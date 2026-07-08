# -*- coding: utf-8 -*-
"""
Outlook Monitor for Voicemail Transcriber Service
Watches Inbox via COM events and queues voicemails for transcription.
"""

import logging
import queue
import threading
import time

import pythoncom
import win32api

from resource_manager import set_low_priority, set_normal_priority

logger = logging.getLogger(__name__)

# Sentinel used to signal the processing thread to stop
_STOP = None


class OutlookMonitor:
    """Monitors Outlook Inbox for new voicemails using COM events.

    Threading model
    ---------------
    * **COM event thread** (STA) — connects to Outlook, hooks
      ``Items.ItemAdd``, pumps Windows messages.
    * **Processing thread** — dequeues items and runs transcription one
      at a time so resource usage stays predictable.

    Parameters
    ----------
    processor : OutlookVoicemailProcessor
        The existing processor instance (from voicemail_transcriber.py).
    on_status_change : callable, optional
        ``callback(state: str, detail: str)`` used by the tray app to
        update icon / show notifications.  *state* is one of:
        ``"idle"``, ``"queued"``, ``"transcribing"``, ``"done"``,
        ``"error"``, ``"disconnected"``, ``"reconnecting"``.
    """

    RECONNECT_INTERVAL = 30  # seconds between reconnection attempts
    MAX_RECONNECT_FAILURES = 10  # stop retrying after this many consecutive failures

    WATCHDOG_INTERVAL = 60  # seconds between COM thread health checks

    def __init__(self, processor, on_status_change=None):
        self.processor = processor
        self.on_status_change = on_status_change or (lambda *a: None)
        self._queue: queue.Queue = queue.Queue()
        self._running = False
        self._paused = False
        self._com_thread: threading.Thread | None = None
        self._process_thread: threading.Thread | None = None
        self._com_restart_count = 0

    # -- public API ----------------------------------------------------------

    def start(self):
        """Start both the COM listener and processing threads."""
        self._running = True
        self._com_restart_count = 0
        self._start_com_thread()
        self._process_thread = threading.Thread(
            target=self._processing_loop, name="Transcription", daemon=True
        )
        self._process_thread.start()
        logger.info("Outlook monitor started")

    def _start_com_thread(self):
        """(Re)start the COM event thread."""
        self._com_thread = threading.Thread(
            target=self._com_event_loop, name="OutlookCOM", daemon=True
        )
        self._com_thread.start()

    def stop(self):
        """Signal both threads to shut down gracefully."""
        logger.info("Stopping Outlook monitor...")
        self._running = False
        self._queue.put(_STOP)  # unblock processing thread

    def pause(self):
        self._paused = True
        logger.info("Monitoring paused")

    def resume(self):
        self._paused = False
        logger.info("Monitoring resumed")

    @property
    def is_paused(self):
        return self._paused

    # -- COM event thread ----------------------------------------------------

    def _com_event_loop(self):
        """Runs on a dedicated STA thread.  Hooks Inbox ItemAdd events and
        pumps Windows messages so the callbacks fire."""
        pythoncom.CoInitialize()
        try:
            self._com_event_loop_inner()
        finally:
            pythoncom.CoUninitialize()

    def _com_event_loop_inner(self):
        consecutive_failures = 0

        while self._running:
            try:
                import win32com.client

                from voicemail_transcriber import collect_store_inboxes

                outlook = win32com.client.DispatchWithEvents(
                    "Outlook.Application", _OutlookAppEvents
                )
                namespace = outlook.GetNamespace("MAPI")

                # Hook the Inbox of EVERY store in the profile so voicemails
                # landing in any mailbox are caught. Falls back to the default
                # store's Inbox if store enumeration yields nothing.
                inboxes = collect_store_inboxes(namespace)
                if not inboxes:
                    logger.warning(
                        "Store enumeration yielded nothing; falling back to "
                        "the default store's Inbox"
                    )
                    inboxes = [("(default store)", namespace.GetDefaultFolder(6))]

                # Bind one ItemAdd sink per Inbox. Keep references to BOTH the
                # Items collections and the sinks — the COM connection points
                # only stay alive while these are referenced.
                handler = _make_items_events_class(self)
                hooked = []  # noqa: F841 — held for the pump's lifetime
                for store_name, inbox in inboxes:
                    try:
                        items = inbox.Items
                        sink = win32com.client.WithEvents(items, handler)
                        hooked.append((store_name, items, sink))
                        logger.info("Listening for new emails in: %s", store_name)
                    except pythoncom.com_error as hook_err:
                        logger.warning(
                            "Could not hook Inbox of %s: %s", store_name, hook_err
                        )
                if not hooked:
                    raise RuntimeError("No Inbox could be hooked in any store")

                consecutive_failures = 0
                self.on_status_change("idle", "")
                logger.info(
                    "Connected to Outlook, listening on %d mailbox(es)", len(hooked)
                )

                # Message pump — keeps COM callbacks alive
                last_heartbeat = time.monotonic()
                while self._running:
                    pythoncom.PumpWaitingMessages()
                    win32api.Sleep(100)

                    # Heartbeat every 5 minutes so logs show the thread is alive
                    now = time.monotonic()
                    if now - last_heartbeat >= 300:
                        logger.info("COM event thread heartbeat — still listening")
                        last_heartbeat = now

            except pythoncom.com_error as e:
                consecutive_failures += 1
                logger.warning(
                    f"Outlook COM error (attempt {consecutive_failures}): {e}"
                )
                self.on_status_change("disconnected", str(e))

                if consecutive_failures >= self.MAX_RECONNECT_FAILURES:
                    logger.error(
                        "COM thread exiting: %d consecutive reconnection failures",
                        consecutive_failures,
                    )
                    self.on_status_change("error", "Impossible de se connecter a Outlook")
                    return

                # Wait before retrying
                for _ in range(self.RECONNECT_INTERVAL * 2):
                    if not self._running:
                        return
                    time.sleep(0.5)

                self.on_status_change("reconnecting", "")

            except Exception as e:
                logger.error(f"Unexpected error in COM event loop: {e}", exc_info=True)
                self.on_status_change("error", str(e))
                # Brief pause then retry
                for _ in range(10):
                    if not self._running:
                        return
                    time.sleep(1)

    # -- Processing thread ---------------------------------------------------

    def _processing_loop(self):
        """Dequeues voicemails and transcribes them one at a time."""
        pythoncom.CoInitialize()
        try:
            self._processing_loop_inner()
        finally:
            pythoncom.CoUninitialize()

    def _processing_loop_inner(self):
        import win32com.client

        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")

        while self._running:
            # Use timeout so we periodically run the watchdog even when idle
            try:
                item_info = self._queue.get(timeout=self.WATCHDOG_INTERVAL)
            except queue.Empty:
                self._watchdog_check()
                continue

            if item_info is _STOP:
                break

            entry_id, store_id, subject = item_info

            try:
                logger.info(f"Processing queued email: {subject}")

                # Re-fetch the mail item on THIS thread's COM apartment.
                # "Message not found" (0x8004010F) means it was already
                # processed and its EntryID changed — skip silently.
                try:
                    mail_item = namespace.GetItemFromID(entry_id, store_id)
                except pythoncom.com_error as fetch_err:
                    logger.info(
                        f"Skipping '{subject}': item no longer found (already processed?): {fetch_err}"
                    )
                    self.on_status_change("idle", "")
                    continue

                # Filter here (safe — item is fully delivered by now)
                if not self.processor.is_voicemail(mail_item):
                    logger.info(f"Skipping (not a voicemail or already processed): {subject}")
                    self.on_status_change("idle", "")
                    continue

                self.on_status_change("transcribing", subject)
                logger.info(f"Transcribing voicemail: {subject}")

                set_low_priority()
                try:
                    success = self.processor.process_single_email(mail_item)
                finally:
                    set_normal_priority()

                if success:
                    self.on_status_change("done", subject)
                else:
                    self.on_status_change("error", f"Echec: {subject}")

            except Exception as e:
                set_normal_priority()
                # BackoffPendingError is a normal retry-scheduling signal, not a failure.
                try:
                    from subprocess_transcriber import BackoffPendingError  # noqa: E402
                except ImportError:
                    BackoffPendingError = None  # type: ignore[assignment]
                if BackoffPendingError is not None and isinstance(e, BackoffPendingError):
                    logger.info(f"Deferred (backoff) '{subject}': {e}")
                    self.on_status_change("idle", "")
                    continue
                logger.error(f"Error processing email '{subject}': {e}", exc_info=True)
                self.on_status_change("error", str(e))

            finally:
                # Small pause between items to keep system responsive
                time.sleep(1)
                if self._queue.empty():
                    self.on_status_change("idle", "")

    def _watchdog_check(self):
        """Restart the COM event thread if it has died."""
        if self._com_thread is not None and not self._com_thread.is_alive():
            self._com_restart_count += 1
            logger.warning(
                "COM thread is dead — restarting (restart #%d)",
                self._com_restart_count,
            )
            self.on_status_change("reconnecting", "Redemarrage du moniteur COM")
            self._start_com_thread()
            logger.info("COM thread restarted successfully")


# ---------------------------------------------------------------------------
# COM event handler classes
# ---------------------------------------------------------------------------

class _OutlookAppEvents:
    """Placeholder for Outlook.Application-level events.
    Required by DispatchWithEvents but we don't need any app events."""
    pass


def _make_items_events_class(monitor: OutlookMonitor):
    """Factory that creates an Items event handler class with a closure
    over the monitor instance."""

    class _InboxItemsEvents:
        """Handles Items.ItemAdd fired when a new item arrives in Inbox."""

        def OnItemAdd(self, item):
            try:
                if monitor._paused:
                    return

                # Queue immediately — no filtering here to avoid COM
                # exceptions on partially-delivered items.  The processing
                # thread will call is_voicemail() after re-fetching.
                entry_id = item.EntryID
                store_id = item.Parent.StoreID
                subject = item.Subject or "(Sans sujet)"
                monitor._queue.put((entry_id, store_id, subject))
                monitor.on_status_change("queued", subject)
                logger.info(f"New email queued for evaluation: {subject}")
            except Exception as e:
                logger.error(f"Error in OnItemAdd handler: {e}", exc_info=True)

    return _InboxItemsEvents
