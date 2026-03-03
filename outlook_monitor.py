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

    def __init__(self, processor, on_status_change=None):
        self.processor = processor
        self.on_status_change = on_status_change or (lambda *a: None)
        self._queue: queue.Queue = queue.Queue()
        self._running = False
        self._paused = False
        self._com_thread: threading.Thread | None = None
        self._process_thread: threading.Thread | None = None

    # -- public API ----------------------------------------------------------

    def start(self):
        """Start both the COM listener and processing threads."""
        self._running = True
        self._com_thread = threading.Thread(
            target=self._com_event_loop, name="OutlookCOM", daemon=True
        )
        self._process_thread = threading.Thread(
            target=self._processing_loop, name="Transcription", daemon=True
        )
        self._com_thread.start()
        self._process_thread.start()
        logger.info("Outlook monitor started")

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

                outlook = win32com.client.DispatchWithEvents(
                    "Outlook.Application", _OutlookAppEvents
                )
                namespace = outlook.GetNamespace("MAPI")
                inbox = namespace.GetDefaultFolder(6)
                items = inbox.Items

                # Create the event handler and bind it
                handler = _make_items_events_class(self)
                sink = win32com.client.WithEvents(items, handler)  # noqa: F841

                consecutive_failures = 0
                self.on_status_change("idle", "")
                logger.info("Connected to Outlook, listening for new emails")

                # Message pump — keeps COM callbacks alive
                while self._running:
                    pythoncom.PumpWaitingMessages()
                    win32api.Sleep(500)

            except pythoncom.com_error as e:
                consecutive_failures += 1
                logger.warning(
                    f"Outlook COM error (attempt {consecutive_failures}): {e}"
                )
                self.on_status_change("disconnected", str(e))

                if consecutive_failures >= self.MAX_RECONNECT_FAILURES:
                    logger.error("Too many reconnection failures, giving up")
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
            item_info = self._queue.get()
            if item_info is _STOP:
                break

            entry_id, store_id, subject = item_info

            try:
                self.on_status_change("transcribing", subject)
                logger.info(f"Processing queued voicemail: {subject}")

                # Re-fetch the mail item on THIS thread's COM apartment
                mail_item = namespace.GetItemFromID(entry_id, store_id)

                # Double-check it still qualifies (not already processed)
                if not self.processor.is_voicemail(mail_item):
                    logger.info(f"Skipping (already processed or no longer voicemail): {subject}")
                    self.on_status_change("idle", "")
                    continue

                set_low_priority()
                success = self.processor.process_single_email(mail_item)
                set_normal_priority()

                if success:
                    self.on_status_change("done", subject)
                else:
                    self.on_status_change("error", f"Echec: {subject}")

            except Exception as e:
                set_normal_priority()
                logger.error(f"Error processing voicemail '{subject}': {e}", exc_info=True)
                self.on_status_change("error", str(e))

            finally:
                # Small pause between items to keep system responsive
                time.sleep(1)
                if self._queue.empty():
                    self.on_status_change("idle", "")


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

                if monitor.processor.is_voicemail(item):
                    entry_id = item.EntryID
                    store_id = item.Parent.StoreID
                    subject = item.Subject or "(Sans sujet)"
                    monitor._queue.put((entry_id, store_id, subject))
                    monitor.on_status_change("queued", subject)
                    logger.info(f"Voicemail queued for transcription: {subject}")
            except Exception as e:
                logger.error(f"Error in OnItemAdd handler: {e}", exc_info=True)

    return _InboxItemsEvents
