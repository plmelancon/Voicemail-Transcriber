# -*- coding: utf-8 -*-
"""
System Tray Application for Voicemail Transcriber Service
Provides a visible tray icon, status feedback, and user controls.
"""

import ctypes
import logging
import os
import subprocess
import threading
from pathlib import Path

import pystray
from PIL import Image, ImageDraw

logger = logging.getLogger(__name__)

# Directory containing this script
import sys as _sys
_BASE_DIR = Path(_sys.executable).parent if getattr(_sys, 'frozen', False) else Path(__file__).parent
# Logs go to %LOCALAPPDATA% (writable even when app is in Program Files)
_LOG_DIR = Path(os.environ.get('LOCALAPPDATA', str(_BASE_DIR))) / "VoicemailTranscriber" / "logs"

# Icon colours for each state
_COLORS = {
    "idle": "#22c55e",       # green
    "queued": "#f59e0b",     # amber
    "transcribing": "#f97316",  # orange
    "done": "#22c55e",       # green (brief flash handled by timer)
    "paused": "#9ca3af",     # gray
    "error": "#ef4444",      # red
    "disconnected": "#f59e0b",  # amber
    "reconnecting": "#f59e0b",
}

# French labels
_STATUS_LABELS = {
    "idle": "En ecoute",
    "queued": "Message en file",
    "transcribing": "Transcription en cours",
    "done": "Transcription terminee",
    "paused": "En pause",
    "error": "Erreur",
    "disconnected": "Outlook deconnecte",
    "reconnecting": "Reconnexion...",
}


class VoicemailTrayApp:
    """System tray icon with status indicators and a right-click menu.

    Parameters
    ----------
    monitor : OutlookMonitor
        The running monitor instance (for pause/resume/stop).
    processor : OutlookVoicemailProcessor
        The processor (for manual scan action).
    device_pref : str
        Current device preference (``"auto"``, ``"npu"``, ``"cpu"``).
    """

    def __init__(self, monitor, processor, device_pref: str = "auto"):
        self.monitor = monitor
        self.processor = processor
        self._device_pref = device_pref
        self._state = "idle"
        self._detail = ""
        self._icon: pystray.Icon | None = None
        self._revert_timer: threading.Timer | None = None
        # Probe NPU availability once at startup
        from resource_manager import is_npu_available
        self._npu_available = is_npu_available()

    # -- Icon image generation -----------------------------------------------

    @staticmethod
    def _make_icon(color: str) -> Image.Image:
        """Generate a 64x64 circle icon with the given hex colour."""
        img = Image.new("RGBA", (64, 64), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)
        draw.ellipse([4, 4, 60, 60], fill=color)
        # Small inner highlight for visibility
        draw.ellipse([20, 20, 44, 44], fill="#ffffff40")
        return img

    # -- Status callback (called from OutlookMonitor) ------------------------

    def on_status_change(self, state: str, detail: str = ""):
        """Called by the monitor threads to update the tray icon/tooltip."""
        self._state = state
        self._detail = detail

        if self._icon is None:
            return

        color = _COLORS.get(state, _COLORS["idle"])
        label = _STATUS_LABELS.get(state, state)
        tooltip = f"Transcripteur vocal - {label}"
        if detail:
            tooltip += f": {detail}"

        try:
            self._icon.icon = self._make_icon(color)
            self._icon.title = tooltip[:127]  # Windows tooltip limit
        except Exception:
            pass  # icon may not be visible yet

        # Show toast notification for completed transcriptions
        if state == "done" and detail:
            self._show_notification("Transcription terminee", detail)
            # Revert icon to idle after 5 seconds
            self._schedule_revert(5)

        if state == "error" and detail:
            self._show_notification("Erreur de transcription", detail)

    def _schedule_revert(self, seconds: float):
        if self._revert_timer is not None:
            self._revert_timer.cancel()
        self._revert_timer = threading.Timer(
            seconds, self.on_status_change, args=("idle", "")
        )
        self._revert_timer.daemon = True
        self._revert_timer.start()

    # -- Windows toast notification (balloon tip) ----------------------------

    def _show_notification(self, title: str, message: str):
        """Show a balloon-tip notification via pystray."""
        try:
            if self._icon is not None:
                self._icon.notify(message, title)
        except Exception as e:
            logger.debug(f"Could not show notification: {e}")

    # -- Menu actions --------------------------------------------------------

    def _toggle_pause(self, icon, item):
        if self.monitor.is_paused:
            self.monitor.resume()
            self.on_status_change("idle", "")
        else:
            self.monitor.pause()
            self.on_status_change("paused", "")

    def _manual_scan(self, icon, item):
        """Run a manual scan of the last 7 days in a background thread."""
        def _scan():
            import pythoncom
            pythoncom.CoInitialize()
            try:
                self.on_status_change("transcribing", "Traitement manuel...")
                # Reset cached COM connection for this thread
                self.processor.outlook = None
                self.processor.namespace = None
                results = self.processor.process_all_voicemails(days_back=7)
                msg = (
                    f"Traites: {results['processed']}, "
                    f"Echecs: {results['failed']}"
                )
                self.on_status_change("done", msg)
            except Exception as e:
                logger.error(f"Manual scan error: {e}", exc_info=True)
                self.on_status_change("error", str(e))
            finally:
                pythoncom.CoUninitialize()

        threading.Thread(target=_scan, name="ManualScan", daemon=True).start()

    def _switch_device(self, device_pref: str):
        """Switch the transcription backend (auto/npu/cpu)."""
        def _do_switch():
            from resource_manager import save_settings, create_transcriber
            old = self._device_pref
            if device_pref == old:
                return
            self._device_pref = device_pref
            save_settings({"device": device_pref})

            # Unload old transcriber
            try:
                self.processor.transcriber.force_unload()
            except Exception:
                pass

            # Create new one
            language = getattr(self.processor, "config", {}).get("language", "fr")
            self.processor.transcriber = create_transcriber(device_pref, language=language)

            labels = {"auto": "Auto", "npu": "NPU", "cpu": "CPU"}
            self._show_notification(
                "Moteur de transcription",
                f"Changé: {labels.get(old, old)} → {labels.get(device_pref, device_pref)}"
            )
            logger.info(f"Transcription engine switched: {old} → {device_pref}")

        threading.Thread(target=_do_switch, name="DeviceSwitch", daemon=True).start()

    def _open_logs(self, icon, item):
        """Open the log directory in Explorer."""
        _LOG_DIR.mkdir(exist_ok=True)
        os.startfile(str(_LOG_DIR))

    def _quit(self, icon, item):
        """Graceful shutdown."""
        logger.info("User requested quit from tray menu")
        self.monitor.stop()
        icon.stop()

    # -- Menu builder --------------------------------------------------------

    def _build_menu(self):
        def pause_label(item):
            return "Reprendre" if self.monitor.is_paused else "Mettre en pause"

        npu_label = "NPU (Intel AI Boost)" if self._npu_available else "NPU (non disponible)"
        device_submenu = pystray.Menu(
            pystray.MenuItem(
                "Auto (recommandé)",
                lambda icon, item: self._switch_device("auto"),
                checked=lambda item: self._device_pref == "auto",
                radio=True,
            ),
            pystray.MenuItem(
                npu_label,
                lambda icon, item: self._switch_device("npu"),
                checked=lambda item: self._device_pref == "npu",
                radio=True,
                enabled=self._npu_available,
            ),
            pystray.MenuItem(
                "CPU (faster-whisper)",
                lambda icon, item: self._switch_device("cpu"),
                checked=lambda item: self._device_pref == "cpu",
                radio=True,
            ),
        )

        return pystray.Menu(
            pystray.MenuItem(
                lambda item: f"Transcripteur vocal - {_STATUS_LABELS.get(self._state, self._state)}",
                None,
                enabled=False,
            ),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem(pause_label, self._toggle_pause),
            pystray.MenuItem("Traiter manuellement (7 jours)", self._manual_scan),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Moteur de transcription", device_submenu),
            pystray.MenuItem("Ouvrir les logs", self._open_logs),
            pystray.MenuItem("Quitter", self._quit),
        )

    # -- Main entry point ----------------------------------------------------

    def run(self):
        """Create the tray icon and enter the event loop (blocks)."""
        self._icon = pystray.Icon(
            name="voicemail_transcriber",
            icon=self._make_icon(_COLORS["idle"]),
            title="Transcripteur vocal - En ecoute",
            menu=self._build_menu(),
        )
        logger.info("System tray icon starting")
        self._icon.run()  # blocks until icon.stop() is called
