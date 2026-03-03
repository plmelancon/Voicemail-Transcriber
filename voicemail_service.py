# -*- coding: utf-8 -*-
"""
Voicemail Transcriber Service — Main Entry Point

Runs a system tray application that monitors Outlook for incoming voicemails
and transcribes them automatically using faster-whisper.

Usage:
    python voicemail_service.py              # Normal mode (system tray)
    python voicemail_service.py --no-tray    # Headless mode (console only, for debugging)
"""

import argparse
import logging
import os
import sys
import threading
from datetime import datetime
from pathlib import Path

# ---------------------------------------------------------------------------
# Logging setup (must happen before any other import that uses logging)
# ---------------------------------------------------------------------------

def _get_app_dir():
    """Return the application directory (works both dev and frozen/PyInstaller)."""
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).parent
    return Path(__file__).parent

APP_DIR = _get_app_dir()
# Logs go to %LOCALAPPDATA% (writable even when app is in Program Files)
LOG_DIR = Path(os.environ.get('LOCALAPPDATA', str(APP_DIR))) / "VoicemailTranscriber" / "logs"
LOG_DIR.mkdir(parents=True, exist_ok=True)
LOG_FILE = LOG_DIR / f"voicemail_service_{datetime.now().strftime('%Y-%m')}.log"

_handlers = [logging.FileHandler(LOG_FILE, encoding="utf-8")]
if sys.stdout is not None:
    _handlers.append(logging.StreamHandler(sys.stdout))

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - [%(threadName)s] %(message)s",
    handlers=_handlers,
)
logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Imports from project modules
# ---------------------------------------------------------------------------

from voicemail_transcriber import CONFIG, OutlookVoicemailProcessor  # noqa: E402
from resource_manager import WhisperModelManager  # noqa: E402
from outlook_monitor import OutlookMonitor  # noqa: E402
from tray_app import VoicemailTrayApp  # noqa: E402


def main():
    parser = argparse.ArgumentParser(
        description="Service de transcription automatique des messages vocaux"
    )
    parser.add_argument(
        "--no-tray",
        action="store_true",
        help="Mode sans icone de barre des taches (console uniquement)",
    )
    args = parser.parse_args()

    logger.info("=" * 60)
    logger.info("  Voicemail Transcriber Service starting")
    logger.info("=" * 60)

    # 1. Build the processor reusing existing CONFIG
    processor = OutlookVoicemailProcessor(CONFIG)

    # 2. Replace the transcriber with our managed faster-whisper version
    processor.transcriber = WhisperModelManager(
        model_name=CONFIG["whisper_model"],
        language=CONFIG["language"],
    )

    # 3. Create the tray app and monitor (wired together via callback)
    if args.no_tray:
        # Headless mode: just log status changes
        def _log_status(state, detail=""):
            logger.info(f"[STATUS] {state}: {detail}")

        monitor = OutlookMonitor(processor, on_status_change=_log_status)
    else:
        # Normal mode: tray app provides the status callback
        # We create the tray app first but wire the monitor's callback to it
        monitor = OutlookMonitor(processor)
        tray = VoicemailTrayApp(monitor, processor)
        monitor.on_status_change = tray.on_status_change

    # 4. Catch-up scan in a background thread (process voicemails that
    #    arrived while the service was not running)
    def _catchup():
        import pythoncom
        pythoncom.CoInitialize()
        try:
            logger.info("Running catch-up scan (last 24 hours)...")
            # Reset cached COM connection so this thread gets its own
            processor.outlook = None
            processor.namespace = None
            results = processor.process_all_voicemails(days_back=1)
            logger.info(f"Catch-up scan complete: {results}")
        except Exception as e:
            logger.error(f"Catch-up scan failed: {e}", exc_info=True)
        finally:
            pythoncom.CoUninitialize()

    catchup_thread = threading.Thread(target=_catchup, name="CatchupScan", daemon=True)
    catchup_thread.start()

    # 5. Start monitoring for new emails
    monitor.start()

    # 6. Enter the main loop
    if args.no_tray:
        logger.info("Running in headless mode. Press Ctrl+C to quit.")
        try:
            # Block the main thread until interrupted
            while True:
                threading.Event().wait(timeout=1)
        except KeyboardInterrupt:
            logger.info("Keyboard interrupt received, shutting down...")
            monitor.stop()
    else:
        logger.info("System tray icon active. Right-click to access menu.")
        tray.run()  # blocks until user clicks Quit

    # 7. Cleanup
    if hasattr(processor.transcriber, "force_unload"):
        processor.transcriber.force_unload()
    logger.info("Service stopped")


if __name__ == "__main__":
    main()
