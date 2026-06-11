# -*- coding: utf-8 -*-
"""
Voicemail Transcriber Service — Main Entry Point

Runs a system tray application that monitors Outlook for incoming voicemails
and transcribes them automatically using faster-whisper (CPU) or OpenVINO (NPU).

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
# CRITICAL: Subprocess modes need stdio
# ---------------------------------------------------------------------------
# PyInstaller's windowed (--noconsole) builds set sys.stdout and sys.stderr
# to None unconditionally, even when the parent process piped them via
# subprocess.Popen(stdout=PIPE). But the underlying OS file descriptors
# (fd 0/1/2) are still valid and point at the pipes. For --worker and
# --probe-devices modes we NEED stdio to work, so reopen fd 0/1/2 as Python
# file objects. This must happen before logging.basicConfig touches stdout.
_NEEDS_STDIO = any(a in sys.argv for a in ("--worker", "--probe-devices"))
if _NEEDS_STDIO:
    if sys.stdout is None:
        try:
            sys.stdout = os.fdopen(1, "w", buffering=1, encoding="utf-8", errors="replace")
        except Exception:
            pass
    if sys.stderr is None:
        try:
            sys.stderr = os.fdopen(2, "w", buffering=1, encoding="utf-8", errors="replace")
        except Exception:
            pass
    if sys.stdin is None:
        try:
            sys.stdin = os.fdopen(0, "r", encoding="utf-8", errors="replace")
        except Exception:
            pass

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

# Pick a log filename based on how this process was invoked. Subprocess
# modes (--worker, --probe-devices) must NOT write to the same file as the
# long-running service, because Python's FileHandler has no cross-process
# locking and concurrent writes would interleave. Detect mode from argv
# here, before logging.basicConfig is called.
_is_worker = "--worker" in sys.argv
_is_probe = "--probe-devices" in sys.argv

if _is_probe:
    # Probe is short-lived and we need stdout to carry a clean JSON line.
    # Keep logging quiet and out of the log directory entirely.
    _handlers = [logging.StreamHandler(sys.stderr)] if sys.stderr is not None else []
    _log_level = logging.WARNING
elif _is_worker:
    _log_file = LOG_DIR / f"voicemail_worker_{datetime.now().strftime('%Y-%m')}.log"
    _handlers = [logging.FileHandler(_log_file, encoding="utf-8")]
    # Do NOT add a StreamHandler on stdout — stdout is the IPC channel.
    if sys.stderr is not None:
        _handlers.append(logging.StreamHandler(sys.stderr))
    _log_level = logging.INFO
else:
    LOG_FILE = LOG_DIR / f"voicemail_service_{datetime.now().strftime('%Y-%m')}.log"
    _handlers = [logging.FileHandler(LOG_FILE, encoding="utf-8")]
    if sys.stdout is not None:
        _handlers.append(logging.StreamHandler(sys.stdout))
    _log_level = logging.INFO

logging.basicConfig(
    level=_log_level,
    format="%(asctime)s - %(levelname)s - [%(threadName)s] %(message)s",
    handlers=_handlers,
)
logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Imports from project modules
# ---------------------------------------------------------------------------

from voicemail_transcriber import CONFIG, OutlookVoicemailProcessor  # noqa: E402


def _run_worker_mode(device: str, language: str) -> int:
    """Long-lived transcription worker. Reads JSON requests from stdin,
    writes JSON responses to stdout, one per line.

    Protocol:
        request  : {"id": str, "cmd": "transcribe", "audio_path": str}
                 | {"id": str, "cmd": "shutdown"}
        response : {"id": str, "ok": true,  "text": str}
                 | {"id": str, "ok": false, "error": str}

    Exits 0 on clean shutdown, non-zero on fatal error. Native crashes
    (cuDNN etc.) propagate as the process exit code, which the parent
    SubprocessTranscriber detects and treats as "this backend is dead".
    """
    import json
    from resource_manager import WORKER_FATAL_EXIT_ENV, create_transcriber  # noqa: E402

    # Arm the fatal-device fast-exit BEFORE the backend is created: on a
    # DEVICE_LOST-class error the backend calls os._exit(3) instead of
    # unwinding (which froze the worker on 2026-06-10). The parent treats
    # the death as WorkerCrashedError and retries on CPU within seconds.
    os.environ[WORKER_FATAL_EXIT_ENV] = "1"

    logger.info(f"[Worker] Starting worker mode on device={device}")
    try:
        transcriber = create_transcriber(device, language=language)
    except Exception as exc:
        logger.error(f"[Worker] Failed to create transcriber: {exc}", exc_info=True)
        sys.stdout.write(json.dumps({"id": "", "ok": False, "error": f"init_failed: {exc}"}) + "\n")
        sys.stdout.flush()
        return 2

    sys.stdout.write(json.dumps({"id": "", "ok": True, "ready": True}) + "\n")
    sys.stdout.flush()
    logger.info("[Worker] Ready")

    for line in sys.stdin:
        line = line.strip()
        if not line:
            continue
        try:
            req = json.loads(line)
        except json.JSONDecodeError as exc:
            sys.stdout.write(json.dumps({"id": "", "ok": False, "error": f"bad_json: {exc}"}) + "\n")
            sys.stdout.flush()
            continue

        req_id = req.get("id", "")
        cmd = req.get("cmd")

        if cmd == "shutdown":
            logger.info("[Worker] Shutdown requested")
            sys.stdout.write(json.dumps({"id": req_id, "ok": True, "bye": True}) + "\n")
            sys.stdout.flush()
            try:
                if hasattr(transcriber, "force_unload"):
                    transcriber.force_unload()
            except Exception:
                pass
            return 0

        if cmd == "transcribe":
            audio_path = req.get("audio_path", "")
            try:
                text = transcriber.transcribe(audio_path)
                sys.stdout.write(json.dumps({"id": req_id, "ok": True, "text": text}) + "\n")
                sys.stdout.flush()
            except Exception as exc:
                logger.error(f"[Worker] Transcription error: {exc}", exc_info=True)
                sys.stdout.write(json.dumps({"id": req_id, "ok": False, "error": str(exc)}) + "\n")
                sys.stdout.flush()
            continue

        sys.stdout.write(json.dumps({"id": req_id, "ok": False, "error": f"unknown_cmd: {cmd}"}) + "\n")
        sys.stdout.flush()

    logger.info("[Worker] stdin closed, exiting")
    return 0


def _run_probe_devices() -> int:
    """One-shot device availability probe. Prints JSON to stdout and exits."""
    import json
    from resource_manager import is_npu_available, is_gpu_available  # noqa: E402
    result = {"npu": bool(is_npu_available()), "gpu": bool(is_gpu_available())}
    sys.stdout.write(json.dumps(result) + "\n")
    sys.stdout.flush()
    return 0


from resource_manager import load_settings  # noqa: E402
from outlook_monitor import OutlookMonitor  # noqa: E402
from tray_app import VoicemailTrayApp  # noqa: E402
from subprocess_transcriber import SubprocessTranscriber  # noqa: E402


def _is_elevated() -> bool:
    """True when this process runs with an elevated (administrator) token.

    An elevated process cannot connect to the user's non-elevated Outlook:
    every COM Dispatch fails with 'Server execution failed' (-2146959355),
    so the service would run as a zombie that never sees a voicemail
    (incident 2026-06-10, 2.5 hours of silent non-operation).
    """
    if sys.platform != "win32":
        return False
    try:
        import ctypes
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def _refuse_elevated_start() -> bool:
    """If elevated, log + show a warning and return True (caller must exit)."""
    if not _is_elevated():
        return False
    logger.error(
        "Refusing to start: process is elevated (run as administrator). "
        "Outlook COM is unreachable from an elevated process; no voicemail "
        "would ever be transcribed. Relaunch normally."
    )
    try:
        import ctypes
        ctypes.windll.user32.MessageBoxW(
            None,
            "Le Transcripteur de messages vocaux ne doit pas être exécuté "
            "en tant qu'administrateur.\n\n"
            "Un processus élevé ne peut pas communiquer avec Outlook : "
            "aucun message vocal ne serait traité.\n\n"
            "Veuillez relancer l'application normalement, sans "
            "« Exécuter en tant qu'administrateur ».",
            "Transcripteur vocal",
            0x00000030 | 0x00040000 | 0x00010000,  # MB_ICONWARNING | MB_TOPMOST | MB_SETFOREGROUND
        )
    except Exception:
        pass
    return True


def main():
    parser = argparse.ArgumentParser(
        description="Service de transcription automatique des messages vocaux"
    )
    parser.add_argument(
        "--no-tray",
        action="store_true",
        help="Mode sans icone de barre des taches (console uniquement)",
    )
    parser.add_argument(
        "--worker",
        action="store_true",
        help="Run as a transcription worker subprocess (internal use)",
    )
    parser.add_argument(
        "--probe-devices",
        action="store_true",
        help="Print NPU/GPU availability as JSON and exit (internal use)",
    )
    parser.add_argument(
        "--device",
        type=str,
        default="auto",
        choices=["auto", "npu", "gpu", "cpu"],
        help="Device for --worker mode",
    )
    args = parser.parse_args()

    # Subprocess modes dispatch before any tray/monitor setup
    if args.probe_devices:
        sys.exit(_run_probe_devices())
    if args.worker:
        sys.exit(_run_worker_mode(args.device, CONFIG["language"]))

    logger.info("=" * 60)
    logger.info("  Voicemail Transcriber Service starting")
    logger.info("=" * 60)

    # Exit code 0 on refusal: if a supervisor spawned us, a clean exit means
    # "user quit" and prevents a refuse/restart loop.
    if _refuse_elevated_start():
        return

    # 1. Build the processor reusing existing CONFIG
    processor = OutlookVoicemailProcessor(CONFIG)

    # 2. Load settings and attach a subprocess-isolated transcriber.
    #    Inference runs in a long-lived child process so native crashes
    #    (e.g. cuDNN access violations) cannot kill the service.
    settings = load_settings()
    device_pref = settings.get("device", "auto")
    logger.info(f"Device preference from settings: {device_pref}")
    processor.transcriber = SubprocessTranscriber(
        device_pref=device_pref,
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
        tray = VoicemailTrayApp(monitor, processor, device_pref=device_pref)
        monitor.on_status_change = tray.on_status_change

    # 4. Catch-up scan in a background thread (process voicemails that
    #    arrived while the service was not running)
    def _catchup():
        import pythoncom
        pythoncom.CoInitialize()
        try:
            logger.info("Running catch-up scan (last 7 days)...")
            # Reset cached COM connection so this thread gets its own
            processor.outlook = None
            processor.namespace = None
            results = processor.process_all_voicemails(days_back=7)
            logger.info(f"Catch-up scan complete: {results}")
        except Exception as e:
            logger.error(f"Catch-up scan failed: {e}", exc_info=True)
        finally:
            pythoncom.CoUninitialize()

    catchup_thread = threading.Thread(target=_catchup, name="CatchupScan", daemon=True)
    catchup_thread.start()

    # 5. Periodic polling safety net — catches anything the live monitor misses
    POLL_INTERVAL_MINUTES = 1

    def _periodic_poll():
        import pythoncom
        pythoncom.CoInitialize()
        try:
            while True:
                threading.Event().wait(timeout=POLL_INTERVAL_MINUTES * 60)
                try:
                    logger.info("Periodic poll: scanning last 24 hours...")
                    # Fresh COM connection for this thread
                    processor.outlook = None
                    processor.namespace = None
                    results = processor.process_all_voicemails(days_back=1)
                    logger.info(f"Periodic poll complete: {results}")
                except Exception as e:
                    logger.error(f"Periodic poll failed: {e}", exc_info=True)
        finally:
            pythoncom.CoUninitialize()

    poll_thread = threading.Thread(target=_periodic_poll, name="PeriodicPoll", daemon=True)
    poll_thread.start()

    # 6. Start monitoring for new emails
    monitor.start()

    # 7. Enter the main loop
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

    # 8. Cleanup
    if hasattr(processor.transcriber, "force_unload"):
        processor.transcriber.force_unload()
    logger.info("Service stopped")


if __name__ == "__main__":
    main()
