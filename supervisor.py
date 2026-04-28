# -*- coding: utf-8 -*-
"""
Voicemail Transcriber Supervisor

Tiny parent process that launches VoicemailTranscriber.exe and restarts it
if it dies. This exists because the transcription service can crash natively
inside cuDNN / OpenVINO — exceptions Python can't catch — and when the
service process dies there is nothing to bring it back until the user logs
in again. With this supervisor in charge of autostart (via the HKCU Run key),
a native crash just causes a brief respawn.

Backoff ladder for rapid-crash loops:
    immediate -> 5s -> 30s -> 2min -> 5min (cap)

Clean exits (return code 0) do not trigger a restart — the user quit the
tray icon deliberately. Any non-zero exit triggers restart with backoff.

Scheduled recycle
-----------------
The long-running service process slowly accrues memory (~9 MB/hour observed
in production — likely pywin32 COM proxy buildup from the periodic Outlook
re-Dispatch). To cap that, the supervisor terminates and respawns the child
once it has been alive longer than _MAX_CHILD_AGE_SECONDS, but only during
the quiet-hours window (_QUIET_HOUR_START..._QUIET_HOUR_END, local time) so
the brief tray-icon flicker happens overnight.
"""

from __future__ import annotations

import json
import logging
import os
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------

def _get_app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).parent


APP_DIR = _get_app_dir()
LOG_DIR = Path(os.environ.get("LOCALAPPDATA", str(APP_DIR))) / "VoicemailTranscriber" / "logs"
LOG_DIR.mkdir(parents=True, exist_ok=True)
LOG_FILE = LOG_DIR / f"voicemail_supervisor_{datetime.now().strftime('%Y-%m')}.log"
LOCK_FILE = Path(os.environ.get("LOCALAPPDATA", str(APP_DIR))) / "VoicemailTranscriber" / "supervisor.lock"

# The child to supervise. In a frozen install both exes live in the same
# directory; in dev we fall back to the python script.
def _find_child_cmd() -> list[str]:
    if getattr(sys, "frozen", False):
        # Installed layout: both exes side by side in {app}
        primary = APP_DIR / "VoicemailTranscriber.exe"
        if primary.exists():
            return [str(primary)]
        # PyInstaller dist layout: supervisor is single-file in dist\,
        # service lives in dist\VoicemailTranscriber\VoicemailTranscriber.exe
        nested = APP_DIR / "VoicemailTranscriber" / "VoicemailTranscriber.exe"
        if nested.exists():
            return [str(nested)]
        raise RuntimeError(f"Cannot find VoicemailTranscriber.exe near {APP_DIR}")
    # Dev mode: run the python entry point
    script = Path(__file__).parent / "voicemail_service.py"
    return [sys.executable, str(script)]


# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - [Supervisor] %(message)s",
    handlers=[logging.FileHandler(LOG_FILE, encoding="utf-8")],
)
logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Single-instance lock
# ---------------------------------------------------------------------------

def _acquire_single_instance_lock() -> bool:
    """Return True if we are the only supervisor running.

    Uses a pid file. On Windows, opening a file with exclusive share mode
    would be cleaner but a simple pid file is enough here — two supervisors
    racing is a rare configuration error, not a routine event.
    """
    try:
        LOCK_FILE.parent.mkdir(parents=True, exist_ok=True)
        if LOCK_FILE.exists():
            try:
                old_pid = int(LOCK_FILE.read_text().strip())
                if _pid_is_running(old_pid):
                    logger.warning(
                        f"Another supervisor is already running (pid={old_pid}). Exiting."
                    )
                    return False
            except (ValueError, OSError):
                pass  # stale or unreadable lock — take it over
        LOCK_FILE.write_text(str(os.getpid()))
        return True
    except Exception as exc:
        logger.warning(f"Could not acquire lock file ({exc}); proceeding anyway")
        return True


def _pid_is_running(pid: int) -> bool:
    if pid <= 0:
        return False
    try:
        if sys.platform == "win32":
            import ctypes
            PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
            h = ctypes.windll.kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
            if h:
                ctypes.windll.kernel32.CloseHandle(h)
                return True
            return False
        else:
            os.kill(pid, 0)
            return True
    except Exception:
        return False


def _release_lock():
    try:
        LOCK_FILE.unlink(missing_ok=True)
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Supervision loop
# ---------------------------------------------------------------------------

# Backoff ladder (seconds). Rapid-crash-loop protection.
_BACKOFF_LADDER = [0, 5, 30, 120, 300]
# If the child runs longer than this before exiting, reset the backoff counter.
_HEALTHY_RUN_SECONDS = 120

# Scheduled recycle: terminate + respawn the child once it has been alive
# this long. Caps memory growth from the slow COM-proxy leak in the main
# service process. Picked at 3 days × ~9 MB/h ≈ 700 MB ceiling.
_MAX_CHILD_AGE_SECONDS = 3 * 24 * 3600

# Quiet-hours window (local time, 24h). The recycle waits for this window
# so the brief tray-icon flicker happens overnight, not during a workday.
# Half-open interval [_QUIET_HOUR_START, _QUIET_HOUR_END).
_QUIET_HOUR_START = 2   # 02:00
_QUIET_HOUR_END = 5     # 05:00

# How long to give the child to exit gracefully after terminate() before
# we resort to kill().
_TERMINATE_GRACE_SECONDS = 30

# Poll interval while waiting on the child. Coarse — we only check age and
# alive status, no need for fine granularity.
_POLL_INTERVAL_SECONDS = 60


def _is_quiet_hours(now: datetime | None = None) -> bool:
    """Return True if the local clock is currently inside the quiet window."""
    h = (now or datetime.now()).hour
    return _QUIET_HOUR_START <= h < _QUIET_HOUR_END


def _terminate_child(proc: subprocess.Popen) -> int:
    """Ask the child to exit, then kill if it doesn't. Returns the exit code."""
    try:
        proc.terminate()
    except Exception as exc:
        logger.warning(f"terminate() raised {exc}; falling through to kill")
    try:
        return proc.wait(timeout=_TERMINATE_GRACE_SECONDS)
    except subprocess.TimeoutExpired:
        logger.warning(
            f"Child (pid={proc.pid}) did not exit within {_TERMINATE_GRACE_SECONDS}s; killing"
        )
        try:
            proc.kill()
        except Exception as exc:
            logger.error(f"kill() raised {exc}")
        try:
            return proc.wait(timeout=10)
        except subprocess.TimeoutExpired:
            logger.error(f"Child (pid={proc.pid}) survived kill; abandoning")
            return -1


def _supervise():
    child_cmd = _find_child_cmd()
    logger.info(f"Child command: {child_cmd}")
    logger.info(
        f"Scheduled recycle: every {_MAX_CHILD_AGE_SECONDS//3600}h, "
        f"during quiet hours [{_QUIET_HOUR_START:02d}:00-{_QUIET_HOUR_END:02d}:00)"
    )

    crash_count = 0
    while True:
        started_at = time.monotonic()
        logger.info("Launching child")

        creationflags = 0
        if sys.platform == "win32":
            # DETACHED so the child doesn't inherit our (hidden) console.
            # Child is the tray app — it runs in the user session.
            creationflags = subprocess.CREATE_NO_WINDOW  # type: ignore[attr-defined]

        try:
            proc = subprocess.Popen(
                child_cmd,
                stdin=subprocess.DEVNULL,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                creationflags=creationflags,
            )
        except Exception as exc:
            logger.error(f"Failed to launch child: {exc}")
            _sleep_backoff(crash_count)
            crash_count = min(crash_count + 1, len(_BACKOFF_LADDER) - 1)
            continue

        logger.info(f"Child started (pid={proc.pid})")

        # Poll the child until it exits, or until we decide to recycle it.
        recycled = False
        while True:
            rc = proc.poll()
            if rc is not None:
                break  # child exited on its own
            age = time.monotonic() - started_at
            if age >= _MAX_CHILD_AGE_SECONDS and _is_quiet_hours():
                logger.info(
                    f"Scheduled recycle: child (pid={proc.pid}) reached age "
                    f"{age:.0f}s during quiet hours; terminating"
                )
                rc = _terminate_child(proc)
                recycled = True
                break
            time.sleep(_POLL_INTERVAL_SECONDS)

        run_duration = time.monotonic() - started_at
        logger.info(f"Child exited (pid={proc.pid}) code={rc} after {run_duration:.0f}s")

        if recycled:
            # We killed it on purpose. Restart immediately, do not bump backoff.
            logger.info("Restarting child after scheduled recycle")
            crash_count = 0
            continue

        if rc == 0:
            logger.info("Child exited cleanly (user quit). Supervisor exiting.")
            return

        # Non-zero exit = crash. Decide backoff.
        if run_duration >= _HEALTHY_RUN_SECONDS:
            # It ran long enough to be considered healthy — reset the ladder.
            crash_count = 0
        _sleep_backoff(crash_count)
        crash_count = min(crash_count + 1, len(_BACKOFF_LADDER) - 1)


def _sleep_backoff(crash_count: int):
    delay = _BACKOFF_LADDER[min(crash_count, len(_BACKOFF_LADDER) - 1)]
    if delay <= 0:
        return
    logger.info(f"Backing off {delay}s before restart (crash_count={crash_count})")
    time.sleep(delay)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main() -> int:
    logger.info("=" * 60)
    logger.info("Voicemail Transcriber Supervisor starting")
    logger.info("=" * 60)

    if not _acquire_single_instance_lock():
        return 1

    try:
        _supervise()
    except KeyboardInterrupt:
        logger.info("Supervisor interrupted")
    except Exception as exc:
        logger.error(f"Supervisor error: {exc}", exc_info=True)
        return 2
    finally:
        _release_lock()
        logger.info("Supervisor exiting")
    return 0


if __name__ == "__main__":
    sys.exit(main())
