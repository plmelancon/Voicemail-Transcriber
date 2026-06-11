# -*- coding: utf-8 -*-
"""
Subprocess-isolated Whisper transcriber.

Runs transcription in a long-lived worker child process so that native
crashes (cuDNN, OpenVINO NPU, ctranslate2) cannot kill the service.
The worker is spawned as ``VoicemailTranscriber.exe --worker --device <d>``
(or ``python voicemail_service.py --worker --device <d>`` in dev mode) and
communicates via one-JSON-line-per-message on stdin/stdout.

Failure handling
----------------
* Worker exits non-zero during a request: the request is retried on CPU.
* Worker exits while idle: silently respawned on next request.
* Repeated crashes on the same EntryID: backoff ladder
      immediate (cpu) -> 30s -> 2min -> 5min -> 15min (forever).
  Never gives up. User requirement: "all voicemails must always be transcribed".
"""

import json
import logging
import os
import subprocess
import sys
import threading
import time
import uuid
from pathlib import Path

logger = logging.getLogger(__name__)


# Backoff ladder (seconds) used after a CPU transcription crash.
# NPU/GPU crashes skip straight to CPU with no delay.
_CPU_BACKOFF_LADDER = [30, 120, 300, 900]  # 30s, 2min, 5min, 15min
_CPU_BACKOFF_FOREVER = 900  # 15min once ladder is exhausted


class _RetryState:
    """Per-EntryID retry bookkeeping. In-memory only — resets on service restart."""

    def __init__(self):
        self._lock = threading.Lock()
        # entry_id -> {"next_attempt": epoch_seconds, "attempts": int, "force_cpu": bool}
        self._state: dict[str, dict] = {}

    def should_attempt(self, entry_id: str) -> bool:
        """Return True if the next retry is due (or if this is the first try)."""
        if not entry_id:
            return True
        with self._lock:
            info = self._state.get(entry_id)
            if info is None:
                return True
            return time.time() >= info["next_attempt"]

    def force_cpu(self, entry_id: str) -> bool:
        """Has a prior NPU/GPU crash forced this entry onto CPU?"""
        if not entry_id:
            return False
        with self._lock:
            info = self._state.get(entry_id)
            return bool(info and info.get("force_cpu"))

    def record_success(self, entry_id: str):
        if not entry_id:
            return
        with self._lock:
            self._state.pop(entry_id, None)

    def record_npu_gpu_fail(self, entry_id: str):
        """Non-CPU backend crashed: next retry should use CPU immediately."""
        if not entry_id:
            return
        with self._lock:
            info = self._state.setdefault(
                entry_id, {"next_attempt": 0.0, "attempts": 0, "force_cpu": False}
            )
            info["force_cpu"] = True
            info["next_attempt"] = 0.0  # retry on next poll

    def record_cpu_fail(self, entry_id: str):
        """CPU transcription crashed: schedule next retry per the backoff ladder."""
        if not entry_id:
            return
        with self._lock:
            info = self._state.setdefault(
                entry_id, {"next_attempt": 0.0, "attempts": 0, "force_cpu": True}
            )
            info["force_cpu"] = True
            idx = min(info["attempts"], len(_CPU_BACKOFF_LADDER) - 1)
            delay = _CPU_BACKOFF_LADDER[idx] if info["attempts"] < len(_CPU_BACKOFF_LADDER) else _CPU_BACKOFF_FOREVER
            info["attempts"] += 1
            info["next_attempt"] = time.time() + delay
            logger.warning(
                "[SubprocessTranscriber] CPU retry #%d for %s scheduled in %ds",
                info["attempts"], entry_id[:16], delay,
            )


class _WorkerProcess:
    """A long-lived worker child, holding the model warm between requests."""

    # How long to wait for a single transcription response before concluding
    # the worker is hung. Large enough to accommodate a cold model load on CPU
    # (faster-whisper large-v3-turbo loads in ~20s, transcribes a 60s clip
    # in ~60s on CPU).
    RESPONSE_TIMEOUT = 600  # 10 minutes

    def __init__(self, device: str, language: str):
        self.device = device
        self.language = language
        self._proc: subprocess.Popen | None = None
        self._lock = threading.Lock()

    def _spawn(self):
        """Launch the worker child and wait for its ready signal."""
        cmd = self._build_cmd()
        logger.info("[SubprocessTranscriber] Spawning worker: device=%s", self.device)
        logger.debug("[SubprocessTranscriber] Command: %s", cmd)

        # CREATE_NO_WINDOW so the child doesn't flash a console
        creationflags = 0
        if sys.platform == "win32":
            creationflags = subprocess.CREATE_NO_WINDOW  # type: ignore[attr-defined]

        self._proc = subprocess.Popen(
            cmd,
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
            errors="replace",
            bufsize=1,  # line-buffered
            creationflags=creationflags,
        )

        # Read the ready handshake line
        ready_line = self._read_line_with_timeout(self.RESPONSE_TIMEOUT)
        if not ready_line:
            self._kill()
            raise RuntimeError("Worker failed to signal ready (no output before timeout)")
        try:
            ready = json.loads(ready_line)
        except json.JSONDecodeError as exc:
            self._kill()
            raise RuntimeError(f"Worker ready signal was not JSON: {ready_line!r} ({exc})")
        if not ready.get("ok") or not ready.get("ready"):
            self._kill()
            raise RuntimeError(f"Worker reported not ready: {ready}")
        logger.info("[SubprocessTranscriber] Worker ready on device=%s", self.device)

    def _build_cmd(self) -> list[str]:
        """Build the command line to re-invoke this program in --worker mode."""
        if getattr(sys, "frozen", False):
            # Frozen exe: re-invoke ourselves with --worker
            exe = sys.executable
            return [exe, "--worker", "--device", self.device, "--no-tray"]
        # Dev mode: re-invoke the script under the same Python
        script = Path(__file__).parent / "voicemail_service.py"
        return [sys.executable, str(script), "--worker", "--device", self.device, "--no-tray"]

    def _read_line_with_timeout(self, timeout: float) -> str:
        """Read one line from stdout with a wall-clock timeout.

        Implemented with a helper thread because Windows has no poll() for
        file handles. Returns the line (without trailing newline) or "" on
        timeout / EOF.
        """
        assert self._proc is not None and self._proc.stdout is not None
        result: list[str] = []
        done = threading.Event()

        def reader():
            try:
                line = self._proc.stdout.readline()  # type: ignore[union-attr]
                result.append(line)
            finally:
                done.set()

        t = threading.Thread(target=reader, name="WorkerReader", daemon=True)
        t.start()
        if not done.wait(timeout):
            return ""  # timeout
        return result[0].rstrip("\n\r") if result and result[0] else ""

    def is_alive(self) -> bool:
        return self._proc is not None and self._proc.poll() is None

    def ensure_alive(self):
        """Spawn or respawn the worker if it isn't currently running."""
        with self._lock:
            if self.is_alive():
                return
            if self._proc is not None:
                # Drain stderr for diagnostic purposes before spawning anew
                try:
                    stderr = self._proc.stderr.read() if self._proc.stderr else ""  # type: ignore[union-attr]
                    if stderr:
                        logger.warning("[SubprocessTranscriber] Previous worker stderr: %s", stderr[:2000])
                except Exception:
                    pass
            self._spawn()

    def transcribe(self, audio_path: str) -> str:
        """Send one transcribe request and wait for the response.

        Raises:
            WorkerCrashedError: the worker died mid-request (native crash).
            WorkerProtocolError: the worker returned an error response.
        """
        with self._lock:
            if not self.is_alive():
                raise WorkerCrashedError("worker not running before request")

            req_id = uuid.uuid4().hex
            req = {"id": req_id, "cmd": "transcribe", "audio_path": audio_path}
            try:
                assert self._proc is not None and self._proc.stdin is not None
                self._proc.stdin.write(json.dumps(req) + "\n")
                self._proc.stdin.flush()
            except (BrokenPipeError, OSError) as exc:
                raise WorkerCrashedError(f"stdin write failed: {exc}")

            line = self._read_line_with_timeout(self.RESPONSE_TIMEOUT)
            if not line:
                # Timeout or EOF — check whether the process is still alive
                if not self.is_alive():
                    rc = self._proc.returncode if self._proc else None
                    raise WorkerCrashedError(f"worker exited with code {rc} during request")
                # Still alive but silent — treat as hung
                self._kill()
                raise WorkerCrashedError("worker hung (no response within timeout)")

            try:
                resp = json.loads(line)
            except json.JSONDecodeError as exc:
                raise WorkerProtocolError(f"bad JSON from worker: {line!r} ({exc})")

            if resp.get("id") != req_id:
                raise WorkerProtocolError(f"id mismatch: sent {req_id}, got {resp.get('id')}")

            if not resp.get("ok"):
                err = resp.get("error", "unknown")
                raise WorkerProtocolError(err)

            return resp.get("text", "")

    def shutdown(self, timeout: float = 5.0):
        """Send a graceful shutdown request, then kill on timeout.

        The worker lock is held for the entire duration of an in-flight
        transcription (up to RESPONSE_TIMEOUT). Waiting for it here is what
        blocked the tray Quit for ~8 minutes on 2026-06-10, so if the lock
        is busy we kill immediately instead of shutting down gracefully.
        """
        if not self._lock.acquire(timeout=1.0):
            logger.warning(
                "[SubprocessTranscriber] Worker busy (request in flight); "
                "killing instead of graceful shutdown"
            )
            self.kill_nowait()
            return
        try:
            if not self.is_alive():
                self._proc = None
                return
            try:
                assert self._proc is not None and self._proc.stdin is not None
                self._proc.stdin.write(json.dumps({"id": "", "cmd": "shutdown"}) + "\n")
                self._proc.stdin.flush()
            except Exception:
                pass
            try:
                if self._proc is not None:
                    self._proc.wait(timeout=timeout)
            except subprocess.TimeoutExpired:
                logger.warning("[SubprocessTranscriber] Worker did not exit in %ds, killing", timeout)
                self._kill()
            self._proc = None
        finally:
            self._lock.release()

    def kill_nowait(self):
        """Kill the worker immediately WITHOUT acquiring the worker lock.

        Safe to call from any thread while a transcription is in flight:
        the death closes the stdout pipe, the in-flight reader gets EOF and
        the blocked transcribe() raises WorkerCrashedError right away.
        ``self._proc`` is left in place so that thread can read the exit
        code; callers discard this _WorkerProcess afterwards.
        """
        proc = self._proc
        if proc is None:
            return
        try:
            proc.kill()
        except Exception:
            pass

    def _kill(self):
        if self._proc is None:
            return
        try:
            self._proc.kill()
        except Exception:
            pass
        try:
            self._proc.wait(timeout=5)
        except Exception:
            pass


class WorkerCrashedError(Exception):
    """The worker process died during a request. The caller should retry."""


class WorkerProtocolError(Exception):
    """The worker returned a well-formed error response (Python exception)."""


class BackoffPendingError(Exception):
    """This EntryID is in the backoff window. Skip it this poll."""


class SubprocessTranscriber:
    """Duck-type-compatible replacement for WhisperModelManager.

    Holds a long-lived worker child process, routes transcriptions through
    it, handles crashes by respawning the worker (on CPU when forced), and
    tracks per-EntryID retry state with backoff.

    Usage from the service is identical to the in-process classes:
        transcriber.transcribe(audio_path)
    except that now an optional entry_id can be passed for retry tracking:
        transcriber.transcribe(audio_path, entry_id="...")
    """

    def __init__(self, device_pref: str = "auto", language: str = "fr"):
        self._user_device_pref = device_pref  # what the user asked for
        self._language = language
        self._retry = _RetryState()
        self._worker: _WorkerProcess | None = None
        self._worker_device: str | None = None  # device the current worker was spawned with
        self._lock = threading.Lock()
        self._closing = False  # set by force_unload: no new workers after this

    # ----- public interface (duck-types WhisperModelManager) -----------------

    def transcribe(self, audio_path: str, entry_id: str = "") -> str:
        """Transcribe audio_path, respecting backoff and falling back to CPU on crash.

        Parameters
        ----------
        audio_path : str
            Path to the audio file.
        entry_id : str
            Outlook EntryID used for retry tracking. If empty, retries are
            still honored but not rate-limited.
        """
        # Backoff check: if this entry has crashed recently, skip
        if not self._retry.should_attempt(entry_id):
            raise BackoffPendingError(f"retry window not elapsed for {entry_id[:16]}")

        target_device = "cpu" if self._retry.force_cpu(entry_id) else self._user_device_pref

        # First attempt on the target device
        try:
            text = self._attempt(target_device, audio_path)
            self._retry.record_success(entry_id)
            return text
        except WorkerCrashedError as exc:
            logger.warning(
                "[SubprocessTranscriber] Worker crashed on %s backend: %s",
                target_device, exc,
            )
            # If the crashing backend was NPU/GPU/auto, fall back to CPU immediately.
            # If it was already CPU, schedule a backoff retry.
            if target_device == "cpu":
                self._retry.record_cpu_fail(entry_id)
                raise
            self._retry.record_npu_gpu_fail(entry_id)
            if self._closing:
                # Quit in progress — don't spawn a CPU retry worker that
                # would outlive the service (observed 2026-06-10 11:12).
                raise
            logger.info("[SubprocessTranscriber] Retrying immediately on CPU")
            try:
                text = self._attempt("cpu", audio_path)
                self._retry.record_success(entry_id)
                return text
            except WorkerCrashedError as cpu_exc:
                logger.error("[SubprocessTranscriber] CPU retry also crashed: %s", cpu_exc)
                self._retry.record_cpu_fail(entry_id)
                raise
        except WorkerProtocolError as exc:
            # Non-crash Python exception in the worker (e.g., malformed audio).
            # Don't treat as fatal for the backend; just propagate.
            logger.error("[SubprocessTranscriber] Worker reported error: %s", exc)
            raise RuntimeError(f"Transcription failed: {exc}") from exc

    def switch_device(self, device_pref: str):
        """Change the user's preferred device. Called from the tray menu."""
        with self._lock:
            logger.info(
                "[SubprocessTranscriber] Device preference: %s -> %s",
                self._user_device_pref, device_pref,
            )
            self._user_device_pref = device_pref
            # Kill current worker so next transcribe spawns with the new device
            if self._worker is not None:
                try:
                    self._worker.shutdown(timeout=3.0)
                except Exception as exc:
                    logger.warning("[SubprocessTranscriber] Shutdown old worker: %s", exc)
                self._worker = None
                self._worker_device = None

    def force_unload(self):
        """Shut down the worker. Called on service quit.

        Must NEVER block: our own lock can be held for up to
        RESPONSE_TIMEOUT by a thread spawning a worker (cold model load),
        and the worker's lock by an in-flight transcription. Kill first,
        lock-free; the in-flight request fails fast with
        WorkerCrashedError and, because _closing is set, no CPU-retry
        worker is respawned mid-shutdown.
        """
        self._closing = True
        worker = self._worker
        if worker is not None:
            worker.kill_nowait()
        if self._lock.acquire(timeout=1.0):
            try:
                self._worker = None
                self._worker_device = None
            finally:
                self._lock.release()

    # ----- internals ---------------------------------------------------------

    def _attempt(self, device: str, audio_path: str) -> str:
        """Ensure a worker is running on the requested device, then transcribe."""
        if self._closing:
            raise WorkerCrashedError("service is shutting down")
        with self._lock:
            if self._worker is None or self._worker_device != device or not self._worker.is_alive():
                # Shut down any existing worker on a different device first
                if self._worker is not None:
                    try:
                        self._worker.shutdown(timeout=2.0)
                    except Exception:
                        pass
                self._worker = _WorkerProcess(device=device, language=self._language)
                self._worker_device = device
                self._worker.ensure_alive()
            worker = self._worker

        # Release the lock during the actual transcription so the tray can
        # still swap devices or force_unload us if needed. The worker's own
        # lock serializes concurrent calls.
        return worker.transcribe(audio_path)
