# -*- coding: utf-8 -*-
"""Tests for the v0.1.1 hardening fixes (incident 2026-06-10).

Run from the repo root:  python -m pytest tests/ -v
"""

import os
import subprocess
import sys
import tempfile
import threading
import time
from pathlib import Path
from unittest import mock

import pytest

# Redirect LOCALAPPDATA BEFORE importing project modules: supervisor.py and
# voicemail_service.py configure FileHandlers on the real log files at import
# time, and test runs must not pollute production logs.
os.environ["LOCALAPPDATA"] = tempfile.mkdtemp(prefix="vmt_test_")

REPO_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO_ROOT))

import resource_manager  # noqa: E402
import subprocess_transcriber  # noqa: E402
from subprocess_transcriber import SubprocessTranscriber, _WorkerProcess  # noqa: E402


# The exact error string the NPU produced on 2026-06-10 at 11:02:00.
INCIDENT_ERROR = (
    'Exception from src\\inference\\src\\cpp\\infer_request.cpp:246:\n'
    'Exception from src\\plugins\\intel_npu\\src\\plugin\\npuw\\just_sync_infer_request.cpp:1286:\n'
    'Execution error: "Exception from src\\plugins\\intel_npu\\src\\utils\\src\\zero\\zero_wrappers.cpp:227:\n'
    'L0 zeCommandQueueExecuteCommandLists result: ZE_RESULT_ERROR_DEVICE_LOST, '
    'code 0x70000001 - device hung, reset, was removed, or driver update occurred\n'
    '" - no fallback possible'
)


# ---------------------------------------------------------------------------
# Fix 2 — worker fatal-device fast-exit
# ---------------------------------------------------------------------------

class TestFatalDeviceFastExit:
    def test_incident_error_matches_keywords(self):
        assert any(k in INCIDENT_ERROR for k in resource_manager.FATAL_DEVICE_KEYWORDS)

    def test_benign_error_does_not_match(self):
        benign = "soundfile could not open 'msg0000.MP3': format not recognised"
        assert not any(k in benign for k in resource_manager.FATAL_DEVICE_KEYWORDS)

    def test_noop_outside_worker_mode(self, monkeypatch):
        monkeypatch.delenv(resource_manager.WORKER_FATAL_EXIT_ENV, raising=False)
        # Must return normally even for a fatal error string
        resource_manager.exit_worker_now_if_fatal_device_error(RuntimeError(INCIDENT_ERROR))

    def test_noop_for_benign_error_in_worker_mode(self, monkeypatch):
        monkeypatch.setenv(resource_manager.WORKER_FATAL_EXIT_ENV, "1")
        resource_manager.exit_worker_now_if_fatal_device_error(RuntimeError("bad audio file"))

    def test_exits_3_for_fatal_error_in_worker_mode(self):
        # os._exit can't run inside the pytest process — use a subprocess.
        code = (
            "import os, sys; sys.path.insert(0, r'%s');"
            "import resource_manager as rm;"
            "os.environ[rm.WORKER_FATAL_EXIT_ENV] = '1';"
            "rm.exit_worker_now_if_fatal_device_error(RuntimeError(%r));"
            "sys.exit(99)  # must NOT be reached"
        ) % (str(REPO_ROOT), INCIDENT_ERROR)
        proc = subprocess.run([sys.executable, "-c", code], timeout=60)
        assert proc.returncode == 3

    def test_does_not_exit_without_env_in_subprocess(self):
        code = (
            "import os, sys; sys.path.insert(0, r'%s');"
            "import resource_manager as rm;"
            "os.environ.pop(rm.WORKER_FATAL_EXIT_ENV, None);"
            "rm.exit_worker_now_if_fatal_device_error(RuntimeError(%r));"
            "sys.exit(99)"
        ) % (str(REPO_ROOT), INCIDENT_ERROR)
        proc = subprocess.run([sys.executable, "-c", code], timeout=60)
        assert proc.returncode == 99


# ---------------------------------------------------------------------------
# Fix 3 — non-blocking quit / kill_nowait
# ---------------------------------------------------------------------------

def _sleeper_popen() -> subprocess.Popen:
    """A child process that sleeps, standing in for a busy worker."""
    return subprocess.Popen(
        [sys.executable, "-c", "import time; time.sleep(300)"],
        stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
        text=True,
    )


class TestKillNowait:
    def test_kill_nowait_returns_fast_with_lock_held(self):
        w = _WorkerProcess(device="cpu", language="fr")
        w._proc = _sleeper_popen()
        try:
            assert w.is_alive()
            # Simulate an in-flight transcription holding the worker lock
            w._lock.acquire()
            try:
                start = time.monotonic()
                w.kill_nowait()
                elapsed = time.monotonic() - start
            finally:
                w._lock.release()
            assert elapsed < 1.0, f"kill_nowait blocked for {elapsed:.1f}s"
            w._proc.wait(timeout=10)
            assert not w.is_alive()
        finally:
            if w._proc and w._proc.poll() is None:
                w._proc.kill()

    def test_kill_nowait_noop_without_proc(self):
        w = _WorkerProcess(device="cpu", language="fr")
        w.kill_nowait()  # must not raise

    def test_shutdown_falls_back_to_kill_when_busy(self):
        w = _WorkerProcess(device="cpu", language="fr")
        w._proc = _sleeper_popen()
        try:
            w._lock.acquire()  # lock busy: graceful path unavailable
            try:
                start = time.monotonic()
                w.shutdown(timeout=5.0)
                elapsed = time.monotonic() - start
            finally:
                w._lock.release()
            assert elapsed < 3.0, f"shutdown blocked for {elapsed:.1f}s"
            w._proc.wait(timeout=10)
        finally:
            if w._proc and w._proc.poll() is None:
                w._proc.kill()

    def test_shutdown_graceful_when_idle_and_dead(self):
        w = _WorkerProcess(device="cpu", language="fr")
        w._proc = None
        w.shutdown(timeout=1.0)  # must not raise


class TestForceUnloadNonBlocking:
    def test_force_unload_fast_while_lock_held(self):
        st = SubprocessTranscriber(device_pref="cpu", language="fr")
        w = _WorkerProcess(device="cpu", language="fr")
        w._proc = _sleeper_popen()
        st._worker = w
        st._worker_device = "cpu"
        try:
            # Simulate _attempt holding the transcriber lock (cold spawn)
            st._lock.acquire()
            try:
                start = time.monotonic()
                st.force_unload()
                elapsed = time.monotonic() - start
            finally:
                st._lock.release()
            assert elapsed < 3.0, f"force_unload blocked for {elapsed:.1f}s"
            assert st._closing is True
            w._proc.wait(timeout=10)
            assert w._proc.poll() is not None, "worker not killed"
        finally:
            if w._proc and w._proc.poll() is None:
                w._proc.kill()

    def test_no_new_worker_after_closing(self):
        st = SubprocessTranscriber(device_pref="cpu", language="fr")
        st._closing = True
        with pytest.raises(subprocess_transcriber.WorkerCrashedError):
            st._attempt("cpu", "fake.mp3")

    def test_no_cpu_retry_while_closing(self):
        st = SubprocessTranscriber(device_pref="npu", language="fr")
        st._closing = True
        # _attempt raises "shutting down" → transcribe must re-raise without
        # attempting the CPU respawn path.
        attempts = []
        def fake_attempt(device, path):
            attempts.append(device)
            raise subprocess_transcriber.WorkerCrashedError("service is shutting down")
        st._attempt = fake_attempt
        with pytest.raises(subprocess_transcriber.WorkerCrashedError):
            st.transcribe("fake.mp3", entry_id="abc")
        assert attempts == ["npu"], f"unexpected retry attempts: {attempts}"


# ---------------------------------------------------------------------------
# Fixes 1 & 4 — elevation guard and already-running notification (supervisor)
# ---------------------------------------------------------------------------

import supervisor  # noqa: E402


class TestSupervisorGuards:
    def test_is_elevated_false_in_test_run(self):
        # The test suite runs non-elevated; both detectors must agree.
        assert supervisor._is_elevated() is False

    def test_main_refuses_elevated_with_message(self, monkeypatch):
        shown = []
        monkeypatch.setattr(supervisor, "_is_elevated", lambda: True)
        monkeypatch.setattr(supervisor, "_show_message", lambda text, flags: shown.append((text, flags)))
        rc = supervisor.main()
        assert rc == 3
        assert len(shown) == 1
        assert "administrateur" in shown[0][0]
        assert shown[0][1] & supervisor._MB_ICONWARNING

    def test_main_shows_message_when_already_running(self, monkeypatch):
        shown = []
        monkeypatch.setattr(supervisor, "_is_elevated", lambda: False)
        monkeypatch.setattr(supervisor, "_acquire_single_instance_lock", lambda: False)
        monkeypatch.setattr(supervisor, "_show_message", lambda text, flags: shown.append((text, flags)))
        rc = supervisor.main()
        assert rc == 1
        assert len(shown) == 1
        assert "déjà en cours" in shown[0][0]
        assert shown[0][1] & supervisor._MB_ICONINFORMATION


# ---------------------------------------------------------------------------
# Fix 1 — elevation guard (service side)
# ---------------------------------------------------------------------------

import voicemail_service  # noqa: E402


class TestServiceElevationGuard:
    def test_is_elevated_false_in_test_run(self):
        assert voicemail_service._is_elevated() is False

    def test_refuse_returns_false_when_not_elevated(self, monkeypatch):
        monkeypatch.setattr(voicemail_service, "_is_elevated", lambda: False)
        assert voicemail_service._refuse_elevated_start() is False

    def test_refuse_returns_true_when_elevated(self, monkeypatch):
        monkeypatch.setattr(voicemail_service, "_is_elevated", lambda: True)
        calls = []
        import ctypes
        fake_windll = mock.MagicMock()
        fake_windll.user32.MessageBoxW.side_effect = lambda *a: calls.append(a) or 1
        monkeypatch.setattr(ctypes, "windll", fake_windll)
        assert voicemail_service._refuse_elevated_start() is True
        assert len(calls) == 1
        assert "administrateur" in calls[0][1]
