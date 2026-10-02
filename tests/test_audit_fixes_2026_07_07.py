# -*- coding: utf-8 -*-
"""Regression tests for the 2026-07-07 audit fixes.

Covers:
  - C1: worker stderr is continuously drained (a chatty worker can no longer
        deadlock on a full stderr pipe) and the post-mortem tail is kept.
  - C2: in-flight EntryID guard — two intake paths can't transcribe the same
        voicemail concurrently.
  - B1: single-instance named mutex is race-free across simultaneous
        process launches and leaves no stale state after death.
  - B2: scheduled recycle fires unconditionally at 2x max age (machines
        asleep during quiet hours).
"""
import json
import os
import subprocess
import sys
import tempfile
import threading
import time
from datetime import datetime
from pathlib import Path
from unittest import mock

import pytest

# Redirect LOCALAPPDATA BEFORE importing project modules (import-time
# FileHandlers must not pollute production logs).
os.environ["LOCALAPPDATA"] = tempfile.mkdtemp(prefix="vmt_test_")

REPO_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO_ROOT))

import subprocess_transcriber  # noqa: E402
import supervisor  # noqa: E402
import voicemail_transcriber  # noqa: E402
from subprocess_transcriber import _WorkerProcess  # noqa: E402


# ---------------------------------------------------------------------------
# C1 — stderr drain
# ---------------------------------------------------------------------------

# A fake worker: prints the ready handshake, then for every request floods
# stderr with far more than any OS pipe buffer before answering. Without a
# drainer thread in the parent, the flood blocks the worker before it can
# respond and the request times out.
_FAKE_WORKER = r"""
import json, sys
print(json.dumps({"ok": True, "ready": True}), flush=True)
for line in sys.stdin:
    req = json.loads(line)
    if req.get("cmd") == "shutdown":
        break
    for i in range(4000):  # ~256 KB of stderr per request
        sys.stderr.write("noise %d: 0123456789012345678901234567890123456789012345678901234567\n" % i)
    sys.stderr.flush()
    print(json.dumps({"id": req["id"], "ok": True, "text": "ok"}), flush=True)
"""


class TestStderrDrain:
    def _make_worker(self):
        wp = _WorkerProcess(device="cpu", language="fr")
        wp._build_cmd = lambda: [sys.executable, "-u", "-c", _FAKE_WORKER]
        return wp

    def test_chatty_worker_does_not_deadlock(self):
        wp = self._make_worker()
        with mock.patch.object(_WorkerProcess, "RESPONSE_TIMEOUT", 30):
            wp.ensure_alive()
            try:
                start = time.monotonic()
                for _ in range(3):
                    assert wp.transcribe("dummy.wav") == "ok"
                elapsed = time.monotonic() - start
                # Undrained, each request would stall RESPONSE_TIMEOUT and fail.
                assert elapsed < 25, f"requests too slow ({elapsed:.1f}s) — pipe backpressure?"
                assert len(wp._stderr_tail) > 0, "drainer captured nothing"
            finally:
                wp.shutdown(timeout=5)

    def test_post_mortem_tail_logged_on_respawn(self, caplog):
        wp = self._make_worker()
        with mock.patch.object(_WorkerProcess, "RESPONSE_TIMEOUT", 30):
            wp.ensure_alive()
            wp.transcribe("dummy.wav")
            wp.kill_nowait()
            wp._proc.wait(timeout=10)
            time.sleep(0.5)  # let the drainer hit EOF
            with caplog.at_level("WARNING"):
                wp.ensure_alive()  # respawn logs the previous worker's tail
            wp.shutdown(timeout=5)
        assert any("Previous worker stderr (tail)" in r.message for r in caplog.records)


# ---------------------------------------------------------------------------
# C2 — in-flight EntryID guard
# ---------------------------------------------------------------------------

class TestInFlightGuard:
    def _bare_processor(self):
        proc = object.__new__(voicemail_transcriber.OutlookVoicemailProcessor)
        proc._in_flight = set()
        proc._in_flight_lock = threading.Lock()
        return proc

    def _mail(self, entry_id="AABB01", subject="Voicemail from 4181234567"):
        m = mock.Mock()
        m.EntryID = entry_id
        m.Subject = subject
        return m

    def test_second_concurrent_call_is_rejected(self):
        proc = self._bare_processor()
        started = threading.Event()
        release = threading.Event()
        results = {}

        def slow_guarded(mail_item, subject, entry_id):
            started.set()
            release.wait(10)
            return True

        proc._process_single_email_guarded = slow_guarded

        t = threading.Thread(
            target=lambda: results.setdefault("first", proc.process_single_email(self._mail()))
        )
        t.start()
        assert started.wait(5)
        # Same EntryID while the first is in flight -> immediate False
        assert proc.process_single_email(self._mail()) is False
        release.set()
        t.join(10)
        assert results["first"] is True

    def test_entry_id_released_after_processing(self):
        proc = self._bare_processor()
        proc._process_single_email_guarded = lambda *a: True
        assert proc.process_single_email(self._mail()) is True
        assert proc.process_single_email(self._mail()) is True  # not stuck in-flight

    def test_entry_id_released_after_exception(self):
        proc = self._bare_processor()

        def boom(*a):
            raise RuntimeError("kaput")

        proc._process_single_email_guarded = boom
        with pytest.raises(RuntimeError):
            proc.process_single_email(self._mail())
        assert not proc._in_flight, "EntryID leaked into the in-flight set"

    def test_different_entry_ids_do_not_block_each_other(self):
        proc = self._bare_processor()
        proc._in_flight.add("OTHER")
        proc._process_single_email_guarded = lambda *a: True
        assert proc.process_single_email(self._mail(entry_id="AABB01")) is True


# ---------------------------------------------------------------------------
# B1 — named mutex single-instance lock
# ---------------------------------------------------------------------------

_MUTEX_PROBE = r"""
import os, sys, time
os.environ["LOCALAPPDATA"] = sys.argv[1]
sys.path.insert(0, sys.argv[2])
import supervisor
got = supervisor._acquire_single_instance_lock()
print("ACQUIRED=%s" % got, flush=True)
if got:
    time.sleep(3)
    supervisor._release_lock()
"""


@pytest.mark.skipif(sys.platform != "win32", reason="Windows named mutex")
class TestMutexRace:
    def test_exactly_one_winner_then_reacquirable(self, tmp_path):
        args = [str(tmp_path), str(REPO_ROOT)]
        procs = [
            subprocess.Popen(
                [sys.executable, "-c", _MUTEX_PROBE] + args,
                stdout=subprocess.PIPE, stderr=subprocess.DEVNULL, text=True,
            )
            for _ in range(5)
        ]
        outs = [p.communicate(timeout=60)[0] for p in procs]
        winners = sum("ACQUIRED=True" in o for o in outs)
        assert winners == 1, f"expected exactly 1 winner, got {winners}: {outs}"
        # All dead -> the OS released the mutex -> a fresh instance acquires.
        p = subprocess.Popen(
            [sys.executable, "-c", _MUTEX_PROBE] + args,
            stdout=subprocess.PIPE, stderr=subprocess.DEVNULL, text=True,
        )
        out = p.communicate(timeout=60)[0]
        assert "ACQUIRED=True" in out


# ---------------------------------------------------------------------------
# B2 — forced recycle at 2x max age
# ---------------------------------------------------------------------------

class TestForcedRecycle:
    def test_below_max_never_recycles(self):
        assert not supervisor._should_recycle(
            supervisor._MAX_CHILD_AGE_SECONDS - 1, datetime(2026, 7, 7, 3, 0)
        ) or supervisor._MAX_CHILD_AGE_SECONDS <= 1

    def test_past_max_only_during_quiet_hours(self):
        age = supervisor._MAX_CHILD_AGE_SECONDS + 60
        assert supervisor._should_recycle(age, datetime(2026, 7, 7, 3, 0))      # 03:00 quiet
        assert not supervisor._should_recycle(age, datetime(2026, 7, 7, 14, 0))  # 14:00 not

    def test_past_double_max_recycles_any_time_of_day(self):
        age = supervisor._FORCE_RECYCLE_AGE_SECONDS + 1
        for hour in (0, 3, 9, 14, 23):
            assert supervisor._should_recycle(age, datetime(2026, 7, 7, hour, 0))

    def test_force_threshold_is_double_max(self):
        assert supervisor._FORCE_RECYCLE_AGE_SECONDS == 2 * supervisor._MAX_CHILD_AGE_SECONDS
