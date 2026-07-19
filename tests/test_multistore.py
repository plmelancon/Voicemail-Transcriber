# -*- coding: utf-8 -*-
"""Tests for multi-store monitoring (all mailboxes, 2026-07-07).

Run from the repo root:  python -m pytest tests/ -v
"""

import os
import sys
import tempfile
from pathlib import Path
from unittest import mock

# Redirect LOCALAPPDATA BEFORE importing project modules (see
# test_hardening_v0_1_1.py for why).
os.environ.setdefault("LOCALAPPDATA", tempfile.mkdtemp(prefix="vmt_test_"))

REPO_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO_ROOT))

import voicemail_transcriber  # noqa: E402
from voicemail_transcriber import collect_store_inboxes  # noqa: E402


class _FakeStore:
    def __init__(self, name, inbox=None, error=None):
        self.DisplayName = name
        self._inbox = inbox
        self._error = error

    def GetDefaultFolder(self, folder_type):
        assert folder_type == 6
        if self._error:
            raise self._error
        return self._inbox


class _FakeStores:
    def __init__(self, stores):
        self._stores = stores

    @property
    def Count(self):
        return len(self._stores)

    def Item(self, i):  # COM collections are 1-based
        return self._stores[i - 1]


class _FakeNamespace:
    def __init__(self, stores, default_inbox=None):
        self.Stores = _FakeStores(stores)
        self._default_inbox = default_inbox

    def GetDefaultFolder(self, folder_type):
        assert folder_type == 6
        return self._default_inbox


# ---------------------------------------------------------------------------
# collect_store_inboxes
# ---------------------------------------------------------------------------

class TestCollectStoreInboxes:
    def test_all_stores_hooked_in_order(self):
        ns = _FakeNamespace([
            _FakeStore("plmelancon@lacoursiereavocats.com", inbox="INBOX-A"),
            _FakeStore("plm@melanconavocats.com", inbox="INBOX-B"),
        ])
        result = collect_store_inboxes(ns)
        assert result == [
            ("plmelancon@lacoursiereavocats.com", "INBOX-A"),
            ("plm@melanconavocats.com", "INBOX-B"),
        ]

    def test_broken_store_is_skipped_not_fatal(self):
        ns = _FakeNamespace([
            _FakeStore("bonne boite", inbox="INBOX-A"),
            _FakeStore("archive cassee", error=OSError("no inbox")),
            _FakeStore("autre bonne boite", inbox="INBOX-C"),
        ])
        result = collect_store_inboxes(ns)
        assert [name for name, _ in result] == ["bonne boite", "autre bonne boite"]

    def test_stores_enumeration_failure_returns_empty(self):
        class _NoStores:
            @property
            def Stores(self):
                raise OSError("MAPI down")

        assert collect_store_inboxes(_NoStores()) == []


# ---------------------------------------------------------------------------
# get_inbox_folders / find_voicemails
# ---------------------------------------------------------------------------

def _make_processor():
    with mock.patch.object(voicemail_transcriber, "WhisperTranscriber"):
        return voicemail_transcriber.OutlookVoicemailProcessor(
            voicemail_transcriber.CONFIG
        )


class _FakeItems:
    def __init__(self, items):
        self._items = items

    def Restrict(self, restriction):
        assert "[ReceivedTime]" in restriction
        return self

    def Sort(self, key, descending):
        pass

    def __iter__(self):
        return iter(self._items)


class _FakeInbox:
    def __init__(self, items):
        self.Items = _FakeItems(items)


class TestGetInboxFolders:
    def test_falls_back_to_default_store_when_no_stores_found(self):
        proc = _make_processor()
        ns = _FakeNamespace([], default_inbox="DEFAULT-INBOX")
        proc.outlook, proc.namespace = object(), ns
        with mock.patch.object(proc, "connect_outlook", return_value=(None, ns)):
            result = proc.get_inbox_folders()
        assert result == [("(store par defaut)", "DEFAULT-INBOX")]


class TestFindVoicemailsAllStores:
    def test_aggregates_voicemails_across_stores(self):
        proc = _make_processor()
        inboxes = [
            ("boite A", _FakeInbox(["vm1", "courriel", "vm2"])),
            ("boite B", _FakeInbox(["vm3"])),
        ]
        with mock.patch.object(proc, "get_inbox_folders", return_value=inboxes), \
             mock.patch.object(proc, "is_voicemail",
                               side_effect=lambda m: m.startswith("vm")):
            result = proc.find_voicemails(days_back=1)
        assert result == ["vm1", "vm2", "vm3"]

    def test_unreadable_inbox_does_not_block_others(self):
        class _BrokenInbox:
            @property
            def Items(self):
                raise OSError("store offline")

        proc = _make_processor()
        inboxes = [
            ("boite cassee", _BrokenInbox()),
            ("boite saine", _FakeInbox(["vm1"])),
        ]
        with mock.patch.object(proc, "get_inbox_folders", return_value=inboxes), \
             mock.patch.object(proc, "is_voicemail", return_value=True):
            result = proc.find_voicemails(days_back=1)
        assert result == ["vm1"]
