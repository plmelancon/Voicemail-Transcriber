# -*- coding: utf-8 -*-
"""
Resource Manager for Voicemail Transcriber Service
Handles process priority and Whisper model lifecycle (faster-whisper).
"""

import gc
import logging
import sys as _sys
import threading
import time
from pathlib import Path

import win32api
import win32process

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Process priority helpers
# ---------------------------------------------------------------------------

def set_low_priority():
    """Set current process to IDLE priority so user apps always get CPU first."""
    handle = win32api.GetCurrentProcess()
    win32process.SetPriorityClass(handle, win32process.IDLE_PRIORITY_CLASS)
    logger.debug("Process priority set to IDLE")


def set_normal_priority():
    """Restore normal process priority."""
    handle = win32api.GetCurrentProcess()
    win32process.SetPriorityClass(handle, win32process.NORMAL_PRIORITY_CLASS)
    logger.debug("Process priority set to NORMAL")


# ---------------------------------------------------------------------------
# Whisper model manager (faster-whisper, with idle auto-unload)
# ---------------------------------------------------------------------------

class WhisperModelManager:
    """Drop-in replacement for WhisperTranscriber using faster-whisper.

    Exposes the same ``transcribe(audio_path) -> str`` interface so it can be
    assigned directly to ``processor.transcriber``.

    Features:
        - Lazy model loading on first transcription
        - Automatic unload after IDLE_TIMEOUT seconds of inactivity
        - Thread-safe (guarded by a Lock)
    """

    IDLE_TIMEOUT = 300  # 5 minutes

    def __init__(self, model_name: str = "turbo", language: str = "fr"):
        self.model_name = model_name
        self.language = language
        self.device = "cpu"
        self.compute_type = "float32"
        self._model = None
        self._lock = threading.Lock()
        self._last_used: float = 0
        self._unload_timer: threading.Timer | None = None

    # -- public interface (duck-type compatible with WhisperTranscriber) ------

    def load_model(self):
        """Load the faster-whisper model (lazy). Returns the model object."""
        with self._lock:
            if self._model is None:
                logger.info(f"Loading faster-whisper model '{self.model_name}' on {self.device}...")
                from faster_whisper import WhisperModel
                import sys as _sys
                model_source = self.model_name
                if getattr(_sys, 'frozen', False):
                    bundled = Path(_sys.executable).parent / "models" / "faster-whisper-large-v3-turbo"
                    if bundled.exists():
                        model_source = str(bundled)
                        logger.info(f"Using bundled model at {model_source}")
                self._model = WhisperModel(
                    model_source,
                    device=self.device,
                    compute_type=self.compute_type,
                )
                logger.info("Model loaded successfully")
            self._last_used = time.time()
            self._reset_unload_timer()
            return self._model

    def transcribe(self, audio_path: str) -> str:
        """Transcribe an audio file and return the text.

        The entire transcription runs under the lock because ctranslate2 is
        NOT thread-safe for concurrent inference — two threads calling
        model.transcribe() simultaneously causes a native crash.
        """
        with self._lock:
            if self._model is None:
                logger.info(f"Loading faster-whisper model '{self.model_name}' on {self.device}...")
                from faster_whisper import WhisperModel
                model_source = self.model_name
                if getattr(_sys, 'frozen', False):
                    bundled = Path(_sys.executable).parent / "models" / "faster-whisper-large-v3-turbo"
                    if bundled.exists():
                        model_source = str(bundled)
                        logger.info(f"Using bundled model at {model_source}")
                self._model = WhisperModel(
                    model_source,
                    device=self.device,
                    compute_type=self.compute_type,
                )
                logger.info("Model loaded successfully")
            self._last_used = time.time()
            self._reset_unload_timer()

            logger.info(f"Transcribing: {audio_path}")
            try:
                segments, _info = self._model.transcribe(
                    audio_path,
                    language=self.language,
                    task="transcribe",
                )
                text = " ".join(seg.text.strip() for seg in segments).strip()
                logger.info(f"Transcription complete: {len(text)} characters")
                return text
            except Exception as e:
                logger.error(f"Transcription error: {e}")
                raise

    def force_unload(self):
        """Immediately unload the model and free resources."""
        with self._lock:
            self._cancel_timer()
            self._do_unload()

    # -- internal ------------------------------------------------------------

    def _reset_unload_timer(self):
        self._cancel_timer()
        self._unload_timer = threading.Timer(self.IDLE_TIMEOUT, self._unload_if_idle)
        self._unload_timer.daemon = True
        self._unload_timer.start()

    def _cancel_timer(self):
        if self._unload_timer is not None:
            self._unload_timer.cancel()
            self._unload_timer = None

    def _unload_if_idle(self):
        with self._lock:
            if self._model is not None and (time.time() - self._last_used >= self.IDLE_TIMEOUT):
                self._do_unload()

    def _do_unload(self):
        if self._model is not None:
            logger.info("Unloading Whisper model (idle timeout)...")
            del self._model
            self._model = None
            gc.collect()
            try:
                import torch
                if torch.cuda.is_available():
                    torch.cuda.empty_cache()
            except ImportError:
                pass
            logger.info("Model unloaded, memory freed")
