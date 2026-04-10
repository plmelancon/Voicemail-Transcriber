# -*- coding: utf-8 -*-
"""
Resource Manager for Voicemail Transcriber Service
Handles process priority and Whisper model lifecycle.
Supports three backends: faster-whisper on CPU, faster-whisper on CUDA GPU,
and OpenVINO GenAI on Intel NPU.
"""

import gc
import json
import logging
import os
import sys as _sys
import threading
import time
from pathlib import Path

import win32api
import win32process

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Settings persistence
# ---------------------------------------------------------------------------
SETTINGS_DIR = Path(os.environ.get("LOCALAPPDATA", str(Path.home()))) / "VoicemailTranscriber"
SETTINGS_FILE = SETTINGS_DIR / "settings.json"


def load_settings() -> dict:
    """Load settings from disk. Returns defaults if file doesn't exist."""
    defaults = {"device": "auto"}
    try:
        if SETTINGS_FILE.exists():
            with open(SETTINGS_FILE, encoding="utf-8") as f:
                data = json.load(f)
            # Merge with defaults for any missing keys
            return {**defaults, **data}
    except Exception as e:
        logger.warning(f"Could not load settings: {e}")
    return defaults


def save_settings(settings: dict):
    """Save settings to disk."""
    try:
        SETTINGS_DIR.mkdir(parents=True, exist_ok=True)
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(settings, f, indent=2)
        logger.info(f"Settings saved: {settings}")
    except Exception as e:
        logger.error(f"Could not save settings: {e}")

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


# ---------------------------------------------------------------------------
# GPU Whisper model manager (faster-whisper on CUDA)
# ---------------------------------------------------------------------------

class GpuWhisperManager:
    """Whisper transcription backend using faster-whisper on a CUDA GPU.

    Same ``transcribe(audio_path) -> str`` interface as WhisperModelManager.

    Features:
        - Lazy model loading on first transcription
        - Automatic unload after IDLE_TIMEOUT seconds of inactivity
        - VRAM freed on unload via torch.cuda.empty_cache()
        - Thread-safe (guarded by a Lock)
    """

    IDLE_TIMEOUT = 300  # 5 minutes

    def __init__(self, model_name: str = "turbo", language: str = "fr"):
        self.model_name = model_name
        self.language = language
        self.device = "cuda"
        self.compute_type = "float16"
        self._model = None
        self._lock = threading.Lock()
        self._last_used: float = 0
        self._unload_timer: threading.Timer | None = None

    def transcribe(self, audio_path: str) -> str:
        """Transcribe an audio file on GPU and return the text."""
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
                logger.info("Model loaded successfully on GPU")
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
        """Immediately unload the model and free VRAM."""
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
            logger.info("Unloading Whisper GPU model (idle timeout)...")
            del self._model
            self._model = None
            gc.collect()
            logger.info("GPU model unloaded, VRAM freed")


# ---------------------------------------------------------------------------
# NPU Whisper model manager (OpenVINO GenAI)
# ---------------------------------------------------------------------------

def _load_audio_16k(audio_path: str):
    """Load an audio file and resample to 16 kHz mono float32.

    Tries soundfile first (WAV/FLAC/OGG), falls back to PyAV for
    formats soundfile can't handle (MP4/M4A/MP3/WMA/etc.).
    """
    import numpy as np
    import soxr

    try:
        import soundfile as sf
        data, sr = sf.read(audio_path, dtype="float32", always_2d=True)
    except Exception:
        # Fallback: decode with PyAV (FFmpeg) for MP4/M4A/MP3/etc.
        data, sr = _load_audio_av(audio_path)

    # Mix to mono
    if data.ndim > 1 and data.shape[1] > 1:
        data = data.mean(axis=1)
    elif data.ndim > 1:
        data = data[:, 0]
    # Resample to 16 kHz if needed
    if sr != 16000:
        data = soxr.resample(data, sr, 16000)
    return data


def _load_audio_av(audio_path: str):
    """Load audio using PyAV (FFmpeg). Returns (float32 ndarray, sample_rate)."""
    import av
    import numpy as np

    container = av.open(audio_path)
    stream = container.streams.audio[0]
    sr = stream.rate or stream.codec_context.sample_rate

    frames = []
    for frame in container.decode(audio=0):
        arr = frame.to_ndarray()
        # av returns shape (channels, samples) for planar, (samples, channels) for interleaved
        if arr.ndim == 2 and arr.shape[0] < arr.shape[1]:
            arr = arr.T  # channels-first → samples-first
        frames.append(arr)
    container.close()

    data = np.concatenate(frames, axis=0).astype(np.float32)
    # Normalize int formats to [-1, 1] range
    if data.max() > 1.0 or data.min() < -1.0:
        max_val = max(abs(data.max()), abs(data.min()))
        if max_val > 0:
            data = data / max_val
    if data.ndim == 1:
        data = data.reshape(-1, 1)
    return data, sr


class NpuWhisperManager:
    """Whisper transcription backend using Intel NPU via OpenVINO GenAI.

    Same ``transcribe(audio_path) -> str`` interface as WhisperModelManager
    so it can be swapped in transparently.

    Features:
        - Background pipeline loading at startup (non-blocking)
        - Pipeline stays loaded permanently (no idle timeout)
        - Thread-safe (guarded by a Lock)
    """

    MODEL_DIR = Path("C:/whisper-npu/models/whisper-large-v3-turbo-int8")
    DEVICE = "NPU"

    def __init__(self, model_name: str = "turbo", language: str = "fr"):
        self.model_name = model_name  # kept for interface compatibility
        self.language = language
        self._model = None
        self._lock = threading.Lock()
        self._ready = threading.Event()
        self._load_error: str | None = None
        # Start loading immediately in background
        self._loader = threading.Thread(
            target=self._background_load, name="NPU-Loader", daemon=True
        )
        self._loader.start()

    def _background_load(self):
        """Load the pipeline in a background thread so the UI stays responsive."""
        try:
            import openvino_genai
            logger.info(f"Loading OpenVINO Whisper pipeline on {self.DEVICE}...")
            logger.info(f"  Model: {self.MODEL_DIR}")
            logger.info("  First load compiles for NPU — may take 2-3 minutes.")
            start = time.time()
            pipe = openvino_genai.WhisperPipeline(
                str(self.MODEL_DIR), self.DEVICE
            )
            with self._lock:
                self._model = pipe
            logger.info(f"Pipeline ready in {time.time() - start:.1f}s")
        except Exception as e:
            self._load_error = str(e)
            logger.error(f"Failed to load NPU pipeline: {e}")
        finally:
            self._ready.set()

    # -- public interface (duck-type compatible with WhisperTranscriber) ------

    def load_model(self):
        """Wait for the background load to finish, then return the pipeline."""
        self._ready.wait()
        if self._load_error:
            raise RuntimeError(f"NPU pipeline failed to load: {self._load_error}")
        return self._model

    def transcribe(self, audio_path: str) -> str:
        """Transcribe an audio file and return the text.

        Waits for the pipeline to be ready (background load), then runs
        inference under the lock (NPU is not thread-safe).
        """
        self._ready.wait()
        if self._load_error:
            raise RuntimeError(f"NPU pipeline failed to load: {self._load_error}")

        with self._lock:
            logger.info(f"Transcribing: {audio_path}")
            try:
                speech = _load_audio_16k(audio_path)

                config = self._model.get_generation_config()
                config.task = "transcribe"
                config.language = f"<|{self.language}|>"
                config.max_new_tokens = 448

                result = self._model.generate(speech, config)
                text = str(result).strip()
                logger.info(f"Transcription complete: {len(text)} characters")
                return text
            except Exception as e:
                logger.error(f"Transcription error: {e}")
                raise

    def force_unload(self):
        """Immediately unload the pipeline and free resources."""
        with self._lock:
            self._do_unload()

    # -- internal ------------------------------------------------------------

    def _do_unload(self):
        if self._model is not None:
            logger.info("Unloading Whisper NPU pipeline (idle timeout)...")
            del self._model
            self._model = None
            gc.collect()
            logger.info("Pipeline unloaded, memory freed")


# ---------------------------------------------------------------------------
# Factory: create the right transcriber based on device preference
# ---------------------------------------------------------------------------

def is_npu_available() -> bool:
    """Check if Intel NPU is available via OpenVINO."""
    try:
        import openvino as ov
        core = ov.Core()
        return "NPU" in core.available_devices
    except Exception:
        return False


def is_gpu_available() -> bool:
    """Check if a CUDA-capable GPU is available via ctranslate2."""
    try:
        import ctranslate2
        return ctranslate2.get_cuda_device_count() > 0
    except Exception:
        return False


class AutoFallbackTranscriber:
    """Tries backends in priority order: NPU → GPU → CPU.

    Each backend is attempted at inference time.  If a fatal device error is
    detected (e.g. ZE_RESULT_ERROR_DEVICE_LOST for NPU, CUDA OOM/device
    errors for GPU), the failing backend is unloaded and the next one in the
    chain is created lazily and retried.
    """

    _NPU_FATAL_KEYWORDS = ("DEVICE_LOST", "DEVICE_HUNG", "device hung")
    _GPU_FATAL_KEYWORDS = (
        "CUDA", "cuda", "out of memory", "device-side assert",
        "CUBLAS", "CUDNN", "NCCL",
    )
    # Modules that OpenVINO's NPU plugin injects as partial stubs into
    # sys.modules. After an NPU failure these stubs linger and break any
    # later faster-whisper / ctranslate2 import in the same process
    # (symptom: "module 'torch' has no attribute 'Tensor'"). Purging them
    # lets the next backend re-import cleanly.
    _POISONED_MODULES = ("torch", "torchvision", "torchaudio")

    def __init__(
        self,
        npu: "NpuWhisperManager | None",
        gpu_available: bool = False,
        language: str = "fr",
    ):
        self._language = language
        self._gpu_available = gpu_available
        # The ordered list of backends to try. Entries are removed as they fail.
        self._backends: list[tuple[str, object | None]] = []
        if npu is not None:
            self._backends.append(("npu", npu))
        if gpu_available:
            self._backends.append(("gpu", None))   # created lazily
        self._backends.append(("cpu", None))        # created lazily

    def _get_or_create(self, name: str, instance: object | None):
        """Return existing instance or create one on demand."""
        if instance is not None:
            return instance
        if name == "gpu":
            return GpuWhisperManager(language=self._language)
        if name == "cpu":
            return WhisperModelManager(language=self._language)
        raise RuntimeError(f"Unknown backend: {name}")

    def _purge_poisoned_modules(self):
        """Remove partial torch stubs from sys.modules so the next backend
        can re-import cleanly. See comment on _POISONED_MODULES above.
        """
        import sys
        purged = []
        for mod in list(sys.modules):
            if mod in self._POISONED_MODULES or any(
                mod.startswith(p + ".") for p in self._POISONED_MODULES
            ):
                sys.modules.pop(mod, None)
                purged.append(mod)
        if purged:
            logger.info("Purged %d poisoned module(s) from sys.modules", len(purged))

    def transcribe(self, audio_path: str) -> str:
        last_exc: Exception | None = None

        for idx, (name, instance) in enumerate(list(self._backends)):
            try:
                backend = self._get_or_create(name, instance)
            except Exception as exc:
                # Backend construction itself failed (e.g. poisoned torch
                # module breaking WhisperModel init). Treat as fatal for
                # this backend and move on.
                logger.warning(
                    "%s backend creation failed, skipping: %s", name.upper(), exc,
                )
                self._backends.pop(idx)
                self._purge_poisoned_modules()
                return self.transcribe(audio_path)

            # Persist the created instance back so we don't recreate it
            self._backends[idx] = (name, backend)
            try:
                return backend.transcribe(audio_path)
            except Exception as exc:
                exc_str = str(exc)
                is_fatal = (
                    (name == "npu" and any(k in exc_str for k in self._NPU_FATAL_KEYWORDS))
                    or (name == "gpu" and any(k in exc_str for k in self._GPU_FATAL_KEYWORDS))
                )
                if is_fatal and idx < len(self._backends) - 1:
                    next_name = self._backends[idx + 1][0]
                    logger.warning(
                        "%s device error, falling back to %s: %s",
                        name.upper(), next_name.upper(), exc,
                    )
                    try:
                        backend.force_unload()
                    except Exception as unload_exc:
                        logger.warning(
                            "%s force_unload failed (ignored): %s",
                            name.upper(), unload_exc,
                        )
                    # Remove the failed backend from the chain
                    self._backends.pop(idx)
                    # NPU failure can leave stub torch modules in sys.modules
                    # that break the next backend's init — purge them.
                    if name == "npu":
                        self._purge_poisoned_modules()
                    # Retry with the rest of the chain (recursive-safe via loop)
                    return self.transcribe(audio_path)
                last_exc = exc

        # All backends exhausted — re-raise the last error
        raise last_exc  # type: ignore[misc]

    def force_unload(self):
        for _name, instance in self._backends:
            if instance is not None and hasattr(instance, "force_unload"):
                instance.force_unload()

    @property
    def active_backend(self) -> str:
        """Return the name of the first (highest-priority) backend still in the chain."""
        return self._backends[0][0] if self._backends else "none"


def create_transcriber(device_pref: str = "auto", language: str = "fr"):
    """Create the appropriate transcriber backend.

    Parameters
    ----------
    device_pref : str
        ``"auto"`` — try NPU → GPU → CPU, with runtime fallback on errors.
        ``"npu"``  — force OpenVINO NPU.
        ``"gpu"``  — force faster-whisper on CUDA GPU.
        ``"cpu"``  — force faster-whisper on CPU.
    language : str
        Language code for transcription (default: ``"fr"``).

    Returns
    -------
    WhisperModelManager, GpuWhisperManager, NpuWhisperManager, or AutoFallbackTranscriber
    """
    if device_pref == "npu":
        logger.info("Device preference: NPU (forced)")
        return NpuWhisperManager(language=language)

    if device_pref == "gpu":
        logger.info("Device preference: GPU (forced)")
        return GpuWhisperManager(language=language)

    if device_pref == "cpu":
        logger.info("Device preference: CPU (forced)")
        return WhisperModelManager(language=language)

    # auto: build the NPU → GPU → CPU fallback chain
    npu_ok = is_npu_available() and NpuWhisperManager.MODEL_DIR.exists()
    gpu_ok = is_gpu_available()

    npu = NpuWhisperManager(language=language) if npu_ok else None
    chain_parts = []
    if npu_ok:
        chain_parts.append("NPU")
    if gpu_ok:
        chain_parts.append("GPU")
    chain_parts.append("CPU")
    logger.info("Device preference: auto → chain: %s", " → ".join(chain_parts))

    # If only CPU is available, skip the wrapper overhead
    if not npu_ok and not gpu_ok:
        return WhisperModelManager(language=language)

    return AutoFallbackTranscriber(
        npu=npu,
        gpu_available=gpu_ok,
        language=language,
    )
