# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec for Voicemail Transcriber Service.
Build with:  pyinstaller voicemail_service.spec --clean
Supports CPU, CUDA GPU, and Intel NPU backends.
"""

import os
import importlib

block_cipher = None

# --- Locate package directories for binary/data collection ---
ct2_dir = os.path.dirname(importlib.import_module('ctranslate2').__file__)
av_dir = os.path.dirname(importlib.import_module('av').__file__)
av_libs = os.path.join(os.path.dirname(av_dir), 'av.libs')
ort_dir = os.path.dirname(importlib.import_module('onnxruntime').__file__)

# CUDA runtime DLLs (needed by ctranslate2 for GPU inference).
# We copy them to a temp dir so PyInstaller never sees a path containing
# "torch" — otherwise it creates a partial torch/ package in the bundle
# which shadows the excluded torch and breaks ctranslate2 imports.
import shutil, tempfile
_torch_lib = os.path.join(os.path.dirname(importlib.import_module('torch').__file__), 'lib')
_cuda_tmp = tempfile.mkdtemp(prefix='cuda_dlls_')
CUDA_DLLS = {}
for _name in ['cublas64_12.dll', 'cublasLt64_12.dll', 'cudart64_12.dll']:
    _src = os.path.join(_torch_lib, _name)
    _dst = os.path.join(_cuda_tmp, _name)
    shutil.copy2(_src, _dst)
    CUDA_DLLS[_name] = _dst
del _torch_lib, _src, _dst, _name

# OpenVINO / GenAI directories (NPU backend)
ov_dir = os.path.dirname(importlib.import_module('openvino').__file__)
ov_libs = os.path.join(ov_dir, 'libs')
ov_genai_dir = os.path.dirname(importlib.import_module('openvino_genai').__file__)
ov_tok_dir = os.path.dirname(importlib.import_module('openvino_tokenizers').__file__)
ov_tok_libs = os.path.join(ov_tok_dir, 'lib')

a = Analysis(
    ['voicemail_service.py'],
    pathex=[],
    binaries=[
        # ctranslate2 native DLLs (CPU backend / CUDA support compiled in)
        (os.path.join(ct2_dir, 'ctranslate2.dll'), 'ctranslate2'),
        (os.path.join(ct2_dir, 'cudnn64_9.dll'), 'ctranslate2'),
        (os.path.join(ct2_dir, 'libiomp5md.dll'), 'ctranslate2'),
        (os.path.join(ct2_dir, '_ext.cp311-win_amd64.pyd'), 'ctranslate2'),
        # CUDA runtime DLLs (required by ctranslate2 for GPU inference)
        (CUDA_DLLS['cublas64_12.dll'], '.'),
        (CUDA_DLLS['cublasLt64_12.dll'], '.'),
        (CUDA_DLLS['cudart64_12.dll'], '.'),
        # OpenVINO GenAI native DLLs (NPU backend)
        (os.path.join(ov_genai_dir, 'openvino_genai.dll'), 'openvino_genai'),
        (os.path.join(ov_genai_dir, 'py_openvino_genai.cp311-win_amd64.pyd'), 'openvino_genai'),
    ],
    datas=[
        # ctranslate2 Python package (full directory)
        (ct2_dir, 'ctranslate2'),
        # av.libs (FFmpeg codec DLLs)
        (av_libs, 'av.libs'),
        # onnxruntime capi (contains onnxruntime.dll etc.)
        (os.path.join(ort_dir, 'capi'), 'onnxruntime/capi'),
        # OpenVINO core libs (includes NPU plugin, CPU plugin, TBB, etc.)
        (ov_libs, 'openvino/libs'),
        # OpenVINO GenAI package
        (ov_genai_dir, 'openvino_genai'),
        # OpenVINO Tokenizers native libs (ICU + tokenizer DLL)
        (ov_tok_libs, 'openvino_tokenizers/lib'),
        # OpenVINO Tokenizers Python package
        (ov_tok_dir, 'openvino_tokenizers'),
    ],
    hiddenimports=[
        # pywin32 COM automation
        'pythoncom',
        'pywintypes',
        'win32api',
        'win32process',
        'win32com',
        'win32com.client',
        'win32com.client.util',
        'win32com.client.dynamic',
        'win32com.client.gencache',
        'win32com.server',
        'win32com.server.util',
        # pystray
        'pystray._win32',
        # faster-whisper and ctranslate2 (CPU backend)
        'faster_whisper',
        'faster_whisper.audio',
        'faster_whisper.feature_extractor',
        'faster_whisper.tokenizer',
        'faster_whisper.transcribe',
        'faster_whisper.utils',
        'faster_whisper.vad',
        'ctranslate2',
        'ctranslate2.converters',
        'ctranslate2.extensions',
        'ctranslate2.specs',
        # onnxruntime (for VAD)
        'onnxruntime',
        'onnxruntime.capi',
        'onnxruntime.capi._pybind_state',
        # OpenVINO GenAI (NPU backend)
        'openvino',
        'openvino.runtime',
        'openvino.runtime.opset14',
        'openvino_genai',
        'openvino_tokenizers',
        # Audio loading for NPU backend (soundfile + soxr, NOT librosa)
        'soundfile',
        'soxr',
        # huggingface model download
        'huggingface_hub',
        'huggingface_hub.file_download',
        'huggingface_hub.hf_api',
        'huggingface_hub.utils',
        'hf_xet',
        # tokenizers
        'tokenizers',
        # av (PyAV) for audio
        'av',
        'av.audio',
        'av.audio.fifo',
        'av.audio.format',
        'av.audio.frame',
        'av.audio.layout',
        'av.audio.resampler',
        'av.audio.stream',
        'av.container',
        # Project modules
        'voicemail_transcriber',
        'resource_manager',
        'outlook_monitor',
        'tray_app',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=['runtime_hook_path.py'],
    excludes=[
        # CRITICAL: Exclude torch and openai-whisper (not needed, saves 4+ GB)
        'torch',
        'torchaudio',
        'torchvision',
        'whisper',
        # Other unnecessary packages
        'tkinter',
        '_tkinter',
        # 'unittest',  # needed by scipy → numpy.testing
        'test',
        'xmlrpc',
        # 'pydoc',  # needed by scipy/librosa import chain
        # 'doctest',  # needed by scipy/librosa import chain
        'flask',
        'flask_cors',
        'uvicorn',
        'starlette',
        'mcp',
        'lxml',
        'pdfminer',
        'pdfplumber',
        'PyMuPDF',
        'fitz',
        'pypdf',
        'pypdfium2',
        'python_docx',
        'docx',
        'keyboard',
        'numba',  # large, only needed by librosa (we use soundfile instead)
        'llvmlite',  # numba dependency
        'tiktoken',
        'waybackpack',
        'pyinstaller',
        'altgraph',
        'pefile',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='VoicemailTranscriber',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,  # System tray app — no console window
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='app_icon.ico',
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='VoicemailTranscriber',
)
