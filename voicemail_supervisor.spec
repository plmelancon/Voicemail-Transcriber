# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec for Voicemail Transcriber Supervisor.
Tiny parent process that relaunches VoicemailTranscriber.exe on crash.
Build with:  pyinstaller voicemail_supervisor.spec --clean
"""

block_cipher = None

a = Analysis(
    ['supervisor.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # Aggressive excludes — this exe does nothing but spawn/wait
        'torch',
        'torchaudio',
        'torchvision',
        'whisper',
        'faster_whisper',
        'ctranslate2',
        'openvino',
        'openvino_genai',
        'openvino_tokenizers',
        'onnxruntime',
        'numpy',
        'scipy',
        'librosa',
        'soundfile',
        'soxr',
        'av',
        'transformers',
        'tokenizers',
        'huggingface_hub',
        'hf_xet',
        'pystray',
        'PIL',
        'pythoncom',
        'win32com',
        'pywintypes',
        'tkinter',
        '_tkinter',
        'test',
        'xmlrpc',
        'flask',
        'mcp',
        'lxml',
        'pdfminer',
        'pdfplumber',
        'PyMuPDF',
        'fitz',
        'pypdf',
        'numba',
        'llvmlite',
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
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='VoicemailSupervisor',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    runtime_tmpdir=None,
    console=False,  # Runs silently in the background
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='app_icon.ico',
)
