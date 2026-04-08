# Runtime hook for PyInstaller: ensure DLLs (ctranslate2, CUDA, etc.)
# can be found when running from the frozen bundle.
import os
import sys

if getattr(sys, 'frozen', False):
    bundle_dir = sys._MEIPASS
    os.environ['PATH'] = bundle_dir + os.pathsep + os.environ.get('PATH', '')
    # Disable cuDNN in ctranslate2 — the bundle only includes cuBLAS DLLs.
    # Without this, ctranslate2 tries to load cudnn sub-libraries that aren't
    # present, causing a hard crash during CUDA inference.
    os.environ['CT2_USE_CUDNN'] = '0'
