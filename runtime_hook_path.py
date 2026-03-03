# Runtime hook for PyInstaller: ensure DLLs (ctranslate2, CUDA, etc.)
# can be found when running from the frozen bundle.
import os
import sys

if getattr(sys, 'frozen', False):
    bundle_dir = sys._MEIPASS
    os.environ['PATH'] = bundle_dir + os.pathsep + os.environ.get('PATH', '')
