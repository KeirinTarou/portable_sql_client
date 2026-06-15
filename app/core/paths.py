from pathlib import Path
import sys

def get_base_dir() -> Path:
    if getattr(sys, 'frozen', False):
        # PyInstaller向け
        return Path(sys.executable).parent
    else:
        # このファイルから見て親の親の親
        return Path(__file__).resolve().parent.parent.parent