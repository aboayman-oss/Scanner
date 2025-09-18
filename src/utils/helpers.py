"""Shared constants and helpers for the RFID Attendance Manager UI."""
import os
import sys
from pathlib import Path

import pandas as pd

def get_runtime_base():
    """Return the folder containing the script or executable."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    module_dir = Path(__file__).resolve().parent
    return str(module_dir.parent)


def get_assets_dir():
    """Locate bundled assets for development and PyInstaller builds.

    PyInstaller one-file executables extract bundled data into a temporary
    folder exposed via ``sys._MEIPASS``. We read assets from there so the
    binary stays portable.
    """
    if getattr(sys, 'frozen', False):
        base = getattr(sys, '_MEIPASS', get_runtime_base())
        return os.path.join(base, 'assets')
    script_dir = get_runtime_base()
    local_assets = os.path.join(script_dir, 'assets')
    if os.path.exists(os.path.join(local_assets, 'logo.png')):
        return local_assets
    parent_dir = os.path.dirname(script_dir)
    parent_assets = os.path.join(parent_dir, 'assets')
    if os.path.exists(os.path.join(parent_assets, 'logo.png')):
        return parent_assets
    return local_assets


RUNTIME_BASE = get_runtime_base()
# Assets live beside the script during development and inside the temporary
# PyInstaller bundle when frozen, so we centralize their path resolution above.
ASSETS_DIR = get_assets_dir()

if getattr(sys, 'frozen', False):
    # Keep user-generated data next to the executable for portability.
    BASE_FOLDER = RUNTIME_BASE
else:
    BASE_FOLDER = os.path.dirname(ASSETS_DIR)

LOGO_FILE        = os.path.join(ASSETS_DIR, 'logo.png')
HOME_BG_FILE     = os.path.join(ASSETS_DIR, 'background.jpg')
SETTINGS_BG_FILE = os.path.join(ASSETS_DIR, 'backgroundnew.jpg')

DATA_FOLDER      = os.path.join(BASE_FOLDER, 'Data')
SESSIONS_FOLDER  = os.path.join(BASE_FOLDER, 'Sessions')
ARCHIVE_FOLDER   = os.path.join(BASE_FOLDER, 'Data archive')
MAPPING_FILE     = os.path.join(ARCHIVE_FOLDER, 'column_map.json')
SETTINGS_FILE    = os.path.join(ARCHIVE_FOLDER, 'app_settings.json')
LAST_DATA_FILE   = os.path.join(ARCHIVE_FOLDER, 'last_data.json')

MIN_DASHBOARD_SIZE     = (980, 640)
MIN_SCAN_SIZE          = (900, 560)
MIN_SETTINGS_SIZE      = (640, 480)
MIN_SESSION_SETUP_SIZE = (360, 240)
MIN_SUMMARY_SIZE       = (380, 320)
MIN_PAST_SESSIONS_SIZE = (720, 480)
for folder in (DATA_FOLDER, SESSIONS_FOLDER, ARCHIVE_FOLDER):
    os.makedirs(folder, exist_ok=True)

SETTINGS = {
    "stage_options":  ["2nd", "3rd"],
    "center_options": [
        "October", "Ferdous", "Helwan", "Hadayek Helwan",
        "Zayed", "Haram", "Dokki", "Maadi", "15 May"
    ],
    "restrictions": {"exam": True, "homework": True},
    "file_type": "xlsx"
}



def bring_window_to_front(window):
    """Raise a toplevel window above its siblings and give it focus."""
    if window is None:
        return
    try:
        window.deiconify()
    except Exception:
        pass
    try:
        window.lift()
    except Exception:
        pass
    try:
        window.focus_force()
    except Exception:
        pass
    try:
        window.attributes('-topmost', True)
        window.after_idle(lambda: window.attributes('-topmost', False))
    except Exception:
        pass


def ensure_initial_size(window, *, min_size=None, padding=(0, 0)):
    """Size a toplevel so its default geometry fits the current layout."""
    if window is None:
        return 0, 0
    window.update_idletasks()
    req_w = max(window.winfo_reqwidth(), window.winfo_width())
    req_h = max(window.winfo_reqheight(), window.winfo_height())
    pad_x, pad_y = padding if isinstance(padding, tuple) else (padding, padding)
    width = max(int(req_w + pad_x), 1)
    height = max(int(req_h + pad_y), 1)
    if min_size:
        min_w, min_h = min_size
        width = max(width, int(min_w))
        height = max(height, int(min_h))
    window.minsize(width, height)
    window.geometry(f"{width}x{height}")
    return width, height

def read_data(path, **kwargs):
    if path.lower().endswith(".xlsx"):
        return pd.read_excel(path, dtype=str, **kwargs)
    else:
        return pd.read_csv(path, dtype=str, **kwargs)

def write_data(df, path, **kwargs):
    if path.lower().endswith(".xlsx"):
        df.to_excel(path, index=False, **kwargs)
    else:
        df.to_csv(path, index=False, **kwargs)
