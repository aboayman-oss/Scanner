"""Shared constants and helpers for the RFID Attendance Manager UI."""
import json
import os
import sys
import ctypes
from ctypes import byref, c_int, c_uint, c_void_p, c_size_t, wintypes
from pathlib import Path

import pandas as pd

INVALID_PATH_CHARS = set('<>::"/\\|?*')

DWMWA_USE_IMMERSIVE_DARK_MODE = 20
DWMWA_USE_IMMERSIVE_DARK_MODE_BEFORE_20H1 = 19
DWMWA_CAPTION_COLOR = 35
DWMWA_TEXT_COLOR = 36
WCA_USEDARKMODECOLORS = 26

_DARK_MODE_APP_INITIALIZED = False


class WINDOWCOMPOSITIONATTRIBDATA(ctypes.Structure):
    _fields_ = [("Attribute", c_int), ("Data", c_void_p), ("SizeOfData", c_size_t)]


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
PAST_SESSIONS_ICON_FILE = os.path.join(ASSETS_DIR, 'past sessions.png')
SETTINGS_ICON_FILE = os.path.join(ASSETS_DIR, 'settings.png')
IMPORT_ICON_FILE = os.path.join(ASSETS_DIR, 'import.png')
NEW_SESSION_ICON_FILE = os.path.join(ASSETS_DIR, 'add.png')
DASHBOARD_ICON_FILE = os.path.join(ASSETS_DIR, 'dashboard.png')
FOLDER_OPEN_ICON_FILE = os.path.join(ASSETS_DIR, 'folder_open.png')
HOME_BG_FILE     = os.path.join(ASSETS_DIR, 'background.jpg')
SETTINGS_BG_FILE = os.path.join(ASSETS_DIR, 'backgroundnew.jpg')
STATUS_OK_ICON_FILE = os.path.join(ASSETS_DIR, 'check_circle.png')
STATUS_INFO_ICON_FILE = os.path.join(ASSETS_DIR, 'warning.png')
REMOVE_ICON_FILE = os.path.join(ASSETS_DIR, 'close.png')
PLUS_ICON_FILE = os.path.join(ASSETS_DIR, 'add.png')

DATA_FOLDER      = os.path.join(BASE_FOLDER, 'Data')
DEFAULT_SESSIONS_FOLDER = os.path.join(BASE_FOLDER, 'Sessions')
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
for folder in (DATA_FOLDER, DEFAULT_SESSIONS_FOLDER, ARCHIVE_FOLDER):
    os.makedirs(folder, exist_ok=True)

SETTINGS = {
    "sessions_folder": DEFAULT_SESSIONS_FOLDER,
    "stage_options":  ["2nd", "3rd"],
    "center_options": [
        "October", "Ferdous", "Helwan", "Hadayek Helwan",
        "Zayed", "Haram", "Dokki", "Maadi", "15 May"
    ],
    "restrictions": {"exam": True, "homework": True},
    "file_type": "xlsx"

}


"""Session folder helpers"""

def _normalize_folder_path(path):
    if not path:
        return DEFAULT_SESSIONS_FOLDER
    normalized = os.path.abspath(path)
    return normalized


def ensure_directory(path):
    os.makedirs(path, exist_ok=True)
    return path


def get_sessions_folder():
    path = SETTINGS.get("sessions_folder") or DEFAULT_SESSIONS_FOLDER
    normalized = _normalize_folder_path(path)
    return ensure_directory(normalized)


def set_sessions_folder(path):
    normalized = _normalize_folder_path(path)
    ensure_directory(normalized)
    SETTINGS["sessions_folder"] = normalized
    return normalized



def _sanitize_path_component(value, fallback):
    raw = (value or '').strip()
    if not raw:
        return fallback
    sanitized = []
    for ch in raw:
        if ch in INVALID_PATH_CHARS:
            sanitized.append('_')
        else:
            sanitized.append(ch)
    sanitized_str = ''.join(sanitized)
    sanitized_str = sanitized_str.replace(os.sep, '_')
    if os.altsep:
        sanitized_str = sanitized_str.replace(os.altsep, '_')
    sanitized_str = sanitized_str.strip()
    sanitized_str = sanitized_str.rstrip('.')
    if sanitized_str in ('', '.', '..'):
        return fallback
    return sanitized_str


def resolve_session_directory(stage=None, center=None, *, create=False):
    base = get_sessions_folder()
    target = base
    if stage:
        stage_component = _sanitize_path_component(stage, 'Stage')
        target = os.path.join(target, stage_component)
        if create:
            ensure_directory(target)
    if center:
        center_component = _sanitize_path_component(center, 'Center')
        target = os.path.join(target, center_component)
        if create:
            ensure_directory(target)
    if not (stage or center) and create:
        ensure_directory(target)
    return target


def resolve_session_file_path(name, *, stage=None, center=None, ext='csv', create=False):
    ext = (ext or '').lstrip('.')
    ext = ext or 'csv'
    directory = resolve_session_directory(stage, center, create=create)
    filename = f"{name}.{ext}"
    return os.path.join(directory, filename)

def save_settings():
    os.makedirs(os.path.dirname(SETTINGS_FILE), exist_ok=True)
    with open(SETTINGS_FILE, "w", encoding="utf-8") as file:
        json.dump(SETTINGS, file, indent=2)




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


def _ensure_app_dark_mode():
    global _DARK_MODE_APP_INITIALIZED
    if _DARK_MODE_APP_INITIALIZED or sys.platform != "win32":
        return
    try:
        uxtheme = ctypes.windll.uxtheme
    except Exception:
        _DARK_MODE_APP_INITIALIZED = True
        return
    try:
        allow_app = getattr(uxtheme, "AllowDarkModeForApp", None)
        if allow_app:
            allow_app.argtypes = [wintypes.BOOL]
            allow_app.restype = wintypes.BOOL
            allow_app(wintypes.BOOL(True))
        set_app_mode = getattr(uxtheme, "SetPreferredAppMode", None)
        if set_app_mode:
            set_app_mode.argtypes = [c_int]
            set_app_mode.restype = c_int
            # APPMODE_ALLOWDARK = 2
            set_app_mode(2)
        flush_themes = getattr(uxtheme, "FlushMenuThemes", None)
        if flush_themes:
            flush_themes()
    except Exception:
        pass
    _DARK_MODE_APP_INITIALIZED = True


def _get_window_handle(window):
    try:
        window.update_idletasks()
    except Exception:
        pass
    try:
        hwnd = window.winfo_id()
    except Exception:
        return None
    if not hwnd:
        return None
    try:
        get_ancestor = ctypes.windll.user32.GetAncestor
        get_ancestor.argtypes = [wintypes.HWND, c_uint]
        get_ancestor.restype = wintypes.HWND
        ancestor = get_ancestor(wintypes.HWND(hwnd), c_uint(2))
        if ancestor:
            hwnd = ancestor
    except Exception:
        try:
            parent = ctypes.windll.user32.GetParent(wintypes.HWND(hwnd))
            if parent:
                hwnd = parent
        except Exception:
            pass
    return int(hwnd) if hwnd else None



def _apply_dark_mode_to_hwnd(hwnd):
    applied = False
    hwnd = wintypes.HWND(hwnd)
    try:
        uxtheme = ctypes.windll.uxtheme
    except Exception:
        uxtheme = None
    if uxtheme:
        try:
            allow_window = getattr(uxtheme, "AllowDarkModeForWindow", None)
            if allow_window:
                allow_window.argtypes = [wintypes.HWND, wintypes.BOOL]
                allow_window.restype = wintypes.BOOL
                allow_window(hwnd, wintypes.BOOL(True))
        except Exception:
            pass
    try:
        dwm = ctypes.windll.dwmapi
    except Exception:
        dwm = None
    if dwm:
        try:
            dwm.DwmSetWindowAttribute.argtypes = [wintypes.HWND, c_uint, c_void_p, c_uint]
        except AttributeError:
            pass
        value = c_int(1)
        size_value = ctypes.sizeof(value)
        for attribute in (DWMWA_USE_IMMERSIVE_DARK_MODE, DWMWA_USE_IMMERSIVE_DARK_MODE_BEFORE_20H1):
            try:
                result = dwm.DwmSetWindowAttribute(hwnd, c_uint(attribute), byref(value), size_value)
            except Exception:
                continue
            if result == 0:
                applied = True
                break
        caption_color = c_uint(0x001A1A1A)
        text_color = c_uint(0x00FFFFFF)
        size_color = ctypes.sizeof(caption_color)
        for attribute, data in ((DWMWA_CAPTION_COLOR, caption_color), (DWMWA_TEXT_COLOR, text_color)):
            try:
                result = dwm.DwmSetWindowAttribute(hwnd, c_uint(attribute), byref(data), size_color)
            except Exception:
                continue
            if result == 0:
                applied = True
    if not applied:
        try:
            user32 = ctypes.windll.user32
            set_attr = getattr(user32, "SetWindowCompositionAttribute")
        except Exception:
            return applied
        try:
            set_attr.argtypes = [wintypes.HWND, ctypes.POINTER(WINDOWCOMPOSITIONATTRIBDATA)]
        except AttributeError:
            pass
        try:
            use_dark = wintypes.BOOL(True)
            data = WINDOWCOMPOSITIONATTRIBDATA(WCA_USEDARKMODECOLORS, ctypes.byref(use_dark), c_size_t(ctypes.sizeof(use_dark)))
            set_attr(hwnd, byref(data))
            applied = True
        except Exception:
            pass
    return applied



def set_dark_title_bar(window, *, attempts=12, interval_ms=90):
    '''Request dark mode for the native title bar, retrying while the window settles.'''
    if window is None or sys.platform != "win32":
        return
    _ensure_app_dark_mode()
    try:
        if not getattr(window, "_dark_title_bar_bound", False):
            def _remap(_event=None):
                set_dark_title_bar(window, attempts=attempts, interval_ms=interval_ms)
            window._dark_title_bar_bound = True
            window.bind("<Map>", _remap, add="+")
            window.bind("<FocusIn>", _remap, add="+")
    except Exception:
        pass
    try:
        job = getattr(window, "_dark_title_job", None)
        if job:
            window.after_cancel(job)
    except Exception:
        pass
    attempts = max(int(attempts), 1)

    def _attempt(remaining):
        hwnd = _get_window_handle(window)
        if hwnd is not None:
            _apply_dark_mode_to_hwnd(hwnd)
        remaining -= 1
        if remaining <= 0:
            return
        try:
            window._dark_title_job = window.after(interval_ms, lambda: _attempt(remaining))
        except Exception:
            pass

    try:
        window._dark_title_job = window.after(0, lambda: _attempt(attempts))
    except Exception:
        _attempt(attempts)


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
