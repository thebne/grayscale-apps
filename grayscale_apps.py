"""
Per-window grayscale overlay using the Windows Magnification API.

Creates a click-through overlay with a magnifier control that re-renders the
target window's content in grayscale. Only the target app appears black & white;
everything else stays in full color.

Configure target process names in config.json.
"""

import ctypes
import ctypes.wintypes as wt
import os
import sys
import json
import time
import atexit

# ---------------------------------------------------------------------------
# DPI awareness (must be set before any window/API calls)
# ---------------------------------------------------------------------------

try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)  # per-monitor v2
except Exception:
    ctypes.windll.user32.SetProcessDPIAware()

# ---------------------------------------------------------------------------
# Libraries
# ---------------------------------------------------------------------------

user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32
mag = ctypes.WinDLL("Magnification.dll")

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

WS_CHILD = 0x40000000
WS_VISIBLE = 0x10000000
WS_POPUP = 0x80000000
WS_CLIPCHILDREN = 0x02000000
WS_EX_TOPMOST = 0x00000008
WS_EX_LAYERED = 0x00080000
WS_EX_TRANSPARENT = 0x00000020
WS_EX_TOOLWINDOW = 0x00000080
WS_EX_NOACTIVATE = 0x08000000

LWA_ALPHA = 0x02
SW_HIDE = 0
SW_SHOWNOACTIVATE = 4
SWP_NOACTIVATE = 0x0010
HWND_TOPMOST = wt.HWND(-1)

WM_DESTROY = 0x0002
PM_REMOVE = 0x0001

PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
MS_SHOWMAGNIFIEDCURSOR = 0x0001
MW_FILTERMODE_INCLUDE = 1

# ---------------------------------------------------------------------------
# Types
# ---------------------------------------------------------------------------

WNDPROC = ctypes.WINFUNCTYPE(
    ctypes.c_long, wt.HWND, ctypes.c_uint, wt.WPARAM, wt.LPARAM
)


HANDLE = ctypes.c_void_p


class WNDCLASSEX(ctypes.Structure):
    _fields_ = [
        ("cbSize", ctypes.c_uint),
        ("style", ctypes.c_uint),
        ("lpfnWndProc", WNDPROC),
        ("cbClsExtra", ctypes.c_int),
        ("cbWndExtra", ctypes.c_int),
        ("hInstance", HANDLE),
        ("hIcon", HANDLE),
        ("hCursor", HANDLE),
        ("hbrBackground", HANDLE),
        ("lpszMenuName", wt.LPCWSTR),
        ("lpszClassName", wt.LPCWSTR),
        ("hIconSm", HANDLE),
    ]


MAGCOLOREFFECT = ctypes.c_float * 25
MAGTRANSFORM = ctypes.c_float * 9

# ---------------------------------------------------------------------------
# Color matrices
# ---------------------------------------------------------------------------

# Windows uses row-vector * matrix, so each ROW holds one input channel's
# contribution to all output channels (R_out, G_out, B_out, A_out, offset).
GRAYSCALE = MAGCOLOREFFECT(
    0.299, 0.299, 0.299, 0, 0,  # R input  → contributes 0.299 to each RGB out
    0.587, 0.587, 0.587, 0, 0,  # G input  → contributes 0.587 to each RGB out
    0.114, 0.114, 0.114, 0, 0,  # B input  → contributes 0.114 to each RGB out
    0,     0,     0,     1, 0,  # A passthrough
    0,     0,     0,     0, 1,  # offset row
)

IDENTITY_TRANSFORM = MAGTRANSFORM(
    1, 0, 0,
    0, 1, 0,
    0, 0, 1,
)

# ---------------------------------------------------------------------------
# API declarations
# ---------------------------------------------------------------------------

mag.MagInitialize.restype = ctypes.c_bool
mag.MagUninitialize.restype = ctypes.c_bool

mag.MagSetColorEffect.argtypes = [wt.HWND, ctypes.POINTER(MAGCOLOREFFECT)]
mag.MagSetColorEffect.restype = ctypes.c_bool

mag.MagSetWindowSource.argtypes = [wt.HWND, wt.RECT]
mag.MagSetWindowSource.restype = ctypes.c_bool

mag.MagSetWindowTransform.argtypes = [wt.HWND, ctypes.POINTER(MAGTRANSFORM)]
mag.MagSetWindowTransform.restype = ctypes.c_bool

mag.MagSetWindowFilterList.argtypes = [
    wt.HWND, wt.DWORD, ctypes.c_int, ctypes.POINTER(wt.HWND)
]
mag.MagSetWindowFilterList.restype = ctypes.c_bool

user32.DefWindowProcW.argtypes = [wt.HWND, ctypes.c_uint, wt.WPARAM, wt.LPARAM]
user32.DefWindowProcW.restype = ctypes.c_long

user32.SetLayeredWindowAttributes.argtypes = [wt.HWND, wt.DWORD, wt.BYTE, wt.DWORD]
user32.SetLayeredWindowAttributes.restype = wt.BOOL

user32.RegisterClassExW.argtypes = [ctypes.POINTER(WNDCLASSEX)]
user32.RegisterClassExW.restype = wt.ATOM

user32.CreateWindowExW.argtypes = [
    wt.DWORD, wt.LPCWSTR, wt.LPCWSTR, wt.DWORD,
    ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int,
    wt.HWND, wt.HMENU, wt.HINSTANCE, wt.LPVOID,
]
user32.CreateWindowExW.restype = wt.HWND

user32.SetWindowPos.argtypes = [
    wt.HWND, wt.HWND,
    ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int,
    ctypes.c_uint,
]
user32.SetWindowPos.restype = wt.BOOL

user32.GetWindowRect.argtypes = [wt.HWND, ctypes.POINTER(wt.RECT)]
user32.GetWindowRect.restype = wt.BOOL

user32.InvalidateRect.argtypes = [wt.HWND, ctypes.c_void_p, wt.BOOL]
user32.InvalidateRect.restype = wt.BOOL

user32.ShowWindow.argtypes = [wt.HWND, ctypes.c_int]
user32.ShowWindow.restype = wt.BOOL

user32.IsIconic.argtypes = [wt.HWND]
user32.IsIconic.restype = wt.BOOL

user32.IsWindowVisible.argtypes = [wt.HWND]
user32.IsWindowVisible.restype = wt.BOOL

user32.DestroyWindow.argtypes = [wt.HWND]
user32.PostQuitMessage.argtypes = [ctypes.c_int]

user32.PeekMessageW.argtypes = [
    ctypes.POINTER(wt.MSG), wt.HWND,
    ctypes.c_uint, ctypes.c_uint, ctypes.c_uint,
]
user32.PeekMessageW.restype = wt.BOOL
user32.TranslateMessage.argtypes = [ctypes.POINTER(wt.MSG)]
user32.DispatchMessageW.argtypes = [ctypes.POINTER(wt.MSG)]

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(SCRIPT_DIR, "config.json")


def load_config():
    with open(CONFIG_PATH) as f:
        cfg = json.load(f)
    targets = {t.lower() for t in cfg.get("targets", [])}
    interval = cfg.get("poll_interval_ms", 16) / 1000.0
    return targets, interval


# ---------------------------------------------------------------------------
# Window helpers
# ---------------------------------------------------------------------------


def get_process_name(pid):
    handle = kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
    if not handle:
        return None
    try:
        buf = ctypes.create_unicode_buffer(260)
        size = wt.DWORD(260)
        if kernel32.QueryFullProcessImageNameW(handle, 0, buf, ctypes.byref(size)):
            return os.path.basename(buf.value).lower()
    finally:
        kernel32.CloseHandle(handle)
    return None


def find_target_window(targets):
    """Return the first visible, non-minimised window belonging to a target process."""
    result = []

    @ctypes.WINFUNCTYPE(wt.BOOL, wt.HWND, wt.LPARAM)
    def _cb(hwnd, _):
        if not user32.IsWindowVisible(hwnd):
            return True
        if user32.IsIconic(hwnd):
            return True
        pid = wt.DWORD()
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        name = get_process_name(pid.value)
        if name and name in targets:
            rect = wt.RECT()
            user32.GetWindowRect(hwnd, ctypes.byref(rect))
            if (rect.right - rect.left) > 50 and (rect.bottom - rect.top) > 50:
                result.append(hwnd)
                return False  # stop enumerating
        return True

    user32.EnumWindows(_cb, 0)
    return result[0] if result else None


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------


def main():
    targets, interval = load_config()

    if not mag.MagInitialize():
        print("ERROR: MagInitialize failed. Try running as Administrator.")
        sys.exit(1)

    # Keep a reference so the GC doesn't collect the callback
    @WNDPROC
    def wnd_proc(hwnd, msg, wp, lp):
        if msg == WM_DESTROY:
            user32.PostQuitMessage(0)
            return 0
        return user32.DefWindowProcW(hwnd, msg, wp, lp)

    hinstance = kernel32.GetModuleHandleW(None)
    cls_name = "GrayscaleOverlayHost"

    wc = WNDCLASSEX()
    wc.cbSize = ctypes.sizeof(WNDCLASSEX)
    wc.lpfnWndProc = wnd_proc
    wc.hInstance = hinstance
    wc.lpszClassName = cls_name

    if not user32.RegisterClassExW(ctypes.byref(wc)):
        print(f"ERROR: RegisterClassExW failed (err {kernel32.GetLastError()})")
        sys.exit(1)

    # Host: top-most, layered, click-through, invisible to taskbar
    hwnd_host = user32.CreateWindowExW(
        WS_EX_TOPMOST | WS_EX_LAYERED | WS_EX_TRANSPARENT
        | WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE,
        cls_name, "GrayscaleOverlay",
        WS_POPUP | WS_CLIPCHILDREN,
        0, 0, 1, 1,
        None, None, hinstance, None,
    )
    if not hwnd_host:
        print(f"ERROR: host CreateWindowExW failed (err {kernel32.GetLastError()})")
        sys.exit(1)

    # Full-opacity layered window (required for click-through)
    user32.SetLayeredWindowAttributes(hwnd_host, 0, 255, LWA_ALPHA)

    # Magnifier child — the Magnifier class is registered by MagInitialize
    hwnd_mag = user32.CreateWindowExW(
        0,
        "Magnifier", "MagnifierChild",
        WS_CHILD | WS_VISIBLE | MS_SHOWMAGNIFIEDCURSOR,
        0, 0, 1, 1,
        hwnd_host, None, hinstance, None,
    )
    if not hwnd_mag:
        print(f"ERROR: magnifier CreateWindowExW failed (err {kernel32.GetLastError()})")
        sys.exit(1)

    # Apply grayscale colour effect
    if not mag.MagSetColorEffect(hwnd_mag, ctypes.byref(GRAYSCALE)):
        print("WARNING: MagSetColorEffect failed — colours may not change")

    # 1 : 1 scale (no magnification)
    mag.MagSetWindowTransform(hwnd_mag, ctypes.byref(IDENTITY_TRANSFORM))

    print("Grayscale Apps running (per-window overlay)")
    print(f"  Targets : {', '.join(sorted(targets))}")
    print(f"  Interval: {int(interval * 1000)}ms")
    print("  Press Ctrl+C to stop\n")

    overlay_visible = False
    prev_target = None
    config_mtime = os.path.getmtime(CONFIG_PATH)

    def cleanup():
        nonlocal overlay_visible
        if overlay_visible:
            user32.ShowWindow(hwnd_host, SW_HIDE)
            overlay_visible = False
        user32.DestroyWindow(hwnd_host)
        mag.MagUninitialize()

    atexit.register(cleanup)

    try:
        while True:
            # ── pump messages ──
            msg = wt.MSG()
            while user32.PeekMessageW(ctypes.byref(msg), None, 0, 0, PM_REMOVE):
                if msg.message == 0x0012:  # WM_QUIT
                    return
                user32.TranslateMessage(ctypes.byref(msg))
                user32.DispatchMessageW(ctypes.byref(msg))

            # ── hot-reload config ──
            try:
                mt = os.path.getmtime(CONFIG_PATH)
                if mt != config_mtime:
                    targets, interval = load_config()
                    config_mtime = mt
                    print(f"  Config reloaded — targets: {', '.join(sorted(targets))}")
            except Exception:
                pass

            # ── find target ──
            hwnd_target = find_target_window(targets)

            if hwnd_target:
                # Update filter list when the target hwnd changes
                if hwnd_target != prev_target:
                    arr = (wt.HWND * 1)(hwnd_target)
                    mag.MagSetWindowFilterList(hwnd_mag, MW_FILTERMODE_INCLUDE, 1, arr)
                    prev_target = hwnd_target
                    print(f"  ● tracking window {hwnd_target}")

                rect = wt.RECT()
                user32.GetWindowRect(hwnd_target, ctypes.byref(rect))
                x, y = rect.left, rect.top
                w, h = rect.right - rect.left, rect.bottom - rect.top

                # Position & size the host over the target
                user32.SetWindowPos(
                    hwnd_host, HWND_TOPMOST,
                    x, y, w, h,
                    SWP_NOACTIVATE,
                )

                # Resize magnifier child to fill host
                user32.SetWindowPos(
                    hwnd_mag, None,
                    0, 0, w, h,
                    SWP_NOACTIVATE,
                )

                # Tell magnifier which screen area to capture
                source = wt.RECT(rect.left, rect.top, rect.right, rect.bottom)
                mag.MagSetWindowSource(hwnd_mag, source)
                user32.InvalidateRect(hwnd_mag, None, True)

                if not overlay_visible:
                    user32.ShowWindow(hwnd_host, SW_SHOWNOACTIVATE)
                    overlay_visible = True
                    print(f"  ● overlay ON")
            else:
                if overlay_visible:
                    user32.ShowWindow(hwnd_host, SW_HIDE)
                    overlay_visible = False
                    prev_target = None
                    print(f"  ○ overlay OFF (target not found)")

            time.sleep(interval)

    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
