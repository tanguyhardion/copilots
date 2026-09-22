"""
Main pywebview Application Entrypoint for the Unified Copilot Suite.
"""

import os
import sys
import ctypes
import webview

# Add current workspace to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from copilots_app.web_api import CopilotBridge


def get_asset_path(relative_path: str) -> str:
    """Resolve absolute path across dev and PyInstaller builds."""
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        base_dir = sys._MEIPASS
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_dir, relative_path)


def get_html_entrypoint() -> str:
    """Resolve absolute path to web/index.html across dev and PyInstaller builds."""
    return get_asset_path(os.path.join("copilots_app", "web", "index.html"))


def set_app_icon(window):
    """Set window and taskbar icon natively on Windows."""
    try:
        ico_path = get_asset_path(os.path.join("assets", "icons", "copilots.ico"))
        if not os.path.exists(ico_path) or sys.platform != "win32":
            return

        # Explicitly set AppUserModelID so Windows taskbar groups and shows the custom icon
        app_id = "copilots.suite.app.1.0"
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(app_id)

        # Apply icon to native Win32 window handles created by pywebview/WebView2
        user32 = ctypes.windll.user32
        IMAGE_ICON = 1
        LR_LOADFROMFILE = 0x00000010
        LR_DEFAULTSIZE = 0x00000040
        WM_SETICON = 0x0080
        ICON_SMALL = 0
        ICON_BIG = 1

        hicon_big = user32.LoadImageW(None, ico_path, IMAGE_ICON, 0, 0, LR_LOADFROMFILE | LR_DEFAULTSIZE)
        hicon_small = user32.LoadImageW(None, ico_path, IMAGE_ICON, 16, 16, LR_LOADFROMFILE)

        # Enumerate top-level windows for this process and set the icon
        current_pid = os.getpid()

        WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)

        def enum_windows_callback(hwnd, lparam):
            window_pid = ctypes.c_ulong()
            user32.GetWindowThreadProcessId(hwnd, ctypes.byref(window_pid))
            if window_pid.value == current_pid and user32.IsWindowVisible(hwnd):
                if hicon_big:
                    user32.SendMessageW(hwnd, WM_SETICON, ICON_BIG, hicon_big)
                if hicon_small:
                    user32.SendMessageW(hwnd, WM_SETICON, ICON_SMALL, hicon_small)
            return True

        user32.EnumWindows(WNDENUMPROC(enum_windows_callback), 0)
    except Exception as e:
        print(f"[Icon] Could not set Windows native icon: {e}")


def on_startup(window):
    set_app_icon(window)
    window.maximize()


def main():
    bridge = CopilotBridge()
    html_path = get_html_entrypoint()

    # Create native WebView window (Edge Chromium WebView2 on Windows)
    window = webview.create_window(
        title="Copilots",
        url=f"file:///{html_path.replace(os.sep, '/')}",
        js_api=bridge,
        width=1280,
        height=820,
        min_size=(1000, 680),
        background_color="#0B0D13",
    )

    # Attach window reference for file dialogs
    bridge.set_window(window)

    # Start pywebview event loop (set icon and maximize on load)
    webview.start(debug=False, func=lambda: on_startup(window))


if __name__ == "__main__":
    main()
