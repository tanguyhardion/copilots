"""
Main pywebview Application Entrypoint for the Unified Copilot Suite.
"""

import os
import sys
import webview

# Add current workspace to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from copilots_app.web_api import CopilotBridge


def get_html_entrypoint() -> str:
    """Resolve absolute path to web/index.html across dev and PyInstaller builds."""
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        base_dir = os.path.join(sys._MEIPASS, "copilots_app", "web")
    else:
        base_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "copilots_app", "web")
    return os.path.join(base_dir, "index.html")


def main():
    bridge = CopilotBridge()
    html_path = get_html_entrypoint()

    # Create native WebView window (Edge Chromium WebView2 on Windows)
    window = webview.create_window(
        title="Copilots — Unified Suite",
        url=f"file:///{html_path.replace(os.sep, '/')}",
        js_api=bridge,
        width=1280,
        height=820,
        min_size=(1000, 680),
        background_color="#0B0D13",
    )

    # Attach window reference for file dialogs
    bridge.set_window(window)

    # Start pywebview event loop (maximize on load)
    webview.start(debug=False, func=lambda: window.maximize())


if __name__ == "__main__":
    main()
