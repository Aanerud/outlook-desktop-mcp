"""Platform-aware entry point for outlook-desktop-mcp.

Detects the OS and imports the correct server module:
- macOS → server_mac (AppleScript automation)
- Windows → server (COM automation)
"""
import sys


def main():
    if sys.platform == "darwin":
        from outlook_desktop_mcp.server_mac import main as _main
    else:
        from outlook_desktop_mcp.server import main as _main
    _main()
