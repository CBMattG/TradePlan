"""Tradeplan desktop app — opens the plan in its own window (no browser, no console).

Double-click this file (or the Tradeplan desktop / Start-menu shortcut). It:
  1. runs the same `server.py` the browser version uses, on a background thread
     (http://127.0.0.1:8765 — the window talks to it exactly like the browser did);
  2. opens a native window (pywebview → Microsoft Edge WebView2) on it;
  3. on close, flushes any pending autosave before the window goes, then exits — which
     stops the server with it.
Only one copy runs at a time: launching it again brings the open window to the front.
If `run.bat` already has a Tradeplan server running on :8765, the window just uses it.

Needs: pip install pywebview   (the WebView2 runtime ships with Windows 11)
Log:   %LOCALAPPDATA%\\Tradeplan\\tradeplan.log  (no console under pythonw)
"""
import ctypes
import os
import sys
import threading
import time
import traceback
import urllib.request
from http.server import ThreadingHTTPServer
from pathlib import Path

APP_DIR = Path(__file__).resolve().parent
TITLE = "Tradeplan"
APP_ID = "CharlesBentley.Tradeplan"          # groups the taskbar button under our icon, not Python's
ICON = APP_DIR / "tradeplan.ico"
HOME = Path(os.environ.get("LOCALAPPDATA", str(Path.home()))) / "Tradeplan"
HOME.mkdir(parents=True, exist_ok=True)

# pythonw has no console: send print/tracebacks to a log file instead
if sys.stdout is None or sys.stderr is None or "pythonw" in Path(sys.executable).name.lower():
    _log = open(HOME / "tradeplan.log", "a", encoding="utf-8", buffering=1)
    sys.stdout = sys.stderr = _log
print(f"\n=== Tradeplan start {time.strftime('%Y-%m-%d %H:%M:%S')} ===")

user32 = ctypes.windll.user32


def message(text, title=TITLE, error=True):
    user32.MessageBoxW(None, text, title, 0x10 if error else 0x40)


# ---- single instance: a second launch just brings the open window forward ----
_mutex = ctypes.windll.kernel32.CreateMutexW(None, False, "Local\\TradeplanDesktopApp")
if ctypes.windll.kernel32.GetLastError() == 183:          # ERROR_ALREADY_EXISTS
    hwnd = user32.FindWindowW(None, TITLE)
    if hwnd:
        user32.ShowWindow(hwnd, 9)                          # SW_RESTORE
        user32.SetForegroundWindow(hwnd)
    sys.exit(0)

try:
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(APP_ID)
except Exception:
    pass

try:
    import webview
except ImportError:
    message("Tradeplan needs the 'pywebview' package.\n\nOpen a command prompt and run:\n\n    pip install pywebview\n\n"
            "(or use run.bat to open Tradeplan in the browser meanwhile).")
    sys.exit(1)

sys.path.insert(0, str(APP_DIR))
os.chdir(APP_DIR)
import server  # noqa: E402  (the app's own server: same data/, same routes)

server.PORT = int(os.environ.get("TRADEPLAN_PORT") or server.PORT)   # override only for testing a copy

URL = f"http://127.0.0.1:{server.PORT}/"


def tradeplan_already_serving():
    try:
        with urllib.request.urlopen(URL, timeout=2) as r:
            return b"Tradeplan" in r.read(4000)
    except Exception:
        return False


httpd = None
try:
    httpd = ThreadingHTTPServer(("127.0.0.1", server.PORT), server.Handler)
    threading.Thread(target=httpd.serve_forever, name="tradeplan-server", daemon=True).start()
    print("server on", URL)
except OSError:
    if tradeplan_already_serving():
        print("port busy — using the Tradeplan server already running on", URL)
    else:
        message(f"Port {server.PORT} is already in use by another program, so Tradeplan can't start.\n\n"
                "Close whatever is using it (or restart the PC) and try again.")
        sys.exit(1)

# ---- closing: let the page finish its autosave first ----
_closing_ok = False


def _flush_then_close(window):
    """Background: ask the page to save anything pending, wait for it, then close for real."""
    global _closing_ok
    status = ""
    try:
        window.evaluate_js("try { savePlanPlace(); } catch (e) {} if (saveTimer) saveNow(); 1")
        for _ in range(100):                                # up to ~15 s
            status = window.evaluate_js(
                "(saveTimer ? 'pending' : '') + document.getElementById('save-status').textContent") or ""
            if not status.startswith("pending") and "Saving" not in status:
                break
            time.sleep(0.15)
    except Exception:
        traceback.print_exc()
    if "fail" in status.lower() and not window.create_confirmation_dialog(
            TITLE, "Your latest changes could not be saved.\n\nClose Tradeplan anyway?"):
        return
    _closing_ok = True
    window.destroy()


def on_closing(window):
    if _closing_ok:
        return True
    threading.Thread(target=_flush_then_close, args=(window,), daemon=True).start()
    return False                                            # cancel this close; the flush closes it


webview.settings["ALLOW_DOWNLOADS"] = True                  # exports open a "Save as" dialog
webview.settings["OPEN_EXTERNAL_LINKS_IN_BROWSER"] = True
# pywebview only allows the browser keys (F5 / Ctrl+R reload, Ctrl+F find, Ctrl +/- zoom) and
# the right-click Copy/Paste menu in its "debug" mode — so run in it, but keep the developer
# tools closed (F12 opens them) and the log at its normal level.
webview.settings["OPEN_DEVTOOLS_IN_DEBUG"] = False
os.environ.setdefault("PYWEBVIEW_LOG", "warning")

win = webview.create_window(TITLE, URL, width=1500, height=900, min_size=(1000, 650), maximized=True,
                            text_select=True, zoomable=True, background_color="#f4f6fa")
win.events.closing += on_closing

try:
    # private_mode off + a fixed storage folder: keeps the app's remembered view (last tab,
    # supplier, search, theme…) between sessions, as the browser did
    webview.start(private_mode=False, storage_path=str(HOME / "webview"),
                  icon=str(ICON) if ICON.exists() else None, debug=True)
except Exception:
    traceback.print_exc()
    message("Tradeplan's window couldn't open (see %LOCALAPPDATA%\\Tradeplan\\tradeplan.log).\n\n"
            "run.bat still opens it in the browser.")
finally:
    if httpd:
        httpd.shutdown()
    print("=== Tradeplan closed ===")
