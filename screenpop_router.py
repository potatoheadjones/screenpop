import argparse
import json
import os
import platform
import re
import subprocess
import threading
import time
from pathlib import Path
from queue import Queue, Full
from urllib.parse import urlparse, unquote

import requests
from flask import Flask, request, jsonify, Response

from PIL import Image, ImageDraw
import pystray

# Optional UI prompts (custom size / seconds)
try:
    import tkinter as tk
    from tkinter import simpledialog, messagebox
    TK_AVAILABLE = True
except Exception:
    TK_AVAILABLE = False

# Optional Windows focus control (best-effort)
ON_WINDOWS = platform.system() == "Windows"
if ON_WINDOWS:
    try:
        import win32gui, win32con, win32process
        HAVE_WIN = True
    except Exception:
        HAVE_WIN = False
else:
    HAVE_WIN = False

# ---------------- App / Config ----------------
app = Flask(__name__)

CONFIG_LOCK = threading.Lock()
CONFIG = {
    "browser": "edge",             # auto|chrome|edge|system  (set this to NOT be your Genesys browser)
    "mode": "first-window-then-tabs",  # new-tab|new-window|first-window-then-tabs
    "fullscreen": False,           # applies to first new-window
    "size": [1400, 900],           # applies to first new-window
    "allowlist": [],               # e.g., ["crm.example.com"]
    "queue_max": 128,
    "dedupe_window_s": 10,         # 0 = off
    "separate_instance": True,     # force dedicated user-data-dir for pop browser
    "app_window": False,           # open first window as --app=<url> (chromeless). Sizing may vary by OS
    "win_no_activate": True,       # Windows only: try to avoid focus steal
}
CONFIG_PATH = Path.cwd() / "screenpop_tray.config.json"

STATE_LOCK = threading.Lock()
STATE = {"first_window_done": False}

JOBQ = Queue(maxsize=CONFIG["queue_max"])
STATS = {
    "enqueued": 0,
    "processed": 0,
    "failed": 0,
    "suppressed": 0,
    "last_error": ""
}

# Dedupe store
_DEDUPE_LOCK = threading.Lock()
_LAST_POP = {}
_PRUNE_EVERY = 200
_opcount = 0

def cfg_get(key=None):
    with CONFIG_LOCK:
        return CONFIG if key is None else CONFIG.get(key)

def cfg_set(key, value):
    with CONFIG_LOCK:
        CONFIG[key] = value
        try:
            CONFIG_PATH.write_text(json.dumps(CONFIG, indent=2), encoding="utf-8")
        except Exception:
            pass

def cfg_update(d):
    with CONFIG_LOCK:
        CONFIG.update(d)
        try:
            CONFIG_PATH.write_text(json.dumps(CONFIG, indent=2), encoding="utf-8")
        except Exception:
            pass

def state_get(key):
    with STATE_LOCK:
        return STATE.get(key)

def state_set(key, value):
    with STATE_LOCK:
        STATE[key] = value

# ---------------- Utilities ----------------
def parse_size(size_str):
    m = re.match(r"^\s*(\d+)\s*[xX]\s*(\d+)\s*$", size_str or "")
    if not m:
        raise ValueError("Invalid size. Use WIDTHxHEIGHT (e.g., 1280x800)")
    return [int(m.group(1)), int(m.group(2))]

def which_exe(candidates):
    for c in candidates:
        if Path(c).exists():
            return c
    import shutil
    for name in candidates:
        exe = shutil.which(Path(name).name)
        if exe:
            return exe
    return None

def chrome_path():
    system = platform.system()
    if system == "Windows":
        return which_exe([
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
            "chrome.exe"
        ])
    elif system == "Darwin":
        return which_exe([
            "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
            "google-chrome"
        ])
    else:
        return which_exe(["google-chrome", "chromium", "chromium-browser"])

def edge_path():
    system = platform.system()
    if system == "Windows":
        return which_exe([
            r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
            r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
            "msedge.exe"
        ])
    elif system == "Darwin":
        return which_exe([
            "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
            "msedge"
        ])
    else:
        return which_exe(["microsoft-edge", "msedge", "edge"])

def resolve_browser_exe():
    b = cfg_get("browser")
    if b == "chrome":
        return chrome_path()
    if b == "edge":
        return edge_path()
    if b == "auto":
        return chrome_path() or edge_path()
    return None  # system

def allowed_host(url):
    allow = cfg_get("allowlist")
    if not allow:
        return True
    try:
        host = urlparse(url).hostname or ""
        return any(host.endswith(a) for a in allow)
    except Exception:
        return False

def open_system_new_tab(url: str):
    import webbrowser
    webbrowser.open_new_tab(url)

# ---------------- Windows focus helpers ----------------
def _win_show_no_activate_for_pid(pid):
    """Best-effort: show window without activation for a process id."""
    if not HAVE_WIN:
        return
    hwnd_targets = []

    def enum_proc(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            try:
                _, wpid = win32process.GetWindowThreadProcessId(hwnd)
                if wpid == pid:
                    hwnd_targets.append(hwnd)
            except Exception:
                pass
        return True

    win32gui.EnumWindows(enum_proc, None)
    for hwnd in hwnd_targets[:3]:  # first few
        try:
            # Show without activating
            win32gui.ShowWindow(hwnd, win32con.SW_SHOWNOACTIVATE)
        except Exception:
            pass

# ---------------- Launchers ----------------
def _user_data_flag():
    if not cfg_get("separate_instance"):
        return None
    p = Path.cwd() / ".screenpop_profile"
    p.mkdir(exist_ok=True)
    return f"--user-data-dir={p}"

def launch_new_tab(url: str):
    exe = resolve_browser_exe()
    if not exe or cfg_get("browser") == "system":
        open_system_new_tab(url)
        return

    flags = ["--new-tab"]
    udir = _user_data_flag()
    if udir:
        flags.append(udir)

    try:
        subprocess.Popen([exe, *flags, url],
                         stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        # New tab goes to existing window; we won't fight focus here except optionally on Windows:
        if ON_WINDOWS and cfg_get("win_no_activate"):
            # Can't easily stop activation on tab-add; skip or add gentle no-activate on owning pid
            pass
    except Exception:
        open_system_new_tab(url)

def launch_new_window(url: str):
    exe = resolve_browser_exe()
    if not exe or cfg_get("browser") == "system":
        open_system_new_tab(url)
        return

    udir = _user_data_flag()
    app_mode = cfg_get("app_window")
    fullscreen = cfg_get("fullscreen")
    size = cfg_get("size")

    # Two styles: normal window OR app window
    flags = ["--disable-first-run-ui", "--no-default-browser-check"]

    if app_mode:
        # app windows ignore some sizing in some OS builds; still try
        flags.append(f"--app={url}")
        if fullscreen:
            flags.append("--start-fullscreen")
        elif size and len(size) == 2:
            w, h = size
            flags.append(f"--window-size={w},{h}")
        if udir:
            flags.append(udir)
        proc = subprocess.Popen([exe, *flags],
                                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    else:
        flags.append("--new-window")
        if fullscreen:
            flags.append("--start-fullscreen")
        elif size and len(size) == 2:
            w, h = size
            flags.append(f"--window-size={w},{h}")
        if udir:
            flags.append(udir)
        proc = subprocess.Popen([exe, *flags, url],
                                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

    if ON_WINDOWS and cfg_get("win_no_activate"):
        # Give the window a moment to appear, then show without activation
        try:
            time.sleep(0.25)
            _win_show_no_activate_for_pid(proc.pid)
        except Exception:
            pass

# ---------------- Dedupe logic ----------------
def _prune_dedupe(now_ts: float, window_s: float):
    cutoff = now_ts - max(window_s, 1.0) * 10.0
    stale = [u for u, ts in _LAST_POP.items() if ts < cutoff]
    for u in stale:
        _LAST_POP.pop(u, None)

def should_suppress(url: str) -> bool:
    window = float(cfg_get("dedupe_window_s") or 0)
    if window <= 0:
        return False
    now_ts = time.time()
    global _opcount
    with _DEDUPE_LOCK:
        last = _LAST_POP.get(url)
        if last is not None and (now_ts - last) < window:
            return True
        _LAST_POP[url] = now_ts
        _opcount += 1
        if _opcount % _PRUNE_EVERY == 0:
            _prune_dedupe(now_ts, window)
    return False

# ---------------- Worker ----------------
def process_job(job):
    try:
        url = job["url"]
        mode = cfg_get("mode")
        if mode == "new-window":
            launch_new_window(url)
        elif mode == "first-window-then-tabs":
            if not state_get("first_window_done"):
                launch_new_window(url)
                state_set("first_window_done", True)
            else:
                launch_new_tab(url)
        else:
            launch_new_tab(url)
        STATS["processed"] += 1
    except Exception as ex:
        STATS["failed"] += 1
        STATS["last_error"] = f"{type(ex).__name__}: {ex}"

def worker_loop():
    while True:
        job = JOBQ.get()
        try:
            process_job(job)
        finally:
            JOBQ.task_done()

threading.Thread(target=worker_loop, name="screenpop-worker", daemon=True).start()

# ---------------- HTTP endpoints ----------------
@app.route("/")
def health():
    b = cfg_get("browser")
    m = cfg_get("mode")
    fs = cfg_get("fullscreen")
    sz = cfg_get("size")
    dedupe = cfg_get("dedupe_window_s")
    first_done = state_get("first_window_done")
    html = f"""
    <html>
      <body style="font-family:system-ui">
        <h3>Screen-pop router running</h3>
        <p><b>Browser:</b> {b} &nbsp; <b>Mode:</b> {m}</p>
        <p><b>Fullscreen:</b> {fs} &nbsp; <b>Size:</b> {sz[0]}x{sz[1]}</p>
        <p><b>Separate instance:</b> {cfg_get('separate_instance')} &nbsp; <b>App window:</b> {cfg_get('app_window')}</p>
        <p><b>Windows no-activate:</b> {cfg_get('win_no_activate')}</p>
        <p><b>Deduplicate window:</b> {dedupe} s (0 = off)</p>
        <p><b>First window done:</b> {first_done}</p>
        <p>GET /open?u=...</p>
      </body>
    </html>
    """
    return Response(html, mimetype="text/html")

@app.route("/stats")
def stats():
    out = dict(STATS)
    out["queue_size"] = JOBQ.qsize()
    out["dedupe_window_s"] = cfg_get("dedupe_window_s")
    out["mode"] = cfg_get("mode")
    out["first_window_done"] = state_get("first_window_done")
    return jsonify(out)

@app.route("/open")
def open_url():
    """
    /open?u=<URL-ENCODED-TARGET>
    Dedupe per exact URL; returns 202.
    """
    raw = request.args.get("u", "").strip()
    if not raw:
        return jsonify({"ok": False, "error": "Missing query parameter u"}), 400

    try:
        target_url = unquote(raw)
        if "%" in target_url:
            target_url = unquote(target_url)
    except Exception:
        target_url = raw

    if not (target_url.startswith("http://") or target_url.startswith("https://")):
        return jsonify({"ok": False, "error": "u must be an absolute http(s) URL"}), 400
    if not allowed_host(target_url):
        return jsonify({"ok": False, "error": "Host not allowed by allowlist"}), 403

    if should_suppress(target_url):
        STATS["suppressed"] += 1
        return jsonify({"ok": True, "status": "suppressed", "target": target_url}), 202

    job = {"url": target_url}
    try:
        JOBQ.put_nowait(job)
        STATS["enqueued"] += 1
    except Full:
        return jsonify({"ok": False, "error": "Queue full. Try again shortly."}), 429

    return jsonify({
        "ok": True,
        "status": "queued",
        "target": target_url,
        "mode": cfg_get("mode"),
        "first_window_done": state_get("first_window_done")
    }), 202

# ---------------- Tray UI ----------------
def make_icon_image():
    img = Image.new('RGBA', (64, 64), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.rounded_rectangle([6, 6, 58, 58], radius=12, fill=(30, 136, 229, 255))
    d.line([20, 40, 44, 16], width=5, fill=(255, 255, 255, 255))
    d.polygon([(44, 16), (44, 28), (32, 16)], fill=(255, 255, 255, 255))
    return img

def tray_set_browser(value):
    def inner(icon, item): cfg_set("browser", value); icon.update_menu()
    return inner

def tray_set_mode(value):
    def inner(icon, item): cfg_set("mode", value); icon.update_menu()
    return inner

def tray_reset_first(icon, item):
    state_set("first_window_done", False); icon.update_menu()

def tray_toggle_fullscreen(icon, item):
    cfg_set("fullscreen", not cfg_get("fullscreen")); icon.update_menu()

def tray_toggle_sep_instance(icon, item):
    cfg_set("separate_instance", not cfg_get("separate_instance")); icon.update_menu()

def tray_toggle_app_window(icon, item):
    cfg_set("app_window", not cfg_get("app_window")); icon.update_menu()

def tray_toggle_win_no_act(icon, item):
    if not ON_WINDOWS:
        return
    cfg_set("win_no_activate", not cfg_get("win_no_activate")); icon.update_menu()

def tray_set_size_preset(w, h):
    def inner(icon, item):
        cfg_update({"size": [w, h], "fullscreen": False})
        icon.update_menu()
    return inner

def tray_set_size_custom(icon, item):
    if not TK_AVAILABLE:
        cfg_update({"size": [1280, 800], "fullscreen": False}); return
    root = tk.Tk(); root.withdraw()
    try:
        val = simpledialog.askstring("Custom size", "Enter size as WIDTHxHEIGHT (e.g., 1600x900):")
        if val:
            try:
                size = parse_size(val)
                cfg_update({"size": size, "fullscreen": False})
            except ValueError as e:
                messagebox.showerror("Invalid size", str(e))
    finally:
        root.destroy()

def tray_open_config(icon, item):
    folder = str(Path.cwd())
    if platform.system() == "Windows":
        subprocess.Popen(["explorer", folder])
    elif platform.system() == "Darwin":
        subprocess.Popen(["open", folder])
    else:
        subprocess.Popen(["xdg-open", folder])

def tray_set_dedupe_seconds(value):
    def inner(icon, item): cfg_set("dedupe_window_s", int(value)); icon.update_menu()
    return inner

def tray_set_dedupe_custom(icon, item):
    if not TK_AVAILABLE:
        cfg_set("dedupe_window_s", 10); return
    root = tk.Tk(); root.withdraw()
    try:
        val = simpledialog.askstring("Deduplicate window", "Seconds (0 = off):")
        if val is not None and val != "":
            try:
                s = int(val)
                if s < 0: raise ValueError("Seconds must be >= 0")
                cfg_set("dedupe_window_s", s)
            except ValueError as e:
                messagebox.showerror("Invalid seconds", str(e))
    finally:
        root.destroy()

def build_menu():
    sz = cfg_get("size")
    items = [
        pystray.MenuItem("Screen-pop Router", lambda: None, enabled=False),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem(
            "Browser",
            pystray.Menu(
                pystray.MenuItem("Auto", tray_set_browser("auto"), checked=lambda item: cfg_get("browser")=="auto"),
                pystray.MenuItem("Chrome", tray_set_browser("chrome"), checked=lambda item: cfg_get("browser")=="chrome"),
                pystray.MenuItem("Edge", tray_set_browser("edge"), checked=lambda item: cfg_get("browser")=="edge"),
                pystray.MenuItem("System default", tray_set_browser("system"), checked=lambda item: cfg_get("browser")=="system"),
            )
        ),
        pystray.MenuItem(
            "Open as",
            pystray.Menu(
                pystray.MenuItem("New tab", tray_set_mode("new-tab"), checked=lambda item: cfg_get("mode")=="new-tab"),
                pystray.MenuItem("New window", tray_set_mode("new-window"), checked=lambda item: cfg_get("mode")=="new-window"),
                pystray.MenuItem("First window, then tabs",
                                 tray_set_mode("first-window-then-tabs"),
                                 checked=lambda item: cfg_get("mode")=="first-window-then-tabs"),
            )
        ),
        pystray.MenuItem("Reset 'first pop' state", tray_reset_first),
        pystray.MenuItem("Fullscreen on first window", tray_toggle_fullscreen, checked=lambda i: cfg_get("fullscreen")),
        pystray.MenuItem(f"First window size (current {sz[0]}x{sz[1]})",
            pystray.Menu(
                pystray.MenuItem("1280 x 800", tray_set_size_preset(1280, 800)),
                pystray.MenuItem("1400 x 900", tray_set_size_preset(1400, 900)),
                pystray.MenuItem("1600 x 900", tray_set_size_preset(1600, 900)),
                pystray.MenuItem("1920 x 1080", tray_set_size_preset(1920, 1080)),
                pystray.MenuItem("Custom…", tray_set_size_custom),
            )
        ),
        pystray.MenuItem("Separate instance (own profile dir)", tray_toggle_sep_instance,
                         checked=lambda i: cfg_get("separate_instance")),
        pystray.MenuItem("App window (chromeless first window)", tray_toggle_app_window,
                         checked=lambda i: cfg_get("app_window")),
        pystray.MenuItem("Windows: no-activate (best-effort)", tray_toggle_win_no_act,
                         checked=lambda i: cfg_get("win_no_activate"), default=False, enabled=ON_WINDOWS and HAVE_WIN),
        pystray.MenuItem(
            "Deduplicate interval",
            pystray.Menu(
                pystray.MenuItem("Off", tray_set_dedupe_seconds(0), checked=lambda i: cfg_get("dedupe_window_s")==0),
                pystray.MenuItem("5 seconds", tray_set_dedupe_seconds(5), checked=lambda i: cfg_get("dedupe_window_s")==5),
                pystray.MenuItem("10 seconds", tray_set_dedupe_seconds(10), checked=lambda i: cfg_get("dedupe_window_s")==10),
                pystray.MenuItem("30 seconds", tray_set_dedupe_seconds(30), checked=lambda i: cfg_get("dedupe_window_s")==30),
                pystray.MenuItem("60 seconds", tray_set_dedupe_seconds(60), checked=lambda i: cfg_get("dedupe_window_s")==60),
                pystray.MenuItem("Custom…", tray_set_dedupe_custom),
            )
        ),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("Open config folder", tray_open_config),
        pystray.MenuItem("Quit", lambda icon, item: icon.stop())
    ]
    return pystray.Menu(*items)

def tray_open_config(icon, item):
    folder = str(Path.cwd())
    if platform.system() == "Windows":
        subprocess.Popen(["explorer", folder])
    elif platform.system() == "Darwin":
        subprocess.Popen(["open", folder])
    else:
        subprocess.Popen(["xdg-open", folder])

def tray_thread(port: int):
    if CONFIG_PATH.exists():
        try:
            obj = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
            cfg_update(obj)
        except Exception:
            pass
    icon = pystray.Icon("screenpop_router", make_icon_image(), f"Screen-pop @ {port}", build_menu())
    icon.run()

# ---------------- Boot ----------------
def run_server(port: int, threads: int):
    from waitress import serve
    print(f"[screen-pop] http://127.0.0.1:{port}  browser={cfg_get('browser')}  mode={cfg_get('mode')}  dedupe={cfg_get('dedupe_window_s')}s")
    serve(app, host="127.0.0.1", port=port, threads=threads, backlog=256, channel_timeout=30)

def main():
    parser = argparse.ArgumentParser(description="Genesys-friendly screen-pop: first window then tabs, other browser")
    parser.add_argument("--port", type=int, default=5588)
    parser.add_argument("--threads", type=int, default=8, help="Waitress threads")
    args = parser.parse_args()

    t = threading.Thread(target=run_server, args=(args.port, args.threads), daemon=True)
    t.start()
    tray_thread(args.port)

if __name__ == "__main__":
    main()
