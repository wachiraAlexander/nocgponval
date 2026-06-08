import subprocess
import sys
import os
import time
import signal

APP_DIR = os.path.dirname(os.path.abspath(__file__))
PY_DIR = APP_DIR

PROCESS = None
MTIMES = {}

def get_watched_files():
    files = {}
    for root, dirs, fnames in os.walk(PY_DIR):
        dirs[:] = [d for d in dirs if not d.startswith(".") and d not in ("__pycache__", "node_modules", "uploads", ".git", ".sixth")]
        for f in fnames:
            if f.endswith(".py") or f.endswith(".html") or f.endswith(".js"):
                p = os.path.join(root, f)
                try:
                    files[p] = os.path.getmtime(p)
                except OSError:
                    pass
    return files

def get_py_files():
    return get_watched_files()

def start_process():
    global PROCESS
    cmd = [sys.executable, "-m", "waitress", "--port=5000", "--host=0.0.0.0", "wsgi:application"]
    print(f"[autoreload] Starting: {' '.join(cmd)}")
    PROCESS = subprocess.Popen(cmd, cwd=APP_DIR)

def kill_process():
    global PROCESS
    if PROCESS and PROCESS.poll() is None:
        try:
            PROCESS.send_signal(signal.SIGTERM)
        except OSError:
            PROCESS.kill()
        PROCESS.wait()
    PROCESS = None

if __name__ == "__main__":
    MTIMES = get_watched_files()
    os.chdir(APP_DIR)
    start_process()
    print(f"[autoreload] Watching {len(MTIMES)} Python/JS/HTML files for changes...")
    try:
        while True:
            time.sleep(2)
            current = get_py_files()
            if current != MTIMES:
                changed = [os.path.relpath(k, APP_DIR) for k in current if k not in MTIMES or current[k] != MTIMES[k]]
                print(f"[autoreload] Changed: {', '.join(changed)} — restarting...")
                MTIMES = current
                kill_process()
                start_process()
            if PROCESS and PROCESS.poll() is not None:
                print(f"[autoreload] Crashed (code {PROCESS.returncode}), restarting...")
                start_process()
    except KeyboardInterrupt:
        print("\n[autoreload] Shutting down...")
        kill_process()
