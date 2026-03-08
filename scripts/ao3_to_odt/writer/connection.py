import os
import sys
import socket
import subprocess
import time
from pathlib import Path

import uno

# Locates LO executable
def find_soffice():
    import shutil, glob
    candidates = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        "/usr/lib/libreoffice/program/soffice",
        "/usr/local/bin/soffice",
    ]
    for pattern in [r"C:\Program Files\LibreOffice*\program\soffice.exe"]:
        found = glob.glob(pattern)
        candidates.extend(sorted(found))
    for c in candidates:
        if Path(c).exists(): return c
    which = shutil.which("soffice")
    if which: return which
    return None

# Polling for LO startup
def is_port_open(port):
    try:
        with socket.create_connection(("localhost", port), timeout=1): return True
    except: return False

# clean up after crash
def clear_lo_locks():
    import glob
    profile_dirs = []
    if sys.platform == "win32":
        appdata = os.environ.get("APPDATA", "")
        if appdata:
            profile_dirs.append(os.path.join(appdata, "LibreOffice", "4", "user"))
    else:
        home = os.path.expanduser("~")
        profile_dirs.append(os.path.join(home, ".config", "libreoffice", "4", "user"))

    removed = []
    for profile in profile_dirs:
        for lock_pattern in [".~lock.*", "*.lock"]:
            for f in glob.glob(os.path.join(profile, lock_pattern)):
                try:
                    os.remove(f)
                    removed.append(f)
                except OSError:
                    pass
    if removed:
        print(f"  Cleared {len(removed)} stale lock file(s)")

# laund LO as server
def start_lo_listener(soffice_path, port=2002):
    clear_lo_locks()
    accept = f"socket,host=localhost,port={port};urp;StarOffice.ServiceManager"
    tmp_profile = os.path.join(os.environ.get("TEMP", os.path.expanduser("~")), "lo_ao3_profile")
    os.makedirs(tmp_profile, exist_ok=True)
    user_install = f"-env:UserInstallation={uno.systemPathToFileUrl(tmp_profile)}"

    cmd = [
        soffice_path,
        "--headless",
        "--norestore",
        "--nofirststartwizard",
        "--nologo",
        user_install,
        f"--accept={accept}",
    ]
    print(f"  Profile: {tmp_profile}")

    kwargs = {}
    if sys.platform == "win32":
        kwargs["creationflags"] = subprocess.CREATE_NO_WINDOW
    proc = subprocess.Popen(
        cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, **kwargs)
    return proc

def connect_uno(port=2002, retries=12, delay=3):
    localContext = uno.getComponentContext()
    resolver = localContext.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", localContext)
    url = f"uno:socket,host=localhost,port={port};urp;StarOffice.ComponentContext"
    last_err = None
    for attempt in range(1, retries + 1):
        try:
            print(f"  Connecting to LO UNO (attempt {attempt}/{retries})...", end="", flush=True)
            ctx = resolver.resolve(url)
            desktop = ctx.ServiceManager.createInstanceWithContext(
                "com.sun.star.frame.Desktop", ctx)
            print(" ok")
            return desktop
        except Exception as e:
            last_err = e
            print(f" waiting {delay}s...")
            time.sleep(delay)
    raise ConnectionError(f"Could not connect after {retries} attempts: {last_err}")