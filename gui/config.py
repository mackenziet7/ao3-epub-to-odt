import json
from pathlib import Path

APP_NAME = "AO3toODT"
DEFAULT_LO_PYTHON = Path(r"C:\Program Files\LibreOffice\program\python.exe")

def config_path() -> Path:
    return Path.home() / "AppData" / "Roaming" / APP_NAME / "config.json"

def load_config() -> dict:
    p = config_path()
    if p.exists():
        try: 
            return json.loads(p.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}

def save_config(data: dict):
    p = config_path()
    p.parent.mkdir (parents=True, exist_ok=True)
    p.write_text(json.dumps(data, indent=2), encoding="utf-8")

def resolve_lo_python() -> Path | None:
    """
    Returns the path to LO's python.exe, or None if it can't be found.
    Priority:
      1. Config file (user-specified or previously saved default)
      2. Default install location (auto-saved to config if found)
    """
    cfg = load_config()

    if cfg.get("lo_python"):
        p = Path(cfg["lo_python"])
        if p.exists():
            return p
        # Saved path is now invalid — return None so the dialog can fire
        return None

    # No config yet — try the default
    if DEFAULT_LO_PYTHON.exists():
        save_config({"lo_python": str(DEFAULT_LO_PYTHON)})
        return DEFAULT_LO_PYTHON

    return None