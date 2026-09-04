import json
from pathlib import Path

APP_NAME = "AO3toODT"
DEFAULT_LO_PYTHON = Path(r"C:\Program Files\LibreOffice\program\python.exe")
INVALID_FILE_CHARS = set('<>:"/\\|?*')

PRESET_SCHEMA = 1

# ------------------------------------------------------------------
# Config
# ------------------------------------------------------------------
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

# ------------------------------------------------------------------
# Presets
# ------------------------------------------------------------------
def load_preset(path: Path | str) -> dict:
    p = Path(path)
    return json.loads(p.read_text(encoding="utf-8"))

def save_preset(name: str, data: dict) -> None:
    full_data = {"preset_name": name, "schema_version": PRESET_SCHEMA, **data}
    p = preset_name_to_path(name)
    p.write_text(json.dumps(full_data, indent=2), encoding="utf-8")

def preset_name_to_path(name: str) -> Path:
    n = "preset_" + name.replace(" ", "_") + ".json"
    return config_path().parent / n

# ------------------------------------------------------------------
# Misc
# ------------------------------------------------------------------
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

def lo_missing_reason(main_window) -> str | None:
    """
    Returns a human-readable reason why LibreOffice/UNO isn't available,
    or None if main_window._lo_python is set and conversion can proceed.

    Centralized so every page that gates a button on LO availability
    (page 0's Quick Convert, page 4's wizard Next/Convert, etc.) shows
    the exact same message.
    """
    if getattr(main_window, "_lo_python", None) is None:
        return "LibreOffice was not found. Set the LibreOffice path in Settings to continue."
    return None