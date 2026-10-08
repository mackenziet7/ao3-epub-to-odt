
SCHEMA_VERSION = 1
REQUIRED_SECTIONS = ("page_setup", "typography_advanced", "additional_options")

def validate_preset(p: dict) -> list[str]:
    """Return a list of problems; empty means usable by the converter."""
    errors = []
    if p.get("schema_version") != SCHEMA_VERSION:
        errors.append(f"schema_version is {p.get('schema_version')!r}, expected {SCHEMA_VERSION}")
    for section in REQUIRED_SECTIONS:
        if section not in p:
            errors.append(f"missing section: {section}")
    return errors