# record_migrator/config.py

import os
import yaml

MODULE_DIR  = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.getenv(
    "CONFIG_PATH",
    os.path.join(MODULE_DIR, "import_config.yaml")
)

# Nuovo file per le impostazioni di import
IMPORT_CONFIG_FILE = os.getenv(
    "IMPORT_CONFIG_PATH",
    os.path.join(MODULE_DIR, "import_config.yaml")
)


DEFAULT_CONFIG = {
    "query_csv": {
        "columns": ["sobject_api", "soql"],
        "prompt": "Carica file query CSV",
        "filetypes": [("CSV file", "*.csv")],
        "separator": ";"
    },
    "ui": {
        "theme": "clam"
    }
}

DEFAULT_IMPORT_CFG = {
    "input_tables": None,
    "import_settings": {},
    "import_order": [],
    "ignore_columns": ["record_id","to_import","sf_id","error","Id"],
    "relationships_file": None
}

def main():
    load_import_config()

def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r") as f:
            cfg = yaml.safe_load(f) or {}
            # merge default for missing keys
            merged = DEFAULT_CONFIG.copy()
            merged.update(cfg)
            # ensure nested ui
            merged["ui"] = DEFAULT_CONFIG["ui"].copy()
            merged["ui"].update(cfg.get("ui", {}))
            return merged
    return DEFAULT_CONFIG.copy()

def save_config(cfg: dict):
    with open(CONFIG_FILE, "w") as f:
        yaml.safe_dump(cfg, f, sort_keys=False, default_flow_style=False)

# --- nuove funzioni per import_config.yaml ---

def load_import_config():
    if os.path.exists(IMPORT_CONFIG_FILE):
        with open(IMPORT_CONFIG_FILE, "r") as f:
            cfg = yaml.safe_load(f) or {}
        # merge default e file
        merged = DEFAULT_IMPORT_CFG.copy()
        merged.update(cfg)
        return merged
    else:
        return DEFAULT_IMPORT_CFG.copy()

def save_import_config(cfg: dict):
    """
    Salva il dict cfg su import_config.yaml.
    """
    with open(IMPORT_CONFIG_FILE, "w") as f:
        yaml.safe_dump(cfg, f, sort_keys=False, default_flow_style=False)

