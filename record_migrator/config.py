# record_migrator/config.py

import os
import yaml

MODULE_DIR  = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(MODULE_DIR, "config.yaml")

# Nuovo file per le impostazioni di import
IMPORT_CONFIG_FILE = os.path.join(MODULE_DIR, "import_config.yaml")

DEFAULT_CONFIG = {
    "query_csv": {
        "columns": ["sobject_api", "soql"],
        "prompt": "Carica file query CSV",
        "filetypes": [("CSV file", "*.csv")],
        "separator": ";"
    }
}

def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r") as f:
            return yaml.safe_load(f)
    return DEFAULT_CONFIG.copy()

def save_config(cfg: dict):
    with open(CONFIG_FILE, "w") as f:
        yaml.safe_dump(cfg, f, sort_keys=False, default_flow_style=False)

# --- nuove funzioni per import_config.yaml ---

def load_import_config():
    """
    Carica le impostazioni di import da import_config.yaml,
    oppure ritorna {} se il file non esiste.
    """
    if os.path.exists(IMPORT_CONFIG_FILE):
        with open(IMPORT_CONFIG_FILE, "r") as f:
            return yaml.safe_load(f) or {}
    return {}

def save_import_config(cfg: dict):
    """
    Salva il dict cfg su import_config.yaml.
    """
    with open(IMPORT_CONFIG_FILE, "w") as f:
        yaml.safe_dump(cfg, f, sort_keys=False, default_flow_style=False)
