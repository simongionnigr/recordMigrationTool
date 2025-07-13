import os
import yaml

CONFIG_FILE = "record_migrator/config.yaml"
DEFAULT_CONFIG = {
    "query_csv": {
        "columns": ["sobject_api", "soql"],
        "prompt": "Carica file query CSV",
        "filetypes": [("CSV file", "*.csv")],
        "separator": ","
    }
}

def load_config():
    """
    Carica config da CONFIG_FILE se esiste, altrimenti ritorna DEFAULT_CONFIG.
    """
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r") as f:
            return yaml.safe_load(f)
    return DEFAULT_CONFIG.copy()

def save_config(cfg: dict):
    """
    Salva il dict cfg su CONFIG_FILE in formato YAML leggibile.
    """
    with open(CONFIG_FILE, "w") as f:
        yaml.safe_dump(cfg, f, sort_keys=False, default_flow_style=False)
