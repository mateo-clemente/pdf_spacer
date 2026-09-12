import json
import os

CONFIG_FILE = "config.json"

DEFAULT_CONFIG = {
    "input_folder": "input",
    "output_folder": "output",
    "multiplier": 2.0,
    "scan_interval_seconds": 60,
    "continous_mode": 1
}

def load_config() -> dict:
    """Carga la configuración desde el JSON o crea uno nuevo si no existe."""
    if not os.path.exists(CONFIG_FILE):
        save_config(DEFAULT_CONFIG)
        return DEFAULT_CONFIG
    
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        print(f"Error al leer {CONFIG_FILE}, usando valores por defecto: {e}")
        return DEFAULT_CONFIG

def save_config(config_data: dict) -> None:
    """Guarda los cambios en el archivo JSON."""
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(config_data, f, indent=4)
        print(" Configuración guardada correctamente.")
    except Exception as e:
        print(f" Error al guardar {CONFIG_FILE}: {e}")