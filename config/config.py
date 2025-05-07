#config/config.py

import os
from pathlib import Path
APP_VERSION = "1.0.1"

API_URL = os.getenv("API_URL", "http://perplan.tech")
EXCEL_BASE_DIR = os.getenv("EXCEL_BASE_DIR", "Z:\\0Digitacoes\\Excel_maker\\")
LOG_FILE = Path.cwd() / "log.txt"  
