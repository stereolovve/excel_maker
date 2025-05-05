#config/config.py

import os
from pathlib import Path
APP_VERSION = "1.0.0"

API_URL = os.getenv("API_URL", "http://perplan.tech")
EXCEL_BASE_DIR = os.getenv("EXCEL_BASE_DIR", "C:\\Users\\lucas.melo\\excel_maker\\output")
LOG_FILE = Path.cwd() / "log.txt"  
