#config/config.py

import os

APP_VERSION = "1.0.0"

API_URL = os.getenv("API_URL", "http://127.0.0.1:8000")
EXCEL_BASE_DIR = os.getenv("EXCEL_BASE_DIR", "C:\\Users\\lucas.melo\\excel_maker\\config")
LOG_FILE = Path.cwd() / "log.txt"  
