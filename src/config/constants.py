"""Constants and configuration values for Reconciliator application."""
import os
from pathlib import Path

# Application Info
APP_NAME = "Reconciliator"
APP_VERSION = "1.0.0"

# Paths
# __file__ is src/config/constants.py → .parent.parent = src/ → .parent = repo root
BASE_DIR = Path(__file__).parent.parent          # src/
PROJECT_ROOT = Path(__file__).parent.parent.parent  # repo root (contains icons/, SQL/, etc.)

# Directory structure — all relative to the repo root
REQUIRED_DIRS = {
    "database": os.path.join(PROJECT_ROOT, "Database"),
    "icons": os.path.join(PROJECT_ROOT, "icons"),
    "sql": os.path.join(PROJECT_ROOT, "SQL"),
    "logs": os.path.join(PROJECT_ROOT, "Log"),
    "warnings": os.path.join(PROJECT_ROOT, "Warning Message"),
    "demo": os.path.join(PROJECT_ROOT, "Demo Files"),
}

# Database settings
DATABASE_NAME = "Reconciliator.db"
DATABASE_PATH = os.path.join(REQUIRED_DIRS["database"], DATABASE_NAME)

# File processing settings
DEFAULT_SHEET_NAME = 0  # First sheet
SUPPORTED_EXCEL_FORMATS = ('.xls', '.xlsx')
MAX_COLUMN_NAME_LENGTH = 100

# SQL settings
PREVALIDATION_SQL_FILE = os.path.join(REQUIRED_DIRS["sql"], "prevalidation.sql")

# UI settings
WINDOW_WIDTH = 1200
WINDOW_HEIGHT = 900
ICON_SIZE = (20, 20)

# Icon paths — absolute, based on PROJECT_ROOT
ICON_PATHS = {
    "minus": os.path.join(PROJECT_ROOT, "icons", "minus-sign.png"),
    "resize": os.path.join(PROJECT_ROOT, "icons", "resize.png"),
    "cancel": os.path.join(PROJECT_ROOT, "icons", "cancel.png"),
    "next": os.path.join(PROJECT_ROOT, "icons", "next.png"),
    "upload": os.path.join(PROJECT_ROOT, "icons", "upload.png"),
    "checklist": os.path.join(PROJECT_ROOT, "icons", "checklist.png"),
    "delete": os.path.join(PROJECT_ROOT, "icons", "delete.png"),
    "reload": os.path.join(PROJECT_ROOT, "icons", "reload.png"),
    "restore": os.path.join(PROJECT_ROOT, "icons", "restore.png"),
    "back": os.path.join(PROJECT_ROOT, "icons", "back.png"),
    "loading": os.path.join(PROJECT_ROOT, "icons", "loading.gif"),
}

# Style settings
COLORS = {
    "dark_bg": "rgb(49, 54, 59)",
    "light_bg": "rgb(246, 246, 246)",
    "button_hover": "rgb(79, 84, 89)",
    "border_color": "rgb(237, 237, 237)",
    "text_color": "rgb(222, 222, 222)",
    "white": "white",
}

# Data type mappings
SQL_DATATYPE_MAP = {
    'object': 'text',
    'int': 'integer',
    'int64': 'integer',
    'float': 'real',
    'bool': 'integer',
}
