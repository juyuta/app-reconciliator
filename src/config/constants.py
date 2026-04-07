"""Constants and configuration values for Reconciliator application."""
import os
from pathlib import Path

# Application Info
APP_NAME = "Reconciliator"
APP_VERSION = "1.0.0"

# Paths
# __file__ is src/config/constants.py → .parent.parent = src/ → .parent = repo root
BASE_DIR = Path(__file__).parent.parent          # src/
PROJECT_ROOT = Path(__file__).parent.parent.parent  # repo root

# Directory structure — all relative to the repo root
# resources/ = static assets shipped with the app
# output/    = runtime-generated files (gitignored)
REQUIRED_DIRS = {
    "database": os.path.join(PROJECT_ROOT, "output", "database"),
    "icons": os.path.join(PROJECT_ROOT, "resources", "icons"),
    "sql": os.path.join(PROJECT_ROOT, "resources", "sql"),
    "logs": os.path.join(PROJECT_ROOT, "output", "logs"),
    "warnings": os.path.join(PROJECT_ROOT, "output", "warnings"),
    "demo": os.path.join(PROJECT_ROOT, "resources", "demo"),
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

# Icon paths — absolute, based on REQUIRED_DIRS["icons"]
ICON_PATHS = {
    "minus": os.path.join(REQUIRED_DIRS["icons"], "minus-sign.png"),
    "resize": os.path.join(REQUIRED_DIRS["icons"], "resize.png"),
    "cancel": os.path.join(REQUIRED_DIRS["icons"], "cancel.png"),
    "next": os.path.join(REQUIRED_DIRS["icons"], "next.png"),
    "upload": os.path.join(REQUIRED_DIRS["icons"], "upload.png"),
    "checklist": os.path.join(REQUIRED_DIRS["icons"], "checklist.png"),
    "delete": os.path.join(REQUIRED_DIRS["icons"], "delete.png"),
    "reload": os.path.join(REQUIRED_DIRS["icons"], "reload.png"),
    "restore": os.path.join(REQUIRED_DIRS["icons"], "restore.png"),
    "back": os.path.join(REQUIRED_DIRS["icons"], "back.png"),
    "loading": os.path.join(REQUIRED_DIRS["icons"], "loading.gif"),
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
