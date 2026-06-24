import os
import sys

def _get_app_dir() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _get_resource_dir() -> str:
    if getattr(sys, "frozen", False):
        return getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


APP_DIR = _get_app_dir()
RESOURCE_DIR = _get_resource_dir()

DATA_DIR = os.path.join(APP_DIR, "data")
LOGS_DIR = os.path.join(APP_DIR, "logs")
TEMPLATES_DIR = os.path.join(APP_DIR, "templates")

RESOURCE_DATA_DIR = os.path.join(RESOURCE_DIR, "data")
RESOURCE_TEMPLATES_DIR = os.path.join(RESOURCE_DIR, "templates")

TRAIN_TYPE_VALUES_FILE = os.path.join(DATA_DIR, "train_type_values.json")
EXCEL_MAPPING_FILE = os.path.join(DATA_DIR, "excel_mapping.json")
METADATA_FILE = os.path.join(DATA_DIR, "metadata.json")
APP_LOG_FILE = os.path.join(LOGS_DIR, "app.log")

TRAIN_TYPE_TEMPLATE_FILE = os.path.join(TEMPLATES_DIR, "train_type_template.xlsx")

BUNDLED_TRAIN_TYPE_VALUES_FILE = os.path.join(RESOURCE_DATA_DIR, "train_type_values.json")
BUNDLED_EXCEL_MAPPING_FILE = os.path.join(RESOURCE_DATA_DIR, "excel_mapping.json")
BUNDLED_METADATA_FILE = os.path.join(RESOURCE_DATA_DIR, "metadata.json")
BUNDLED_TRAIN_TYPE_TEMPLATE_FILE = os.path.join(RESOURCE_TEMPLATES_DIR, "train_type_template.xlsx")

SHAREPOINT_BASE_URL: str = "https://your-tenant.sharepoint.com/sites/your-site/Shared%20Documents"
SHAREPOINT_TRAIN_TYPE_VALUES_URL: str = f"{SHAREPOINT_BASE_URL}/train_type_values.json"
SHAREPOINT_EXCEL_MAPPING_URL: str = f"{SHAREPOINT_BASE_URL}/excel_mapping.json"

REQUEST_TIMEOUT_SECONDS: int = 10
