import os
import sys

def _get_base_dir() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


BASE_DIR = _get_base_dir()

DATA_DIR = os.path.join(BASE_DIR, "data")
TEMPLATES_DIR = os.path.join(BASE_DIR, "templates")

TRAIN_TYPE_VALUES_FILE = os.path.join(DATA_DIR, "train_type_values.json")
EXCEL_MAPPING_FILE = os.path.join(DATA_DIR, "excel_mapping.json")
METADATA_FILE = os.path.join(DATA_DIR, "metadata.json")

TRAIN_TYPE_TEMPLATE_FILE = os.path.join(TEMPLATES_DIR, "train_type_template.xlsx")

SHAREPOINT_BASE_URL: str = "https://your-tenant.sharepoint.com/sites/your-site/Shared%20Documents"
SHAREPOINT_TRAIN_TYPE_VALUES_URL: str = f"{SHAREPOINT_BASE_URL}/train_type_values.json"
SHAREPOINT_EXCEL_MAPPING_URL: str = f"{SHAREPOINT_BASE_URL}/excel_mapping.json"

REQUEST_TIMEOUT_SECONDS: int = 10
