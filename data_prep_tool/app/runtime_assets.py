import logging
import os
import shutil

from app.config import (
    APP_DIR,
    BUNDLED_EXCEL_MAPPING_FILE,
    BUNDLED_METADATA_FILE,
    BUNDLED_TRAIN_TYPE_TEMPLATE_FILE,
    BUNDLED_TRAIN_TYPE_VALUES_FILE,
    DATA_DIR,
    LOGS_DIR,
    METADATA_FILE,
    TEMPLATES_DIR,
    TRAIN_TYPE_TEMPLATE_FILE,
    EXCEL_MAPPING_FILE,
    TRAIN_TYPE_VALUES_FILE,
)

logger = logging.getLogger(__name__)


def _copy_if_missing(source_path: str, target_path: str) -> None:
    if os.path.abspath(source_path) == os.path.abspath(target_path):
        return
    if os.path.exists(target_path):
        return
    if not os.path.exists(source_path):
        logger.warning("Bundled asset missing: %s", source_path)
        return

    shutil.copy2(source_path, target_path)
    logger.info("Initialized runtime asset: %s", target_path)


def ensure_runtime_assets() -> None:
    os.makedirs(APP_DIR, exist_ok=True)
    os.makedirs(DATA_DIR, exist_ok=True)
    os.makedirs(LOGS_DIR, exist_ok=True)
    os.makedirs(TEMPLATES_DIR, exist_ok=True)

    _copy_if_missing(BUNDLED_TRAIN_TYPE_VALUES_FILE, TRAIN_TYPE_VALUES_FILE)
    _copy_if_missing(BUNDLED_EXCEL_MAPPING_FILE, EXCEL_MAPPING_FILE)
    _copy_if_missing(BUNDLED_METADATA_FILE, METADATA_FILE)
    _copy_if_missing(BUNDLED_TRAIN_TYPE_TEMPLATE_FILE, TRAIN_TYPE_TEMPLATE_FILE)