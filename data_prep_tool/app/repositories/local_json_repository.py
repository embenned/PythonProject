import json
import logging
from typing import Any

from app.config import TRAIN_TYPE_VALUES_FILE, EXCEL_MAPPING_FILE
from app.repositories.base_repository import BaseRepository

logger = logging.getLogger(__name__)


class LocalJsonRepository(BaseRepository):

    def _load_file(self, path: str) -> dict[str, Any]:
        logger.info("Loading local file: %s", path)
        with open(path, encoding="utf-8") as fh:
            return json.load(fh)

    def load_train_type_values(self) -> dict[str, Any]:
        return self._load_file(TRAIN_TYPE_VALUES_FILE)

    def load_excel_mapping(self) -> dict[str, Any]:
        return self._load_file(EXCEL_MAPPING_FILE)
