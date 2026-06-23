import logging
import os
import shutil
from typing import Any

import openpyxl

from app.config import TRAIN_TYPE_TEMPLATE_FILE, EXCEL_MAPPING_FILE
from app.models.train_type import TrainTypeValues
from app.repositories.base_repository import BaseRepository

logger = logging.getLogger(__name__)


class ExcelGenerator:

    def __init__(self, repository: BaseRepository) -> None:
        self._repository = repository
        self._mapping: list[dict[str, Any]] | None = None

    def _ensure_mapping_loaded(self) -> None:
        if self._mapping is None:
            raw = self._repository.load_excel_mapping()
            self._mapping = raw.get("mappings", [])

    def generate(self, train_type_values: TrainTypeValues, output_path: str) -> None:
        self._ensure_mapping_loaded()

        if not os.path.exists(TRAIN_TYPE_TEMPLATE_FILE):
            raise FileNotFoundError(
                f"Template not found: {TRAIN_TYPE_TEMPLATE_FILE}"
            )

        output_dir = os.path.dirname(output_path)
        if output_dir:
            os.makedirs(output_dir, exist_ok=True)

        shutil.copy2(TRAIN_TYPE_TEMPLATE_FILE, output_path)
        logger.info("Copied template to: %s", output_path)

        wb = openpyxl.load_workbook(output_path)

        for entry in self._mapping:
            sheet_name: str = entry.get("sheet", "")
            cell: str = entry.get("cell", "")
            field_key: str = entry.get("field", "")

            if sheet_name not in wb.sheetnames:
                logger.warning("Sheet '%s' not found in template — skipping.", sheet_name)
                continue

            ws = wb[sheet_name]
            value = train_type_values.get_value(field_key)
            if value is None:
                logger.warning(
                    "No value for field '%s' (train type '%s') — cell %s!%s left unchanged.",
                    field_key,
                    train_type_values.train_type,
                    sheet_name,
                    cell,
                )
                continue

            ws[cell] = value
            logger.debug("Wrote %s -> %s!%s", field_key, sheet_name, cell)

        wb.save(output_path)
        logger.info("Excel file saved: %s", output_path)
