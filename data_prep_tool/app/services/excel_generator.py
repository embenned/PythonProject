import logging
import os
import shutil

import openpyxl

from app.config import TRAIN_TYPE_TEMPLATE_FILE
from app.models.train_type import DataValidationError, ExcelMappingEntry, TrainTypeValues
from app.repositories.base_repository import BaseRepository

logger = logging.getLogger(__name__)


class ExcelGenerator:

    def __init__(self, repository: BaseRepository) -> None:
        self._repository = repository
        self._mapping: list[ExcelMappingEntry] | None = None

    def _ensure_mapping_loaded(self) -> None:
        if self._mapping is None:
            raw = self._repository.load_excel_mapping()
            mappings = raw.get("mappings", [])
            if not isinstance(mappings, list):
                raise DataValidationError("The top-level 'mappings' value must be a list.")
            self._mapping = [ExcelMappingEntry.from_mapping(entry) for entry in mappings]

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
            if entry.sheet not in wb.sheetnames:
                logger.warning("Sheet '%s' not found in template — skipping.", entry.sheet)
                continue

            ws = wb[entry.sheet]
            value = train_type_values.get_value(entry.field)
            if value is None:
                logger.warning(
                    "No value for field '%s' (train type '%s') — cell %s!%s left unchanged.",
                    entry.field,
                    train_type_values.train_type,
                    entry.sheet,
                    entry.cell,
                )
                continue

            ws[entry.cell] = value
            logger.debug("Wrote %s -> %s!%s", entry.field, entry.sheet, entry.cell)

        wb.save(output_path)
        logger.info("Excel file saved: %s", output_path)
