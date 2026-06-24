from dataclasses import dataclass, field
from typing import Any, Mapping

from openpyxl.utils.cell import coordinate_from_string
from openpyxl.utils.exceptions import CellCoordinatesException


class DataValidationError(ValueError):
    pass


@dataclass
class TrainTypeValues:
    train_type: str
    values: dict[str, Any] = field(default_factory=dict)

    def get_value(self, key: str) -> Any:
        return self.values.get(key)

    @classmethod
    def from_mapping(cls, train_type: str, raw_values: Mapping[str, Any]) -> "TrainTypeValues":
        if not isinstance(train_type, str) or not train_type.strip():
            raise DataValidationError("Train type name must be a non-empty string.")
        if not isinstance(raw_values, Mapping):
            raise DataValidationError(f"Train type '{train_type}' must map to an object.")

        normalized: dict[str, Any] = {}
        for key, value in raw_values.items():
            if not isinstance(key, str) or not key.strip():
                raise DataValidationError(
                    f"Train type '{train_type}' contains an invalid field name: {key!r}."
                )
            normalized[key] = value

        return cls(train_type=train_type, values=normalized)


@dataclass(slots=True)
class ExcelMappingEntry:
    sheet: str
    cell: str
    field: str

    @classmethod
    def from_mapping(cls, raw_entry: Mapping[str, Any]) -> "ExcelMappingEntry":
        if not isinstance(raw_entry, Mapping):
            raise DataValidationError("Excel mapping entries must be objects.")

        sheet = raw_entry.get("sheet", "")
        cell = raw_entry.get("cell", "")
        field = raw_entry.get("field", "")

        if not isinstance(sheet, str) or not sheet.strip():
            raise DataValidationError("Excel mapping entry is missing a valid 'sheet'.")
        if not isinstance(cell, str) or not cell.strip():
            raise DataValidationError(f"Excel mapping entry for sheet '{sheet}' is missing a valid 'cell'.")
        if not isinstance(field, str) or not field.strip():
            raise DataValidationError(f"Excel mapping entry for sheet '{sheet}' is missing a valid 'field'.")

        try:
            coordinate_from_string(cell)
        except (ValueError, CellCoordinatesException) as exc:
            raise DataValidationError(
                f"Excel mapping entry for sheet '{sheet}' uses an invalid cell reference: {cell!r}."
            ) from exc

        return cls(sheet=sheet, cell=cell, field=field)
