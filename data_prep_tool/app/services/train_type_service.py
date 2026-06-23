import logging
from typing import Any

from app.models.train_type import TrainTypeValues
from app.repositories.base_repository import BaseRepository

logger = logging.getLogger(__name__)


class TrainTypeService:

    def __init__(self, repository: BaseRepository) -> None:
        self._repository = repository
        self._raw: dict[str, Any] | None = None

    def _ensure_loaded(self) -> None:
        if self._raw is None:
            self._raw = self._repository.load_train_type_values()

    def get_available_train_types(self) -> list[str]:
        self._ensure_loaded()
        return list(self._raw.get("train_types", {}).keys())

    def get_train_type_values(self, train_type: str) -> TrainTypeValues:
        self._ensure_loaded()
        all_types: dict = self._raw.get("train_types", {})
        if train_type not in all_types:
            raise ValueError(f"Unknown train type: '{train_type}'")
        return TrainTypeValues(
            train_type=train_type,
            values=all_types[train_type],
        )
