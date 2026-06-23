from abc import ABC, abstractmethod
from typing import Any


class BaseRepository(ABC):

    @abstractmethod
    def load_train_type_values(self) -> dict[str, Any]:
        """Return the full train_type_values structure as a dict."""

    @abstractmethod
    def load_excel_mapping(self) -> dict[str, Any]:
        """Return the excel_mapping structure as a dict."""
