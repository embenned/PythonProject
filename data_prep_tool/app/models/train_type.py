from dataclasses import dataclass, field
from typing import Any


@dataclass
class TrainTypeValues:
    train_type: str
    values: dict[str, Any] = field(default_factory=dict)

    def get_value(self, key: str) -> Any:
        return self.values.get(key)
