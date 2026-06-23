import json
import logging
import os
from datetime import datetime, timezone
from typing import Any

from app.config import METADATA_FILE

logger = logging.getLogger(__name__)


class DataUpdateService:

    def get_metadata(self) -> dict[str, Any]:
        if not os.path.exists(METADATA_FILE):
            logger.warning("metadata.json not found at: %s", METADATA_FILE)
            return {}
        with open(METADATA_FILE, encoding="utf-8") as fh:
            return json.load(fh)

    def record_generation_event(self, train_type: str, output_path: str) -> None:
        metadata = self.get_metadata()
        history: list = metadata.setdefault("generation_history", [])
        history.append(
            {
                "timestamp": datetime.now(timezone.utc).isoformat(),
                "train_type": train_type,
                "output_path": output_path,
            }
        )
        with open(METADATA_FILE, "w", encoding="utf-8") as fh:
            json.dump(metadata, fh, indent=2, ensure_ascii=False)
        logger.info("Recorded generation event for train type '%s'.", train_type)
