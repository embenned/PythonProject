import json
import logging
from typing import Any

import urllib.request
import urllib.error

from app.config import (
    SHAREPOINT_TRAIN_TYPE_VALUES_URL,
    SHAREPOINT_EXCEL_MAPPING_URL,
    REQUEST_TIMEOUT_SECONDS,
)
from app.repositories.base_repository import BaseRepository
from app.repositories.local_json_repository import LocalJsonRepository

logger = logging.getLogger(__name__)


class SharePointRepository(BaseRepository):

    def __init__(self) -> None:
        self._local_fallback = LocalJsonRepository()

    def _fetch_url(self, url: str) -> dict[str, Any]:
        logger.info("Attempting to fetch from SharePoint: %s", url)
        req = urllib.request.Request(url, headers={"Accept": "application/json"})
        with urllib.request.urlopen(req, timeout=REQUEST_TIMEOUT_SECONDS) as response:
            raw = response.read()
        data = json.loads(raw.decode("utf-8"))
        logger.info("Successfully loaded from SharePoint: %s", url)
        return data

    def _fetch_with_fallback(
        self,
        url: str,
        local_loader,
        label: str,
    ) -> dict[str, Any]:
        try:
            data = self._fetch_url(url)
            logger.info("Source used for %s: SharePoint", label)
            return data
        except Exception as exc:
            logger.warning(
                "SharePoint unavailable for %s (%s). Falling back to local.", label, exc
            )
            data = local_loader()
            logger.info("Source used for %s: local JSON", label)
            return data

    def load_train_type_values(self) -> dict[str, Any]:
        return self._fetch_with_fallback(
            SHAREPOINT_TRAIN_TYPE_VALUES_URL,
            self._local_fallback.load_train_type_values,
            "train_type_values",
        )

    def load_excel_mapping(self) -> dict[str, Any]:
        return self._fetch_with_fallback(
            SHAREPOINT_EXCEL_MAPPING_URL,
            self._local_fallback.load_excel_mapping,
            "excel_mapping",
        )
