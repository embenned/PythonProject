import logging
import os
import tkinter as tk

from app.config import APP_LOG_FILE, LOGS_DIR
from app.runtime_assets import ensure_runtime_assets
from app.repositories.sharepoint_repository import SharePointRepository
from app.services.data_update_service import DataUpdateService
from app.services.excel_generator import ExcelGenerator
from app.services.train_type_service import TrainTypeService
from app.ui.main_window import MainWindow

logger = logging.getLogger(__name__)


def main() -> None:
    ensure_runtime_assets()
    os.makedirs(LOGS_DIR, exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        handlers=[
            logging.StreamHandler(),
            logging.FileHandler(APP_LOG_FILE, encoding="utf-8"),
        ],
    )
    logger.info("Starting Train Type Data Preparation Tool.")

    repository = SharePointRepository()

    train_type_service = TrainTypeService(repository)
    excel_generator = ExcelGenerator(repository)
    data_update_service = DataUpdateService()

    root = tk.Tk()
    window = MainWindow(
        root=root,
        train_type_service=train_type_service,
        excel_generator=excel_generator,
        data_update_service=data_update_service,
    )
    window.run()

    logger.info("Application closed.")


if __name__ == "__main__":
    main()
