from __future__ import annotations

import logging
from pathlib import Path

from .config import APP_NAME
from .logging_config import configure_logging
from .ui.main_window import MainWindow


def main() -> int:
    project_root = Path(__file__).resolve().parents[2]
    configure_logging(project_root)
    logging.getLogger(__name__).info("Starting %s", APP_NAME)

    app = MainWindow()
    app.mainloop()
    return 0
