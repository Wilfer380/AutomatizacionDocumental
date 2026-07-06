from __future__ import annotations

import os
from pathlib import Path


def ensure_directory(path: Path) -> Path:
    path.mkdir(parents=True, exist_ok=True)
    return path


def open_folder(path: Path) -> None:
    if hasattr(os, "startfile"):
        os.startfile(path)  # type: ignore[attr-defined]
        return
    raise OSError("Opening folders is only supported on Windows in this build.")
