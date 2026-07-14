from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


@dataclass(slots=True)
class UpdateSettings:
    enabled: bool = False
    shared_root: Path | None = None
    manifest_name: str = "latest.json"
    installer_name: str = ""
    channel: str = "stable"


@dataclass(slots=True)
class UpdateRelease:
    version: str
    installer_path: Path
    notes: str = ""
    published_at: str = ""
    mandatory: bool = False


@dataclass(slots=True)
class UpdateCheckResult:
    checked: bool
    available: bool
    current_version: str
    release: UpdateRelease | None = None
    reason: str = ""
