"""Typed JSON settings and runtime-progress persistence."""

from __future__ import annotations

import json
import logging
import os
from dataclasses import asdict, dataclass, fields
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, TypeVar

logger = logging.getLogger(__name__)
T = TypeVar("T")
AWS_LOGIN_URL = (
    "https://skillbuilder.aws/login?redirect=https%3A%2F%2Fskillbuilder.aws%2F"
)


@dataclass(slots=True)
class AppSettings:
    """User-editable settings shown in the main window."""

    excel_file: str = ""
    sheet: str = ""
    start_cell: str = "B5"
    url: str = AWS_LOGIN_URL
    settings_version: int = 2
    selector: str = "textarea"
    workflow_mode: str = "aws_skill_builder"
    browser_channel: str = "msedge"
    input_method: str = "fill"
    stable_empty_ms: int = 500
    status_check_interval_ms: int = 100


@dataclass(slots=True)
class RuntimeProgress:
    """Crash-recovery data. Recovery UI is intentionally outside the MVP."""

    excel_file: str = ""
    sheet: str = ""
    start_cell: str = ""
    current_cell: str = ""
    current_value: str = ""
    last_completed_cell: str = ""
    last_skipped_cell: str = ""
    processed_count: int = 0
    skipped_count: int = 0
    total_items: int = 0
    url: str = ""
    selector: str = ""
    workflow_mode: str = "aws_skill_builder"
    phase: str = "IDLE"
    updated_at: str = ""
    confirmed_training_url: str = ""
    terminal_outcome: str = ""
    last_error: str = ""
    error_report_backup: str = ""


def _from_mapping(model_type: type[T], data: dict[str, Any]) -> T:
    known_names = {field.name for field in fields(model_type)}
    return model_type(
        **{key: value for key, value in data.items() if key in known_names}
    )


class SettingsStore:
    """Read and atomically write CellRelay JSON files."""

    def __init__(self, config_dir: Path) -> None:
        self.config_dir = config_dir
        self.settings_path = config_dir / "settings.json"
        self.progress_path = config_dir / "progress.json"

    def load_settings(self) -> AppSettings:
        data = self._read_json(self.settings_path)
        if data is None:
            settings = AppSettings()
            self.save_settings(settings)
            return settings
        try:
            if data.get("settings_version", 1) < 2:
                if (
                    data.get("workflow_mode", "aws_skill_builder")
                    == "aws_skill_builder"
                ):
                    data["url"] = AWS_LOGIN_URL
                data["settings_version"] = 2
            return _from_mapping(AppSettings, data)
        except (TypeError, ValueError):
            logger.exception("Invalid settings file: %s", self.settings_path)
            return AppSettings()

    def save_settings(self, settings: AppSettings) -> None:
        self._write_json(self.settings_path, asdict(settings))

    def load_progress(self) -> RuntimeProgress | None:
        data = self._read_json(self.progress_path)
        if data is None:
            return None
        try:
            return _from_mapping(RuntimeProgress, data)
        except (TypeError, ValueError):
            logger.exception("Invalid progress file: %s", self.progress_path)
            return None

    def save_progress(self, progress: RuntimeProgress) -> None:
        progress.updated_at = datetime.now(timezone.utc).isoformat()
        self._write_json(self.progress_path, asdict(progress))

    @staticmethod
    def _read_json(path: Path) -> dict[str, Any] | None:
        if not path.exists():
            return None
        try:
            with path.open("r", encoding="utf-8") as handle:
                value = json.load(handle)
            if not isinstance(value, dict):
                raise TypeError("JSON root must be an object")
            return value
        except (OSError, json.JSONDecodeError, TypeError):
            logger.exception("Could not read JSON file: %s", path)
            return None

    @staticmethod
    def _write_json(path: Path, data: dict[str, Any]) -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
        temp_path = path.with_name(f"{path.name}.tmp")
        try:
            with temp_path.open("w", encoding="utf-8", newline="\n") as handle:
                json.dump(data, handle, ensure_ascii=False, indent=2)
                handle.write("\n")
                handle.flush()
                os.fsync(handle.fileno())
            os.replace(temp_path, path)
        except OSError:
            logger.exception("Could not write JSON file: %s", path)
            raise
