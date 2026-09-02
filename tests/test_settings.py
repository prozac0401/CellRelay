import json
from pathlib import Path

from app.config.settings import AppSettings, RuntimeProgress, SettingsStore


def test_settings_and_progress_round_trip(tmp_path: Path) -> None:
    store = SettingsStore(tmp_path / "config")
    settings = AppSettings(
        excel_file="C:/work/data.xlsx",
        sheet="Sheet1",
        start_cell="B5",
        url="https://example.com",
        selector="textarea",
    )
    store.save_settings(settings)
    assert store.load_settings() == settings

    progress = RuntimeProgress(
        excel_file=settings.excel_file,
        sheet=settings.sheet,
        start_cell=settings.start_cell,
        current_cell="B17",
        last_completed_cell="B16",
        last_skipped_cell="B15",
        processed_count=12,
        skipped_count=1,
        phase="WAITING_FOR_CLEAR",
    )
    store.save_progress(progress)
    loaded = store.load_progress()
    assert loaded is not None
    assert loaded.current_cell == "B17"
    assert loaded.last_completed_cell == "B16"
    assert loaded.last_skipped_cell == "B15"
    assert loaded.skipped_count == 1
    assert loaded.updated_at

    raw = json.loads(store.progress_path.read_text(encoding="utf-8"))
    assert raw["phase"] == "WAITING_FOR_CLEAR"


def test_unknown_settings_keys_are_ignored(tmp_path: Path) -> None:
    store = SettingsStore(tmp_path)
    store.settings_path.write_text(
        '{"start_cell": "C3", "future_option": true}',
        encoding="utf-8",
    )
    assert store.load_settings().start_cell == "C3"
