"""
Application configuration and persistent settings manager.
"""

import os
import json
from pathlib import Path
from dataclasses import dataclass, asdict
from typing import Dict, Any


@dataclass
class AppSettings:
    dark_mode: bool = True
    default_copilot: str = "powerpoint"
    auto_backup: bool = True
    svg_cache_dir: str = "icons"
    recent_files: list[str] = None

    def __post_init__(self):
        if self.recent_files is None:
            self.recent_files = []


class ConfigManager:
    """Manages reading and writing application settings to JSON."""

    CONFIG_FILE = Path(os.path.expanduser("~/.copilot_suite_config.json"))

    _instance = None

    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
            cls._instance.settings = cls._instance._load()
        return cls._instance

    def _load(self) -> AppSettings:
        if self.CONFIG_FILE.exists():
            try:
                with open(self.CONFIG_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    return AppSettings(**data)
            except Exception:
                pass
        return AppSettings()

    def save(self):
        try:
            with open(self.CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(asdict(self.settings), f, indent=2)
        except Exception as err:
            print(f"[config] Failed to save settings: {err}")

    def add_recent_file(self, file_path: str):
        if not file_path:
            return
        if file_path in self.settings.recent_files:
            self.settings.recent_files.remove(file_path)
        self.settings.recent_files.insert(0, file_path)
        self.settings.recent_files = self.settings.recent_files[:10]
        self.save()
