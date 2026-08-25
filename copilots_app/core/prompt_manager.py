"""
Core Prompt Manager: Handles resolution, user overrides in AppData, saving,
resetting, and clipboard preparation across Dev and Production (.exe) builds.
"""

import os
import sys
from pathlib import Path
from typing import Dict, Optional


class PromptManager:
    """Manages prompt resolution between bundled defaults and user custom overrides."""

    PROMPTS_DIR_NAME = "prompts"
    APP_DATA_FOLDER_NAME = "CopilotsApp"

    # Map copilot keys to default filenames
    PROMPT_FILES: Dict[str, str] = {
        "powerpoint": "powerpoint_prompt.md",
        "word": "word_prompt.md",
        "excel": "excel_prompt.md",
        "cv": "cv_prompt.md",
    }

    PROMPT_TITLES: Dict[str, str] = {
        "powerpoint": "PowerPoint Copilot System Prompt",
        "word": "Word Copilot System Prompt",
        "excel": "Excel Copilot System Prompt",
        "cv": "CV Copilot System Prompt",
    }

    _instance = None

    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
            cls._instance._init_paths()
        return cls._instance

    def _init_paths(self):
        # 1. Base directory for bundled defaults (handles PyInstaller sys._MEIPASS)
        if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
            self.bundled_dir = Path(sys._MEIPASS) / "copilots_app" / self.PROMPTS_DIR_NAME
        else:
            # Dev workspace: relative to this file
            self.bundled_dir = Path(__file__).resolve().parent.parent / self.PROMPTS_DIR_NAME

        # 2. User AppData directory (read-write for persistent user customizations)
        if sys.platform == "win32":
            appdata = os.environ.get("APPDATA") or os.path.expanduser("~")
            self.user_dir = Path(appdata) / self.APP_DATA_FOLDER_NAME / self.PROMPTS_DIR_NAME
        else:
            self.user_dir = Path(os.path.expanduser("~/.config")) / self.APP_DATA_FOLDER_NAME / self.PROMPTS_DIR_NAME

        # Ensure user folder exists
        try:
            self.user_dir.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            print(f"[PromptManager] Could not create user prompts dir: {e}")

    def get_prompt_file_name(self, copilot_key: str) -> str:
        return self.PROMPT_FILES.get(copilot_key, f"{copilot_key}_prompt.md")

    def get_user_prompt_path(self, copilot_key: str) -> Path:
        filename = self.get_prompt_file_name(copilot_key)
        return self.user_dir / filename

    def get_bundled_prompt_path(self, copilot_key: str) -> Path:
        filename = self.get_prompt_file_name(copilot_key)
        return self.bundled_dir / filename

    def is_customized(self, copilot_key: str) -> bool:
        """Returns True if the user has a custom override saved."""
        user_path = self.get_user_prompt_path(copilot_key)
        return user_path.exists() and user_path.is_file()

    def get_prompt(self, copilot_key: str) -> str:
        """
        Retrieves prompt text. Checks user custom override first,
        then falls back to bundled default template.
        """
        user_path = self.get_user_prompt_path(copilot_key)
        if user_path.exists() and user_path.is_file():
            try:
                with open(user_path, "r", encoding="utf-8") as f:
                    return f.read()
            except Exception as err:
                print(f"[PromptManager] Error reading user prompt {user_path}: {err}")

        bundled_path = self.get_bundled_prompt_path(copilot_key)
        if bundled_path.exists() and bundled_path.is_file():
            try:
                with open(bundled_path, "r", encoding="utf-8") as f:
                    return f.read()
            except Exception as err:
                print(f"[PromptManager] Error reading bundled prompt {bundled_path}: {err}")

        return f"# System Prompt for {copilot_key.capitalize()} Copilot\n\n(No prompt template found)"

    def get_default_prompt(self, copilot_key: str) -> str:
        """Always returns the factory default prompt from the bundled directory."""
        bundled_path = self.get_bundled_prompt_path(copilot_key)
        if bundled_path.exists() and bundled_path.is_file():
            try:
                with open(bundled_path, "r", encoding="utf-8") as f:
                    return f.read()
            except Exception as err:
                print(f"[PromptManager] Error reading bundled prompt: {err}")
        return ""

    def save_user_prompt(self, copilot_key: str, content: str) -> bool:
        """Saves custom prompt content to user AppData directory."""
        try:
            self.user_dir.mkdir(parents=True, exist_ok=True)
            user_path = self.get_user_prompt_path(copilot_key)
            with open(user_path, "w", encoding="utf-8") as f:
                f.write(content)
            return True
        except Exception as err:
            print(f"[PromptManager] Failed to save custom prompt: {err}")
            return False

    def reset_to_default(self, copilot_key: str) -> bool:
        """Removes the user override file, restoring the bundled default."""
        user_path = self.get_user_prompt_path(copilot_key)
        if user_path.exists():
            try:
                user_path.unlink()
                return True
            except Exception as err:
                print(f"[PromptManager] Failed to remove user prompt override: {err}")
                return False
        return True

    def open_prompts_directory(self):
        """Opens the user prompts directory in OS File Explorer."""
        try:
            self.user_dir.mkdir(parents=True, exist_ok=True)
            if sys.platform == "win32":
                os.startfile(str(self.user_dir))
            elif sys.platform == "darwin":
                os.system(f'open "{self.user_dir}"')
            else:
                os.system(f'xdg-open "{self.user_dir}"')
        except Exception as err:
            print(f"[PromptManager] Failed to open explorer: {err}")
