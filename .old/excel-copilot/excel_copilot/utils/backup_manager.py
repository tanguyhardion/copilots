"""Backup Manager: Handles automated pre-execution workbook backups and restores."""

import os
import shutil
import datetime
from typing import List, Dict, Optional


class BackupManager:
    """Manages file backups before Excel file modifications."""

    def __init__(self, backup_dir: str = ".backups"):
        self.backup_dir = os.path.abspath(backup_dir)
        os.makedirs(self.backup_dir, exist_ok=True)

    def create_backup(self, file_path: str) -> Optional[str]:
        """Create timestamped copy of specified Excel file."""
        if not os.path.exists(file_path):
            return None

        filename = os.path.basename(file_path)
        name, ext = os.path.splitext(filename)
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_filename = f"{name}_backup_{timestamp}{ext}"
        backup_path = os.path.join(self.backup_dir, backup_filename)

        shutil.copy2(file_path, backup_path)
        return backup_path

    def list_backups(self) -> List[Dict[str, str]]:
        """List all available backup files sorted by creation time."""
        if not os.path.exists(self.backup_dir):
            return []

        backups = []
        for fn in os.listdir(self.backup_dir):
            if fn.endswith((".xlsx", ".xlsm")):
                fp = os.path.join(self.backup_dir, fn)
                ctime = os.path.getctime(fp)
                dt_str = datetime.datetime.fromtimestamp(ctime).strftime("%Y-%m-%d %H:%M:%S")
                backups.append({
                    "filename": fn,
                    "path": fp,
                    "created_at": dt_str,
                    "size_bytes": os.path.getsize(fp),
                })

        backups.sort(key=lambda x: x["created_at"], reverse=True)
        return backups

    def restore_backup(self, backup_path: str, target_file_path: str) -> bool:
        """Restore a backup file to original destination."""
        if not os.path.exists(backup_path):
            return False

        shutil.copy2(backup_path, target_file_path)
        return True
