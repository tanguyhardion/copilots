"""
Python Copilot Execution Runner & Workspace Sandbox Manager.
Manages isolated execution folders, running scripts, capturing stdout/stderr,
tracking generated files, formatting folder context for LLMs, and explorer integration.
"""

import os
import sys
import time
import shutil
import tempfile
import subprocess
from pathlib import Path
from typing import Dict, Any, List, Optional


class PythonSandboxRunner:
    """
    Manages execution of pasted Python scripts in a dedicated sandbox environment.
    Supports:
    - Session-based temporary execution with auto-cleanup (persist=False by default)
    - Persistent project directory mode (persist=True)
    - Capturing stdout, stderr, exit code, and execution time
    - Tracking newly generated or modified files
    - Generating markdown context of all files in the directory for copying to LLMs
    - Opening current folder in Windows File Explorer
    """

    DEFAULT_SCRIPT_NAME = "runner_script.py"

    def __init__(self):
        self.persist_mode: bool = False
        # Base persistent directory in AppData or dev folder
        appdata = os.environ.get("APPDATA") or os.path.expanduser("~")
        self.persistent_dir = Path(appdata) / "CopilotsApp" / "python_sandbox"
        self.persistent_dir.mkdir(parents=True, exist_ok=True)

        # Active temporary session dir
        self._temp_dir: Optional[tempfile.TemporaryDirectory] = None
        self._current_working_dir: Path = self._create_temp_workspace()

    def _create_temp_workspace(self) -> Path:
        """Create a fresh temporary folder for scratchpad execution."""
        if self._temp_dir is not None:
            try:
                self._temp_dir.cleanup()
            except Exception:
                pass
        self._temp_dir = tempfile.TemporaryDirectory(prefix="copilots_py_")
        return Path(self._temp_dir.name)

    @property
    def current_dir(self) -> Path:
        """Return the current active working directory (persistent or temp)."""
        if self.persist_mode:
            self.persistent_dir.mkdir(parents=True, exist_ok=True)
            return self.persistent_dir
        if not self._current_working_dir.exists():
            self._current_working_dir = self._create_temp_workspace()
        return self._current_working_dir

    def set_persistence(self, persist: bool) -> Dict[str, Any]:
        """Toggle persistence mode."""
        self.persist_mode = persist
        if persist:
            # If switching to persist, copy any non-script files from temp to persistent dir
            if self._current_working_dir.exists():
                for item in self._current_working_dir.iterdir():
                    if item.name != self.DEFAULT_SCRIPT_NAME:
                        dest = self.persistent_dir / item.name
                        try:
                            if item.is_dir():
                                if dest.exists():
                                    shutil.rmtree(dest)
                                shutil.copytree(item, dest)
                            else:
                                shutil.copy2(item, dest)
                        except Exception as e:
                            print(f"[PythonRunner] Error copying to persistent dir: {e}")
        return {
            "success": True,
            "persist": self.persist_mode,
            "folder_path": str(self.current_dir),
            "files": self.list_files()
        }

    def clear_sandbox(self) -> Dict[str, Any]:
        """Clear current sandbox workspace."""
        try:
            target = self.current_dir
            for item in target.iterdir():
                try:
                    if item.is_dir():
                        shutil.rmtree(item)
                    else:
                        item.unlink()
                except Exception as e:
                    print(f"[PythonRunner] Error removing {item}: {e}")

            if not self.persist_mode:
                self._current_working_dir = self._create_temp_workspace()

            return {
                "success": True,
                "message": "Sandbox workspace cleared.",
                "folder_path": str(self.current_dir),
                "files": []
            }
        except Exception as err:
            return {"success": False, "error": f"Failed to clear sandbox: {err}"}

    def open_in_explorer(self) -> Dict[str, Any]:
        """Open the active working directory in Windows File Explorer."""
        folder = self.current_dir
        folder.mkdir(parents=True, exist_ok=True)
        try:
            if sys.platform == "win32":
                os.startfile(str(folder))
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(folder)])
            else:
                subprocess.Popen(["xdg-open", str(folder)])
            return {"success": True, "path": str(folder)}
        except Exception as err:
            return {"success": False, "error": f"Failed to open explorer: {err}"}

    def run_code(self, code: str, persist: Optional[bool] = None) -> Dict[str, Any]:
        """
        Execute Python code in current sandbox directory.
        Captures stdout, stderr, return code, execution time, and produced files.
        """
        if persist is not None:
            self.persist_mode = persist

        work_dir = self.current_dir
        work_dir.mkdir(parents=True, exist_ok=True)

        # Snapshot files before execution to detect newly created or modified files
        before_files = {f.name: f.stat().st_mtime for f in work_dir.glob("**/*") if f.is_file()}

        # Write code to script file in the sandbox
        script_file = work_dir / self.DEFAULT_SCRIPT_NAME
        try:
            with open(script_file, "w", encoding="utf-8") as f:
                f.write(code)
        except Exception as err:
            return {
                "success": False,
                "error": f"Could not write script to disk: {err}",
                "stdout": "",
                "stderr": str(err),
                "exit_code": -1,
                "duration_ms": 0,
                "files": []
            }

        start_time = time.time()
        try:
            # Use current Python interpreter (respects venv)
            python_bin = sys.executable
            # Ensure unbuffered output (-u)
            process = subprocess.Popen(
                [python_bin, "-u", str(script_file.name)],
                cwd=str(work_dir),
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                encoding="utf-8",
                errors="replace"
            )

            stdout, stderr = process.communicate(timeout=60)
            exit_code = process.returncode
        except subprocess.TimeoutExpired:
            process.kill()
            stdout, stderr = process.communicate()
            exit_code = -1
            stderr += "\n[Execution Timed Out after 60 seconds]"
        except Exception as err:
            stdout = ""
            stderr = f"Execution error: {err}"
            exit_code = -1

        duration_ms = int((time.time() - start_time) * 1000)

        # Detect files after execution
        after_files = self.list_files()

        # Identify newly created or modified files
        produced_files = []
        for f in after_files:
            if f["name"] == self.DEFAULT_SCRIPT_NAME:
                continue
            name = f["name"]
            if name not in before_files or f["mtime"] > before_files[name]:
                produced_files.append(f)

        return {
            "success": exit_code == 0,
            "stdout": stdout,
            "stderr": stderr,
            "exit_code": exit_code,
            "duration_ms": duration_ms,
            "produced_files": produced_files,
            "all_files": [f for f in after_files if f["name"] != self.DEFAULT_SCRIPT_NAME],
            "folder_path": str(work_dir),
            "persist": self.persist_mode
        }

    def list_files(self) -> List[Dict[str, Any]]:
        """List files in the current working directory with metadata."""
        folder = self.current_dir
        if not folder.exists():
            return []

        results = []
        try:
            for item in folder.iterdir():
                if item.name == self.DEFAULT_SCRIPT_NAME:
                    continue
                try:
                    stat = item.stat()
                    size_bytes = stat.st_size if item.is_file() else 0
                    results.append({
                        "name": item.name,
                        "path": str(item),
                        "is_dir": item.is_dir(),
                        "size_bytes": size_bytes,
                        "size_formatted": self._format_size(size_bytes) if item.is_file() else "Folder",
                        "mtime": stat.st_mtime,
                        "ext": item.suffix.lower().lstrip("."),
                    })
                except Exception:
                    continue
        except Exception as e:
            print(f"[PythonRunner] Error listing files: {e}")

        # Sort files by name
        results.sort(key=lambda x: (not x["is_dir"], x["name"].lower()))
        return results

    def get_folder_context_for_llm(self) -> Dict[str, Any]:
        """
        Generate clean structured markdown description of the folder and its files
        ready to copy & paste into ChatGPT or other LLMs.
        """
        folder = self.current_dir
        files = self.list_files()

        lines = [
            f"### Working Directory Context",
            f"**Path**: `{folder}`",
            f"**Persistence Mode**: {'Enabled (Persistent)' if self.persist_mode else 'Disabled (Temporary Sandbox)'}",
            f"**Total Files**: {len(files)}",
            "",
            "#### Files in Directory:",
        ]

        if not files:
            lines.append("*(Directory is currently empty. No documents or output files present.)*")
        else:
            lines.append("| Filename | Type | Size |")
            lines.append("| :--- | :--- | :--- |")
            for f in files:
                ftype = "Folder" if f["is_dir"] else (f["ext"].upper() or "File")
                lines.append(f"| `{f['name']}` | {ftype} | {f['size_formatted']} |")

            lines.append("")
            lines.append("#### Quick Reference File Names (copyable list):")
            lines.append("```text")
            for f in files:
                lines.append(f["name"])
            lines.append("```")

        context_md = "\n".join(lines)
        return {
            "success": True,
            "folder_path": str(folder),
            "file_count": len(files),
            "files": files,
            "context_markdown": context_md
        }

    def open_specific_file(self, filename: str) -> Dict[str, Any]:
        """Open a specific file from the sandbox in the default OS application."""
        file_path = self.current_dir / filename
        if not file_path.exists():
            return {"success": False, "error": f"File '{filename}' does not exist."}

        try:
            if sys.platform == "win32":
                os.startfile(str(file_path))
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(file_path)])
            else:
                subprocess.Popen(["xdg-open", str(file_path)])
            return {"success": True, "message": f"Opened {filename}"}
        except Exception as err:
            return {"success": False, "error": f"Could not open file: {err}"}

    @staticmethod
    def _format_size(size_bytes: int) -> str:
        """Format byte size into human readable string."""
        if size_bytes < 1024:
            return f"{size_bytes} B"
        elif size_bytes < 1024 * 1024:
            return f"{size_bytes / 1024:.1f} KB"
        elif size_bytes < 1024 * 1024 * 1024:
            return f"{size_bytes / (1024 * 1024):.1f} MB"
        return f"{size_bytes / (1024 * 1024 * 1024):.1f} GB"
