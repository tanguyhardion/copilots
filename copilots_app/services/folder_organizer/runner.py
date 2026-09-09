"""
Folder/File Organizer runner.
Builds recursive folder context for LLM prompting and executes copy-only
reorganization plans described in a simple DSL.
"""

from __future__ import annotations

import os
import re
import shutil
import subprocess
import sys
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple


class FolderOrganizerRunner:
    """Analyzes folders and safely executes copy-only reorganization plans."""

    IMAGE_EXTENSIONS = {
        ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".svg", ".webp", ".tiff", ".ico", ".heic"
    }
    INSTRUCTION_RE = re.compile(
        r'^\s*MOVE\s+(?P<src>"[^"]+"|\S+)\s*->\s*(?P<dst>"[^"]+"|\S+)\s*$',
        re.IGNORECASE,
    )

    def __init__(self):
        self.source_root: Optional[Path] = None
        self.last_output_root: Optional[Path] = None

    def set_source_root(self, folder_path: str) -> Dict[str, Any]:
        path = Path(folder_path).expanduser().resolve()
        if not path.exists() or not path.is_dir():
            return {"success": False, "error": "Selected folder does not exist or is not a directory."}
        self.source_root = path
        return {"success": True, "source_root": str(path)}

    def get_status(self) -> Dict[str, Any]:
        return {
            "success": True,
            "source_root": str(self.source_root) if self.source_root else "",
            "last_output_root": str(self.last_output_root) if self.last_output_root else "",
        }

    def generate_context_for_llm(self) -> Dict[str, Any]:
        if not self.source_root:
            return {"success": False, "error": "No source folder selected yet."}

        source_root = self.source_root
        all_files = self._list_relative_files(source_root)
        tree_lines = self._render_tree(source_root)

        excerpt_lines = ["## File Excerpts (non-image files)", ""]
        excerpted_count = 0
        skipped_count = 0
        for rel_file in all_files:
            if rel_file.suffix.lower() in self.IMAGE_EXTENSIONS:
                continue

            excerpt, skip_reason = self._read_text_excerpt(source_root / rel_file)
            excerpt_lines.append(f"### `{rel_file.as_posix()}`")
            if excerpt is not None:
                excerpted_count += 1
                excerpt_lines.append("```text")
                excerpt_lines.append(excerpt)
                excerpt_lines.append("```")
            else:
                skipped_count += 1
                excerpt_lines.append(f"_Skipped: {skip_reason}_")
            excerpt_lines.append("")

        if excerpted_count == 0 and skipped_count == 0:
            excerpt_lines.append("_No non-image files found._")

        context_markdown = "\n".join(
            [
                "# Folder Organizer Context",
                f"Source root: `{source_root}`",
                "",
                "## Recursive Tree",
                "```text",
                *tree_lines,
                "```",
                "",
                *excerpt_lines,
            ]
        )

        return {
            "success": True,
            "source_root": str(source_root),
            "file_count": len(all_files),
            "excerpted_count": excerpted_count,
            "skipped_count": skipped_count,
            "context_markdown": context_markdown,
        }

    def parse_plan(self, dsl_text: str) -> Dict[str, Any]:
        instructions: List[Dict[str, str]] = []
        errors: List[str] = []

        for idx, raw_line in enumerate(dsl_text.splitlines(), start=1):
            line = raw_line.strip()
            if not line or line.startswith("#") or line.startswith("//"):
                continue

            m = self.INSTRUCTION_RE.match(raw_line)
            if not m:
                errors.append(f"Line {idx}: invalid instruction format.")
                continue

            src_token = self._strip_quotes(m.group("src"))
            dst_token = self._strip_quotes(m.group("dst"))
            src_rel = self._validate_relative_path(src_token)
            dst_rel = self._validate_relative_path(dst_token)

            if not src_rel:
                errors.append(f"Line {idx}: invalid source path '{src_token}'.")
                continue
            if not dst_rel:
                errors.append(f"Line {idx}: invalid destination path '{dst_token}'.")
                continue

            instructions.append(
                {
                    "source_rel": src_rel,
                    "dest_rel": dst_rel,
                    "line": idx,
                }
            )

        return {
            "success": len(instructions) > 0,
            "instructions": instructions,
            "errors": errors,
            "instruction_count": len(instructions),
        }

    def execute_plan(self, dsl_text: str) -> Dict[str, Any]:
        if not self.source_root:
            return {"success": False, "error": "No source folder selected yet."}

        parsed = self.parse_plan(dsl_text)
        instructions = parsed["instructions"]
        errors = list(parsed["errors"])
        source_root = self.source_root

        if not instructions:
            return {
                "success": False,
                "error": "No valid MOVE instructions were parsed.",
                "errors": errors,
                "copied_count": 0,
                "skipped_count": 0,
            }

        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        # Place the output folder next to the source folder (sibling)
        output_base = source_root.parent
        output_root = output_base / f"{source_root.name}_organized_{stamp}"
        output_root.mkdir(parents=True, exist_ok=True)

        copied_count = 0
        skipped_count = 0
        report_lines = []

        for inst in instructions:
            src_rel = inst["source_rel"]
            dst_rel = inst["dest_rel"]
            line_no = inst["line"]

            src_path = (source_root / src_rel).resolve()
            if not src_path.exists() or not src_path.is_file():
                skipped_count += 1
                report_lines.append(f"Line {line_no}: source missing, skipped '{src_rel}'.")
                continue

            if not self._is_within_root(src_path, source_root):
                skipped_count += 1
                report_lines.append(f"Line {line_no}: source path escapes source root, skipped '{src_rel}'.")
                continue

            dst_path = (output_root / dst_rel).resolve()
            if not self._is_within_root(dst_path, output_root):
                skipped_count += 1
                report_lines.append(f"Line {line_no}: destination path escapes output root, skipped '{dst_rel}'.")
                continue

            dst_path.parent.mkdir(parents=True, exist_ok=True)
            try:
                shutil.copy2(src_path, dst_path)
                copied_count += 1
                report_lines.append(f"Line {line_no}: copied '{src_rel}' -> '{dst_rel}'.")
            except Exception as err:
                skipped_count += 1
                report_lines.append(f"Line {line_no}: copy failed for '{src_rel}' ({err}).")

        self.last_output_root = output_root
        success = copied_count > 0
        message = f"Copied {copied_count} file(s) to '{output_root}'. Skipped {skipped_count}."
        if errors:
            message += f" Parse issues: {len(errors)}."

        return {
            "success": success,
            "message": message,
            "output_root": str(output_root),
            "copied_count": copied_count,
            "skipped_count": skipped_count,
            "errors": errors,
            "report_lines": report_lines,
        }

    def open_output_folder(self) -> Dict[str, Any]:
        """Open the last generated output folder in the system file explorer."""
        if not self.last_output_root or not self.last_output_root.exists():
            return {"success": False, "error": "No output folder generated yet."}
        folder = self.last_output_root
        try:
            if sys.platform == "win32":
                os.startfile(str(folder))
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(folder)])
            else:
                subprocess.Popen(["xdg-open", str(folder)])
            return {"success": True, "path": str(folder)}
        except Exception as err:
            return {"success": False, "error": f"Failed to open folder: {err}"}

    def _render_tree(self, source_root: Path) -> List[str]:
        lines = [f"{source_root.name}/"]
        for root, dirs, files in os.walk(source_root):
            dirs.sort(key=lambda d: d.lower())
            files.sort(key=lambda f: f.lower())
            rel_root = Path(root).relative_to(source_root)
            depth = len(rel_root.parts)
            prefix = "  " * depth

            for d in dirs:
                lines.append(f"{prefix}- {d}/")
            for f in files:
                lines.append(f"{prefix}- {f}")
        return lines

    def _list_relative_files(self, source_root: Path) -> List[Path]:
        files: List[Path] = []
        for root, _, names in os.walk(source_root):
            root_path = Path(root)
            for name in sorted(names, key=lambda n: n.lower()):
                file_path = root_path / name
                files.append(file_path.relative_to(source_root))
        return sorted(files, key=lambda p: p.as_posix().lower())

    @staticmethod
    def _strip_quotes(token: str) -> str:
        t = token.strip()
        if len(t) >= 2 and t[0] == '"' and t[-1] == '"':
            return t[1:-1]
        return t

    @staticmethod
    def _validate_relative_path(path_str: str) -> Optional[str]:
        if not path_str:
            return None
        normalized = path_str.replace("\\", "/").strip()
        if not normalized or normalized.startswith("/") or normalized.startswith("~") or ":" in normalized:
            return None
        parts = [p for p in normalized.split("/") if p not in ("", ".")]
        if not parts or any(p == ".." for p in parts):
            return None
        return Path(*parts).as_posix()

    @staticmethod
    def _is_within_root(path: Path, root: Path) -> bool:
        try:
            path.relative_to(root)
            return True
        except ValueError:
            return False

    @staticmethod
    def _read_text_excerpt(file_path: Path, max_chars: int = 700) -> Tuple[Optional[str], Optional[str]]:
        try:
            with open(file_path, "rb") as f:
                blob = f.read(4096)
            if b"\x00" in blob:
                return None, "binary or unsupported encoding"
            with open(file_path, "r", encoding="utf-8", errors="replace") as f:
                content = f.read(max_chars)
            excerpt = content.strip()
            if not excerpt:
                excerpt = "[empty file]"
            return excerpt, None
        except Exception as err:
            return None, str(err)

