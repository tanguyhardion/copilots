"""
PowerPoint Edit DSL Parser: Parses edit-mode operations targeting existing shapes.
Supports modify, replace, delete, insert_after, insert_before, and slide background directives.
Operations can target the active slide or specific slides (`edit slide=N`), enabling multi-slide edits.
"""

import re
from typing import List, Dict, Any, Tuple, Optional

from copilots_app.services.powerpoint.parser import (
    tokenize_dsl_line,
    extract_fields,
    parse_dsl,
    build_shape_from_fields,
    build_slide_command,
    dsl_resolve_color,
)

ID_COMMENT_RE = re.compile(r"//\s*id=(\d+)\s*$")
# Slide headers emitted by the extractor, e.g. `// === Slide 3/12 ===` or
# `// === Extracted from Active Slide (3/12) ===`
SLIDE_HEADER_RE = re.compile(r"^//\s*===\s*(?:Slide\s+|Extracted from Active Slide\s*\()(\d+)\s*/", re.IGNORECASE)


def _parse_slide_target(fields: Dict[str, str]) -> Optional[int]:
    """
    Resolve the target slide of an `edit` header directive.
    `edit target=active` -> None (active slide), `edit slide=N` / `edit target=N` -> N.
    """
    raw = fields.get("slide") or fields.get("target") or "active"
    raw = raw.strip().lower()
    if raw == "active":
        return None
    try:
        return int(raw)
    except ValueError:
        print(f"[edit-dsl] invalid edit target '{raw}', defaulting to active slide")
        return None


def looks_like_edit_dsl(dsl_string: str) -> bool:
    """
    Checks whether the DSL string contains edit-mode commands.
    """
    for line in dsl_string.strip().split("\n"):
        line = line.strip()
        if not line or line.startswith("//") or line.startswith("#"):
            continue
        lower = line.lower()
        if lower.startswith("edit") or lower.startswith("modify") or lower.startswith("delete id=") or lower.startswith("replace id=") or lower.startswith("insert_after") or lower.startswith("insert_before") or lower.startswith("clear slide"):
            return True
    return False


def _read_block_until_endblock(raw_lines: List[str], i: int) -> Tuple[List[str], int]:
    """
    Starting just after an op line, collect raw lines until 'endblock' (case-insensitive)
    or until the next edit op keyword. Returns (block_lines, next_index).
    """
    block_lines = []
    edit_op_starters = ("delete", "replace", "modify", "insert_before", "insert_after", "insert_at", "clear", "edit")
    while i < len(raw_lines):
        stripped = raw_lines[i].strip()
        if stripped.lower() == "endblock":
            return block_lines, i + 1
        # In case the LLM forgot 'endblock' and started the next operation:
        first_word = stripped.split()[0].lower() if stripped else ""
        if first_word in edit_op_starters and "=" in stripped:
            print("[edit-dsl] notice: implicit endblock detected before next operation")
            return block_lines, i
        block_lines.append(raw_lines[i])
        i += 1
    return block_lines, i


def parse_edit_dsl(dsl_string: str) -> List[Dict[str, Any]]:
    """
    Parses edit-mode DSL into a structured list of operation dictionaries.
    Every operation carries a "slide" key: the 1-based slide index it targets, or None for
    the active slide. The target is set by `edit slide=N` / `edit target=active` headers or by
    extractor slide header comments, and applies to all following operations.
      {"op": "modify", "target_id": N, "fields": {...}, "text_part": "..."}
      {"op": "replace", "target_id": N, "elements": [...]}
      {"op": "delete", "target_id": N}
      {"op": "insert_after", "target_id": N, "elements": [...]}
      {"op": "insert_before", "target_id": N, "elements": [...]}
      {"op": "insert_at", "position": "end", "elements": [...]}
      {"op": "slide_background", "background_color": "#HEX"}
      {"op": "clear_slide"}
    """
    raw_lines = dsl_string.strip().split("\n")
    ops = []
    current_slide: Optional[int] = None

    def emit(op: Dict[str, Any]) -> None:
        op["slide"] = current_slide
        ops.append(op)

    i = 0

    while i < len(raw_lines):
        raw = raw_lines[i]
        line = raw.strip()

        header_match = SLIDE_HEADER_RE.match(line)
        if header_match:
            current_slide = int(header_match.group(1))
            i += 1
            continue

        if not line or line.startswith("//") or line.startswith("#") or line == "---":
            i += 1
            continue

        tokens, text_part = tokenize_dsl_line(line)
        if not tokens:
            i += 1
            continue

        op_type, fields = extract_fields(tokens)
        if not op_type:
            i += 1
            continue

        # Header directive: `edit target=active` or `edit slide=N`
        if op_type == "edit":
            current_slide = _parse_slide_target(fields)
            i += 1
            continue

        # Clear slide directive: `clear slide`
        if op_type == "clear" and (fields.get("slide") or (len(tokens) > 1 and tokens[1].lower() == "slide")):
            emit({"op": "clear_slide"})
            i += 1
            continue

        # Slide background directive: `slide background=...`
        if op_type == "slide":
            slide_cmd = build_slide_command(fields, i + 1)
            if slide_cmd:
                emit({
                    "op": "slide_background",
                    "background_color": slide_cmd["background_color"],
                })
            i += 1
            continue

        # Delete operation: `delete id=N`
        if op_type == "delete":
            target_id = None
            if "id" in fields:
                try:
                    target_id = int(fields["id"])
                except ValueError:
                    pass
            if target_id is not None:
                emit({"op": "delete", "target_id": target_id})
            else:
                print(f"[edit-dsl] 'delete' missing valid id= on line {i+1}")
            i += 1
            continue

        # Modify operation: `modify id=N [left=...] [color=...] [| "new text"...]`
        if op_type == "modify":
            target_id = None
            if "id" in fields:
                try:
                    target_id = int(fields["id"])
                except ValueError:
                    pass
            if target_id is not None:
                emit({
                    "op": "modify",
                    "target_id": target_id,
                    "fields": fields,
                    "text_part": text_part,
                    "_source_line": i + 1,
                })
            else:
                print(f"[edit-dsl] 'modify' missing valid id= on line {i+1}")
            i += 1
            continue

        # Replace operation: `replace id=N` ... `endblock`
        if op_type == "replace":
            target_id = None
            if "id" in fields:
                try:
                    target_id = int(fields["id"])
                except ValueError:
                    pass

            # Read block
            block_lines, i = _read_block_until_endblock(raw_lines, i + 1)
            elements = parse_dsl("\n".join(block_lines))
            if target_id is not None:
                emit({"op": "replace", "target_id": target_id, "elements": elements})
            else:
                print(f"[edit-dsl] 'replace' missing valid id= on line {i}")
            continue

        # Insert operations: `insert_after id=N`, `insert_before id=N`
        if op_type in ("insert_after", "insert_before"):
            target_id = None
            if "id" in fields:
                try:
                    target_id = int(fields["id"])
                except ValueError:
                    pass

            block_lines, i = _read_block_until_endblock(raw_lines, i + 1)
            elements = parse_dsl("\n".join(block_lines))
            if target_id is not None:
                emit({"op": op_type, "target_id": target_id, "elements": elements})
            else:
                print(f"[edit-dsl] '{op_type}' missing valid id= on line {i}")
            continue

        # Insert at: `insert_at position=end`
        if op_type == "insert_at":
            position = fields.get("position") or (
                tokens[1].lower() if len(tokens) > 1 and "=" not in tokens[1] else "end"
            )
            block_lines, i = _read_block_until_endblock(raw_lines, i + 1)
            elements = parse_dsl("\n".join(block_lines))
            emit({"op": "insert_at", "position": position, "elements": elements})
            continue

        if op_type == "endblock":
            i += 1
            continue

        # If a standard shape line has an explicit id=N, treat as a replace or modify!
        if "id" in fields:
            try:
                target_id = int(fields["id"])
                shape_def = build_shape_from_fields(op_type, fields, text_part, i + 1)
                emit({"op": "replace", "target_id": target_id, "elements": [shape_def]})
                i += 1
                continue
            except ValueError:
                pass

        # Check comment for // id=N
        id_match = ID_COMMENT_RE.search(raw)
        if id_match:
            target_id = int(id_match.group(1))
            shape_def = build_shape_from_fields(op_type, fields, text_part, i + 1)
            emit({"op": "replace", "target_id": target_id, "elements": [shape_def]})
            i += 1
            continue

        # Default fallback: regular shape insertion
        single_shape = parse_dsl(raw)
        if single_shape:
            emit({"op": "insert_at", "position": "end", "elements": single_shape})
        i += 1

    return ops

