"""
PowerPoint Editor: Executes edit operations directly on the active PowerPoint presentation (active slide or specific slides).
Applies surgical in-place modifications, shape replacements, deletions, and additions via COM.
"""

from typing import List, Dict, Any, Optional, Callable
import pythoncom

from copilots_app.services.powerpoint.constants import (
    PPT_THEME_MAP,
    THEME_COLORS,
)
from copilots_app.services.powerpoint.parser import (
    dsl_resolve_color,
    is_theme_color,
    hex_to_rgb,
    rgb_to_bgr_int,
    parse_rich_text,
)
from copilots_app.services.powerpoint.connector import PowerPointConnector
from copilots_app.services.powerpoint.edit_parser import parse_edit_dsl
from copilots_app.services.powerpoint.animations import parse_animation_fields, apply_shape_animation


class PowerPointEditor:
    """Executes edit operations on the active PowerPoint slide or on specific slides."""

    def __init__(self, connector: Optional[PowerPointConnector] = None):
        self.connector = connector or PowerPointConnector()

    def apply_edits_to_active_slide(
        self,
        dsl_string: str,
        status_cb: Optional[Callable[[str], None]] = None,
    ) -> Dict[str, Any]:
        """
        Parses edit DSL and applies the operations in place. Operations target the active slide
        by default, or specific slides when grouped under `edit slide=N` headers.
        Returns a dict with success status, counts of ops performed, and a message.
        """
        pythoncom.CoInitialize()
        try:
            edit_ops = parse_edit_dsl(dsl_string)
            if not edit_ops:
                raise Exception("No valid edit operations found in the input DSL.")

            if status_cb:
                status_cb("Connecting to PowerPoint…")
            self.connector.connect()
            ppt = self.connector.ppt_app

            if ppt.Presentations.Count == 0:
                raise Exception("No presentations are open. Please open one in PowerPoint first.")

            # Resolve every targeted slide up front so an invalid index aborts before any change
            slides = {}
            for op in edit_ops:
                target = op.get("slide")
                if target not in slides:
                    slides[target] = self._resolve_slide(ppt, target)

            # Prefetch icons for all newly inserted/replaced shapes
            all_new_elements = []
            for op in edit_ops:
                if "elements" in op:
                    all_new_elements.extend(op["elements"])
            icon_cache = self.connector._prefetch_icons(all_new_elements, status_cb)

            counts = {"modify": 0, "replace": 0, "delete": 0, "insert": 0, "background": 0}
            total = len(edit_ops)
            touched_indices = []

            for idx, op in enumerate(edit_ops):
                op_type = op["op"]
                slide = slides[op.get("slide")]
                slide_index = slide.SlideIndex
                if slide_index not in touched_indices:
                    touched_indices.append(slide_index)
                if status_cb:
                    status_cb(f"Applying edit {idx+1}/{total} on slide {slide_index}: {op_type}…")
                self._apply_op(slide, op, icon_cache, counts)

            summary_parts = []
            if counts["modify"]:
                summary_parts.append(f"{counts['modify']} modified")
            if counts["replace"]:
                summary_parts.append(f"{counts['replace']} replaced")
            if counts["delete"]:
                summary_parts.append(f"{counts['delete']} deleted")
            if counts["insert"]:
                summary_parts.append(f"{counts['insert']} inserted")
            if counts["background"]:
                summary_parts.append(f"{counts['background']} background(s) updated")

            if len(touched_indices) == 1:
                target_label = f"slide {touched_indices[0]}"
            else:
                target_label = f"{len(touched_indices)} slides ({', '.join(str(n) for n in sorted(touched_indices))})"
            msg = f"✓ Edits applied to {target_label}: " + (", ".join(summary_parts) if summary_parts else "completed")
            if status_cb:
                status_cb(msg)

            return {
                "success": True,
                "message": msg,
                "counts": counts,
                "slide_index": touched_indices[0],
                "slide_indices": touched_indices,
            }
        finally:
            pythoncom.CoUninitialize()

    def _resolve_slide(self, ppt, target: Optional[int]):
        """Return the slide object for a 1-based index, or the active slide when target is None."""
        if target is None:
            try:
                return ppt.ActiveWindow.View.Slide
            except Exception:
                raise Exception("Could not get active slide. Please select a slide in PowerPoint first.")
        slide_count = ppt.ActivePresentation.Slides.Count
        if not 1 <= target <= slide_count:
            raise Exception(f"Slide {target} does not exist (presentation has {slide_count} slide(s)).")
        return ppt.ActivePresentation.Slides(target)

    def _apply_op(self, slide, op: Dict[str, Any], icon_cache, counts: Dict[str, int]):
        """Apply a single parsed edit operation to the given slide, updating counts."""
        op_type = op["op"]

        if op_type == "clear_slide":
            # Delete all shapes on the slide
            while slide.Shapes.Count > 0:
                slide.Shapes(1).Delete()

        elif op_type == "slide_background":
            bg_color = op.get("background_color")
            self.connector._apply_slide_background(slide, bg_color)
            counts["background"] += 1

        elif op_type == "delete":
            target_id = op.get("target_id")
            shape = self._find_shape(slide, target_id)
            if shape:
                shape.Delete()
                counts["delete"] += 1
            else:
                print(f"[editor] Could not find shape with id={target_id} on slide {slide.SlideIndex} to delete")

        elif op_type == "modify":
            target_id = op.get("target_id")
            shape = self._find_shape(slide, target_id)
            if shape:
                self._modify_shape(slide, shape, op.get("fields", {}), op.get("text_part"))
                counts["modify"] += 1
            else:
                print(f"[editor] Could not find shape with id={target_id} on slide {slide.SlideIndex} to modify")

        elif op_type == "replace":
            target_id = op.get("target_id")
            shape = self._find_shape(slide, target_id)
            if shape:
                shape.Delete()
            for elem in op.get("elements", []):
                self.connector._create_single_shape(slide, elem, icon_cache)
            counts["replace"] += 1

        elif op_type in ("insert_after", "insert_before", "insert_at"):
            for elem in op.get("elements", []):
                self.connector._create_single_shape(slide, elem, icon_cache)
            counts["insert"] += len(op.get("elements", []))

    def _find_shape(self, slide, target_id: Optional[int]):
        if target_id is None:
            return None
        try:
            for s in slide.Shapes:
                if s.Id == target_id:
                    return s
        except Exception as e:
            print(f"[editor] Error finding shape {target_id}: {e}")
        return None

    def _modify_shape(self, slide, shape, fields: Dict[str, str], text_part: Optional[str]):
        """Modify shape properties, animations, and text in-place."""
        # 1. Geometry
        if "left" in fields:
            try:
                shape.Left = float(fields["left"])
            except ValueError:
                pass
        if "top" in fields:
            try:
                shape.Top = float(fields["top"])
            except ValueError:
                pass
        if "width" in fields:
            try:
                shape.Width = float(fields["width"])
            except ValueError:
                pass
        if "height" in fields:
            try:
                shape.Height = float(fields["height"])
            except ValueError:
                pass
        if "rotation" in fields:
            try:
                shape.Rotation = float(fields["rotation"])
            except ValueError:
                pass

        # 2. Fill
        if "color" in fields:
            try:
                col = dsl_resolve_color(fields["color"])
                if col:
                    shape.Fill.Visible = True
                    if is_theme_color(col):
                        shape.Fill.ForeColor.ObjectThemeColor = PPT_THEME_MAP[col]
                    else:
                        r, g, b = hex_to_rgb(col)
                        shape.Fill.ForeColor.RGB = rgb_to_bgr_int(r, g, b)
                    shape.Fill.Solid()
            except Exception as e:
                print(f"[editor] Error modifying fill: {e}")

        if "transparency" in fields:
            try:
                shape.Fill.Transparency = float(fields["transparency"])
            except ValueError:
                pass

        # 3. Outline
        if "outline" in fields:
            try:
                val = fields["outline"].strip()
                if val.lower() == "none" or val == "false":
                    shape.Line.Visible = False
                else:
                    parts = val.split(",")
                    col = dsl_resolve_color(parts[0].strip())
                    weight = float(parts[1].strip()) if len(parts) > 1 else 1.0
                    shape.Line.Visible = True
                    shape.Line.Weight = weight
                    if col:
                        if is_theme_color(col):
                            shape.Line.ForeColor.ObjectThemeColor = PPT_THEME_MAP[col]
                        else:
                            r, g, b = hex_to_rgb(col)
                            shape.Line.ForeColor.RGB = rgb_to_bgr_int(r, g, b)
            except Exception as e:
                print(f"[editor] Error modifying outline: {e}")

        # 4. Text content
        if text_part is not None:
            try:
                rich = parse_rich_text(text_part)
                shape_def = {
                    "type": "rect",
                    "rich_text": rich,
                }
                if "halign" in fields:
                    shape_def["text_align"] = fields["halign"]
                if "valign" in fields:
                    shape_def["vertical_align"] = fields["valign"]
                if "font" in fields:
                    shape_def["font"] = fields["font"]
                if "bullet" in fields:
                    b_val = fields["bullet"].strip()
                    shape_def["bullet"] = True if b_val.lower() == "true" else b_val
                if "line_height" in fields:
                    shape_def["line_height"] = fields["line_height"]

                self.connector._apply_rich_text(shape, shape_def)
            except Exception as e:
                print(f"[editor] Error modifying text: {e}")

        # 5. Animation
        if "anim" in fields or "animation" in fields:
            anim_val = (fields.get("anim") or fields.get("animation") or "").strip().lower()
            if anim_val in ("none", "false", "0", "remove"):
                # Remove existing animation for this shape from slide timeline
                try:
                    timeline = getattr(slide, "TimeLine", None)
                    if timeline:
                        main_seq = timeline.MainSequence
                        for i in range(main_seq.Count, 0, -1):
                            eff = main_seq(i)
                            if getattr(eff, "Shape", None) and eff.Shape.Id == shape.Id:
                                eff.Delete()
                except Exception as e:
                    print(f"[editor] Error removing animation: {e}")
            else:
                anim_cfg = parse_animation_fields(fields)
                if anim_cfg:
                    try:
                        # Remove prior effect on this shape first to replace it cleanly
                        timeline = getattr(slide, "TimeLine", None)
                        if timeline:
                            main_seq = timeline.MainSequence
                            for i in range(main_seq.Count, 0, -1):
                                eff = main_seq(i)
                                if getattr(eff, "Shape", None) and eff.Shape.Id == shape.Id:
                                    eff.Delete()
                        apply_shape_animation(slide, shape, anim_cfg)
                    except Exception as e:
                        print(f"[editor] Error updating animation: {e}")

