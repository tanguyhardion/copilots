"""
PowerPoint Animation Engine: Constants, mappings, parsing, and COM automation helpers.
Designed to be modular and scalable for entrance, emphasis, and exit animations.
"""

from typing import Dict, Any, Optional, Tuple, List

# ── MSO ANIMATION EFFECT ENUMS ───────────────────────────────────────────────
# PowerPoint msoAnimEffect enum values
# Entrance effects
MSO_ANIM_EFFECT_MAP: Dict[str, int] = {
    "appear": 1,        # msoAnimEffectAppear
    "fly_in": 2,        # msoAnimEffectFly
    "blind": 3,         # msoAnimEffectBlinds
    "box": 4,           # msoAnimEffectBox
    "checkerboard": 5,  # msoAnimEffectCheckerboard
    "circle": 6,        # msoAnimEffectCircle
    "crawl": 7,         # msoAnimEffectCrawl
    "dissolve": 8,      # msoAnimEffectDissolve
    "flash_once": 9,    # msoAnimEffectFlashOnce
    "fade": 10,         # msoAnimEffectFade
    "diamond": 11,      # msoAnimEffectDiamond
    "split": 12,        # msoAnimEffectSplit
    "wedge": 13,        # msoAnimEffectWedge
    "wheel": 14,        # msoAnimEffectWheel
    "wipe": 15,         # msoAnimEffectWipe
    "zoom": 16,         # msoAnimEffectZoom
    "random_bars": 17,  # msoAnimEffectRandomBars
    "float_in": 74,     # msoAnimEffectFloat
    "expand": 63,       # msoAnimEffectExpand
}

REVERSE_ANIM_EFFECT_MAP: Dict[int, str] = {v: k for k, v in MSO_ANIM_EFFECT_MAP.items()}

# ── TRIGGER TYPES ────────────────────────────────────────────────────────────
# PowerPoint msoAnimTriggerType enum values
MSO_ANIM_TRIGGER_MAP: Dict[str, int] = {
    "on_click": 1,         # msoAnimTriggerOnPageClick
    "with_previous": 2,    # msoAnimTriggerWithPrevious
    "after_previous": 3,   # msoAnimTriggerAfterPrevious
}

REVERSE_ANIM_TRIGGER_MAP: Dict[int, str] = {v: k for k, v in MSO_ANIM_TRIGGER_MAP.items()}

# Aliases for trigger types
TRIGGER_ALIASES: Dict[str, str] = {
    "click": "on_click",
    "onclick": "on_click",
    "on_click": "on_click",
    "with_prev": "with_previous",
    "withprev": "with_previous",
    "with_previous": "with_previous",
    "after_prev": "after_previous",
    "afterprev": "after_previous",
    "after_previous": "after_previous",
}


def parse_animation_fields(fields: Dict[str, str]) -> Optional[Dict[str, Any]]:
    """
    Parses animation configuration from DSL key-value fields.
    Supports:
      anim=appear
      anim=appear,on_click
      anim=appear,after_previous,0.5
      animation=appear
      anim_trigger=on_click|with_previous|after_previous
      anim_delay=N (seconds)
      anim_duration=N (seconds)
      anim_order=N (int)
    """
    raw_anim = fields.get("anim") or fields.get("animation")
    if not raw_anim:
        return None

    effect_name = None
    trigger = "on_click"
    delay = 0.0
    duration = 0.0
    order = None

    # Handle comma-separated compact shorthand: anim=appear,after_previous,0.5
    parts = [p.strip() for p in raw_anim.split(",") if p.strip()]
    if parts:
        effect_name = parts[0].lower()
        if len(parts) > 1:
            raw_trig = parts[1].lower()
            trigger = TRIGGER_ALIASES.get(raw_trig, raw_trig)
        if len(parts) > 2:
            try:
                delay = float(parts[2])
            except ValueError:
                pass

    if not effect_name:
        return None

    # Normalise effect name if valid or keep unknown effect for robustness
    effect_key = effect_name if effect_name in MSO_ANIM_EFFECT_MAP else "appear"

    # Explicit field overrides
    if "anim_trigger" in fields:
        raw_trig = fields["anim_trigger"].strip().lower()
        trigger = TRIGGER_ALIASES.get(raw_trig, trigger)

    if trigger not in MSO_ANIM_TRIGGER_MAP:
        trigger = "on_click"

    if "anim_delay" in fields:
        try:
            delay = float(fields["anim_delay"])
        except ValueError:
            pass

    if "anim_duration" in fields:
        try:
            duration = float(fields["anim_duration"])
        except ValueError:
            pass

    if "anim_order" in fields:
        try:
            order = int(fields["anim_order"])
        except ValueError:
            pass

    return {
        "effect": effect_key,
        "trigger": trigger,
        "delay": max(0.0, delay),
        "duration": max(0.0, duration),
        "order": order,
    }


def apply_shape_animation(slide, ppt_shape, anim_config: Dict[str, Any]) -> Optional[Any]:
    """
    Applies the configured animation to the given PowerPoint shape on the slide timeline.
    Returns the created Effect object or None if failed.
    """
    if not ppt_shape or not anim_config:
        return None

    try:
        timeline = getattr(slide, "TimeLine", None)
        if not timeline:
            return None

        main_seq = timeline.MainSequence
        effect_key = anim_config.get("effect", "appear")
        effect_id = MSO_ANIM_EFFECT_MAP.get(effect_key, 1)  # Default to 1 (Appear)
        trigger_name = anim_config.get("trigger", "on_click")
        trigger_id = MSO_ANIM_TRIGGER_MAP.get(trigger_name, 1)  # Default to 1 (OnPageClick)

        # AddEffect(Shape, effectId, Level, trigger)
        # Level 0 = msoAnimateLevelNone (animate shape as a single unit)
        effect = main_seq.AddEffect(ppt_shape, effect_id, 0, trigger_id)

        timing = effect.Timing
        delay = anim_config.get("delay", 0.0)
        if delay > 0:
            timing.TriggerDelayTime = float(delay)

        duration = anim_config.get("duration", 0.0)
        if duration > 0:
            timing.Duration = float(duration)

        return effect
    except Exception as e:
        print(f"[animation] Error applying animation '{anim_config}': {e}")
        return None


def extract_slide_animations(slide) -> Dict[int, Dict[str, Any]]:
    """
    Scans slide.TimeLine.MainSequence and returns a dict mapping shape ID -> animation dict.
    """
    result: Dict[int, Dict[str, Any]] = {}
    try:
        timeline = getattr(slide, "TimeLine", None)
        if not timeline:
            return result

        main_seq = timeline.MainSequence
        count = main_seq.Count
        for i in range(1, count + 1):
            try:
                eff = main_seq(i)
                shape = getattr(eff, "Shape", None)
                if not shape:
                    continue
                shape_id = shape.Id
                # If shape already has an animation registered, keep the first primary entrance effect
                if shape_id in result:
                    continue

                eff_type = getattr(eff, "EffectType", 1)
                eff_name = REVERSE_ANIM_EFFECT_MAP.get(eff_type, "appear")

                timing = getattr(eff, "Timing", None)
                trigger_name = "on_click"
                delay = 0.0
                duration = 0.0

                if timing:
                    trig_type = getattr(timing, "TriggerType", 1)
                    trigger_name = REVERSE_ANIM_TRIGGER_MAP.get(trig_type, "on_click")
                    delay = round(float(getattr(timing, "TriggerDelayTime", 0.0)), 2)
                    duration = round(float(getattr(timing, "Duration", 0.0)), 2)

                result[shape_id] = {
                    "effect": eff_name,
                    "trigger": trigger_name,
                    "delay": delay,
                    "duration": duration,
                    "sequence_index": i,
                }
            except Exception as e:
                print(f"[animation] Error extracting effect index {i}: {e}")
    except Exception as e:
        print(f"[animation] Error accessing timeline for slide animations: {e}")

    return result


def format_animation_dsl(anim_config: Dict[str, Any]) -> str:
    """
    Formats an animation configuration dict into concise DSL tokens.
    """
    if not anim_config:
        return ""

    effect = anim_config.get("effect", "appear")
    trigger = anim_config.get("trigger", "on_click")
    delay = anim_config.get("delay", 0.0)

    # Shorthand or explicit tokens
    if trigger == "on_click" and delay == 0.0:
        return f"anim={effect}"
    elif delay == 0.0:
        return f"anim={effect} anim_trigger={trigger}"
    else:
        return f"anim={effect} anim_trigger={trigger} anim_delay={delay}"

