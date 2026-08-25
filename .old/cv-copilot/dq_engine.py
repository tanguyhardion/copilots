"""
dq_engine.py
Deterministic DQ engine for the Europass CV Builder.

Severity:
  ERROR   – something is definitively wrong
  WARNING – worth surfacing but not necessarily blocking

Run order:
  1. JSON schema validation  (hard-stop on failure)
  2. Date format sanity
  3. All other deterministic rules
"""

from __future__ import annotations

import json
import re
from datetime import date

import jsonschema

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

_DATE_RE = re.compile(r"^\d{4}-\d{2}$")
_TODAY = date.today()


def _flag(
    source: str,
    section: str,
    rule_code: str,
    severity: str,
    message: str,
    entry: str | None = None,
    field: str | None = None,
) -> dict:
    return {
        "source": source,
        "section": section,
        "entry": entry,
        "field": field,
        "rule_code": rule_code,
        "severity": severity,
        "message": message,
    }


def _app_error(section, rule_code, message, entry=None, field=None):
    return _flag("APP", section, rule_code, "ERROR", message, entry, field)


def _app_warn(section, rule_code, message, entry=None, field=None):
    return _flag("APP", section, rule_code, "WARNING", message, entry, field)


def _parse_ym(value: str, today: date = _TODAY) -> date | None:
    """
    Parse a YYYY-MM string or 'Present' into a date (first of month).
    Returns None if unparseable.
    """
    if value == "Present":
        return today.replace(day=1)
    if not _DATE_RE.match(value):
        return None
    year, month = int(value[:4]), int(value[5:7])
    if not (1 <= month <= 12):
        return None
    return date(year, month, 1)


def _duration_inclusive(d_from: date, d_to: date) -> int:
    """Inclusive month count: Jun-25 → Sep-25 = 4."""
    return (d_to.year - d_from.year) * 12 + (d_to.month - d_from.month) + 1


def _next_month(d: date) -> date:
    """First day of the month after d."""
    if d.month == 12:
        return date(d.year + 1, 1, 1)
    return date(d.year, d.month + 1, 1)


def _entry_label(item: dict, idx: int) -> str:
    """Human-readable entry identifier for flag messages."""
    for key in ("project_title", "job_title", "qualification_title", "role"):
        if item.get(key):
            return item[key]
    return f"entry #{idx + 1}"


# ---------------------------------------------------------------------------
# Rule implementations
# ---------------------------------------------------------------------------


def _check_date_formats(cv: dict) -> list[dict]:
    """
    DATE_FORMAT_INVALID – YYYY-MM values that pass the regex but have
    month 00 or 13+, or any other non-Present string that doesn't match
    the pattern at all.
    """
    flags = []
    sections = {
        "work_experience": ("date_from", "date_to"),
        "education": ("date_from", "date_to"),
        "experience_overview": ("date_from", "date_to"),
        "project_experience": ("date_from", "date_to"),
    }
    for section, fields in sections.items():
        for idx, item in enumerate(cv.get(section) or []):
            label = _entry_label(item, idx)
            for field in fields:
                val = item.get(field)
                if val is None:
                    continue
                if val == "Present":
                    continue
                if not _DATE_RE.match(str(val)):
                    flags.append(_app_error(
                        section, "DATE_FORMAT_INVALID",
                        f"'{field}' value '{val}' is not a valid YYYY-MM date.",
                        entry=label, field=field,
                    ))
                    continue
                month = int(val[5:7])
                if not (1 <= month <= 12):
                    flags.append(_app_error(
                        section, "DATE_FORMAT_INVALID",
                        f"'{field}' value '{val}' has an invalid month ({month}).",
                        entry=label, field=field,
                    ))
    return flags


def _check_date_logic(cv: dict) -> list[dict]:
    """
    DATE_FROM_AFTER_DATE_TO  – date_from > date_to
    DATE_IN_FUTURE           – non-Present date_to is after today
    """
    flags = []
    today_first = _TODAY.replace(day=1)
    sections = {
        "work_experience": ("date_from", "date_to"),
        "education": ("date_from", "date_to"),
        "experience_overview": ("date_from", "date_to"),
        "project_experience": ("date_from", "date_to"),
    }
    for section, (f_from, f_to) in sections.items():
        for idx, item in enumerate(cv.get(section) or []):
            label = _entry_label(item, idx)
            raw_from = item.get(f_from)
            raw_to = item.get(f_to)

            d_from = _parse_ym(raw_from) if raw_from else None
            d_to = _parse_ym(raw_to) if raw_to else None

            # Skip if either date couldn't be parsed (DATE_FORMAT_INVALID
            # will already have fired for those)
            if d_from is None or d_to is None:
                continue

            if d_from > d_to:
                flags.append(_app_error(
                    section, "DATE_FROM_AFTER_DATE_TO",
                    f"date_from ({raw_from}) is after date_to ({raw_to}).",
                    entry=label, field=f_from,
                ))

            if raw_to != "Present" and d_to > today_first:
                flags.append(_app_error(
                    section, "DATE_IN_FUTURE",
                    f"date_to ({raw_to}) is in the future.",
                    entry=label, field=f_to,
                ))
    return flags


def _check_employment_overlap(cv: dict) -> list[dict]:
    """
    EMPLOYMENT_OVERLAP  – two work_experience entries share any month
    MULTIPLE_PRESENT    – more than one work_experience entry has date_to = 'Present'
    """
    flags = []
    entries = cv.get("work_experience") or []

    # Collect parseable entries
    parsed: list[tuple[date, date, str]] = []  # (from, to, label)
    present_count = 0

    for idx, item in enumerate(entries):
        label = _entry_label(item, idx)
        raw_from = item.get("date_from")
        raw_to = item.get("date_to")
        if raw_to == "Present":
            present_count += 1
        d_from = _parse_ym(raw_from) if raw_from else None
        d_to = _parse_ym(raw_to) if raw_to else None
        if d_from and d_to:
            parsed.append((d_from, d_to, label))

    if present_count > 1:
        flags.append(_app_error(
            "work_experience", "MULTIPLE_PRESENT",
            f"{present_count} work_experience entries have date_to = 'Present'. "
            "Only one is allowed unless allocation_percent is specified "
            "(allocation is not a field on work_experience — use project_experience).",
        ))

    # Check every pair for overlap (shared boundary month counts as overlap)
    for i in range(len(parsed)):
        for j in range(i + 1, len(parsed)):
            a_from, a_to, a_label = parsed[i]
            b_from, b_to, b_label = parsed[j]
            # Overlap if ranges share at least one month
            overlap_start = max(a_from, b_from)
            overlap_end = min(a_to, b_to)
            if overlap_start <= overlap_end:
                flags.append(_app_error(
                    "work_experience", "EMPLOYMENT_OVERLAP",
                    f"Entries '{a_label}' and '{b_label}' overlap "
                    f"(shared period starts {overlap_start.strftime('%Y-%m')}).",
                ))
    return flags


def _check_work_experience_gaps(cv: dict) -> list[dict]:
    """
    WORK_EXP_GAP – any gap of ≥1 month between consecutive work_experience
    entries (when sorted by date_from).
    Allowed: entry A ends YYYY-MM, entry B starts the very next month.
    """
    flags = []
    entries = cv.get("work_experience") or []

    parsed: list[tuple[date, date, str]] = []
    for idx, item in enumerate(entries):
        label = _entry_label(item, idx)
        d_from = _parse_ym(item.get("date_from", ""))
        d_to = _parse_ym(item.get("date_to", ""))
        if d_from and d_to:
            parsed.append((d_from, d_to, label))

    if len(parsed) < 2:
        return flags

    parsed.sort(key=lambda x: x[0])

    for i in range(len(parsed) - 1):
        _, a_to, a_label = parsed[i]
        b_from, _, b_label = parsed[i + 1]
        expected_next = _next_month(a_to)
        if b_from > expected_next:
            gap_months = (b_from.year - a_to.year) * 12 + (b_from.month - a_to.month) - 1
            flags.append(_app_warn(
                "work_experience", "WORK_EXP_GAP",
                f"Gap of {gap_months} month(s) between '{a_label}' "
                f"(ends {a_to.strftime('%Y-%m')}) and '{b_label}' "
                f"(starts {b_from.strftime('%Y-%m')}).",
            ))
    return flags


def _check_duration(cv: dict) -> list[dict]:
    """
    DURATION_MISMATCH        – duration_months doesn't match inclusive calculation
    RELEVANT_EXCEEDS_DURATION – relevant_months > duration_months
    RELEVANT_MONTHS_ZERO     – relevant_months = 0
    """
    flags = []
    for idx, item in enumerate(cv.get("experience_overview") or []):
        label = _entry_label(item, idx)
        raw_from = item.get("date_from")
        raw_to = item.get("date_to")
        stated = item.get("duration_months")
        relevant = item.get("relevant_months")

        d_from = _parse_ym(raw_from) if raw_from else None
        d_to = _parse_ym(raw_to) if raw_to else None

        if d_from and d_to and stated is not None:
            expected = _duration_inclusive(d_from, d_to)
            if stated != expected:
                flags.append(_app_error(
                    "experience_overview", "DURATION_MISMATCH",
                    f"'{label}': duration_months is {stated} but calculated "
                    f"inclusive duration is {expected} "
                    f"({raw_from} → {raw_to}).",
                    entry=label, field="duration_months",
                ))

        if stated is not None and relevant is not None:
            if relevant > stated:
                flags.append(_app_error(
                    "experience_overview", "RELEVANT_EXCEEDS_DURATION",
                    f"'{label}': relevant_months ({relevant}) exceeds "
                    f"duration_months ({stated}).",
                    entry=label, field="relevant_months",
                ))
            if relevant == 0:
                flags.append(_app_warn(
                    "experience_overview", "RELEVANT_MONTHS_ZERO",
                    f"'{label}': relevant_months is 0. "
                    "Verify this entry should be included.",
                    entry=label, field="relevant_months",
                ))
    return flags


def _check_overview_project_mismatch(cv: dict) -> list[dict]:
    """
    EXP_OVERVIEW_PROJECT_MISMATCH – count, dates, or role don't align
    between experience_overview and project_experience.
    One flag per discrepancy found.
    """
    flags = []
    overview = cv.get("experience_overview") or []
    projects = cv.get("project_experience") or []

    if len(overview) != len(projects):
        flags.append(_app_error(
            "experience_overview", "EXP_OVERVIEW_PROJECT_MISMATCH",
            f"experience_overview has {len(overview)} entries but "
            f"project_experience has {len(projects)}. They must be parallel arrays.",
        ))
        # Can't do pair-wise checks if counts differ
        return flags

    for idx, (ov, pr) in enumerate(zip(overview, projects)):
        ov_label = _entry_label(ov, idx)
        pr_label = _entry_label(pr, idx)
        pair = f"pair #{idx + 1} (overview: '{ov_label}' / project: '{pr_label}')"

        if ov.get("date_from") != pr.get("date_from"):
            flags.append(_app_error(
                "experience_overview", "EXP_OVERVIEW_PROJECT_MISMATCH",
                f"{pair}: date_from mismatch — "
                f"overview '{ov.get('date_from')}' vs project '{pr.get('date_from')}'.",
                entry=pair, field="date_from",
            ))

        if ov.get("date_to") != pr.get("date_to"):
            flags.append(_app_error(
                "experience_overview", "EXP_OVERVIEW_PROJECT_MISMATCH",
                f"{pair}: date_to mismatch — "
                f"overview '{ov.get('date_to')}' vs project '{pr.get('date_to')}'.",
                entry=pair, field="date_to",
            ))

        if ov.get("role") != pr.get("role"):
            flags.append(_app_error(
                "experience_overview", "EXP_OVERVIEW_PROJECT_MISMATCH",
                f"{pair}: role mismatch — "
                f"overview '{ov.get('role')}' vs project '{pr.get('role')}'.",
                entry=pair, field="role",
            ))
    return flags


def _check_parallel_projects(cv: dict) -> list[dict]:
    """
    PARALLEL_PROJECT_NO_ALLOCATION – overlapping projects where at least
                                     one has no allocation_percent
    ALLOCATION_EXCEEDS_100         – overlapping projects whose numeric
                                     allocations sum to > 100 %
    ALLOCATION_UNPARSEABLE         – allocation_percent present but not
                                     a parseable percentage string
    """
    flags = []
    projects = cv.get("project_experience") or []

    parsed: list[tuple[date, date, str, str | None]] = []  # from, to, label, alloc_raw
    for idx, item in enumerate(projects):
        label = _entry_label(item, idx)
        d_from = _parse_ym(item.get("date_from", ""))
        d_to = _parse_ym(item.get("date_to", ""))
        alloc = item.get("allocation_percent")
        if d_from and d_to:
            parsed.append((d_from, d_to, label, alloc))

    for i in range(len(parsed)):
        for j in range(i + 1, len(parsed)):
            a_from, a_to, a_label, a_alloc = parsed[i]
            b_from, b_to, b_label, b_alloc = parsed[j]

            overlap_start = max(a_from, b_from)
            overlap_end = min(a_to, b_to)
            if overlap_start > overlap_end:
                continue  # no overlap

            pair = f"'{a_label}' and '{b_label}'"

            # --- missing allocation ---
            if not a_alloc or not b_alloc:
                missing = []
                if not a_alloc:
                    missing.append(a_label)
                if not b_alloc:
                    missing.append(b_label)
                flags.append(_app_error(
                    "project_experience", "PARALLEL_PROJECT_NO_ALLOCATION",
                    f"Projects {pair} overlap (from "
                    f"{overlap_start.strftime('%Y-%m')}) but "
                    f"{' and '.join(missing)} has no allocation_percent.",
                ))
                continue  # skip sum check if we can't parse both

            # --- parse allocations ---
            def _parse_alloc(raw: str, label: str) -> float | None:
                cleaned = raw.strip().rstrip("%").strip()
                try:
                    return float(cleaned)
                except ValueError:
                    flags.append(_app_warn(
                        "project_experience", "ALLOCATION_UNPARSEABLE",
                        f"'{label}': allocation_percent '{raw}' cannot be "
                        "parsed as a percentage.",
                        entry=label, field="allocation_percent",
                    ))
                    return None

            a_val = _parse_alloc(a_alloc, a_label)
            b_val = _parse_alloc(b_alloc, b_label)

            if a_val is not None and b_val is not None:
                total = a_val + b_val
                if total > 100:
                    flags.append(_app_warn(
                        "project_experience", "ALLOCATION_EXCEEDS_100",
                        f"Projects {pair} overlap and their combined "
                        f"allocation is {total:.0f}% "
                        f"({a_alloc} + {b_alloc}).",
                    ))
    return flags


def _check_cefr(cv: dict) -> list[dict]:
    """CEFR_INVALID – belt-and-suspenders over schema enum."""
    flags = []
    valid = {"A1", "A2", "B1", "B2", "C1", "C2"}
    fields = ["listening", "reading", "spoken_interaction",
              "spoken_production", "writing"]
    for idx, lang in enumerate(cv.get("languages", {}).get("other") or []):
        name = lang.get("language", f"language #{idx + 1}")
        for field in fields:
            val = lang.get(field)
            if val and val not in valid:
                flags.append(_app_error(
                    "languages", "CEFR_INVALID",
                    f"'{name}' has invalid CEFR level '{val}' for '{field}'. "
                    f"Must be one of {sorted(valid)}.",
                    entry=name, field=field,
                ))
    return flags


def _check_certifications(cv: dict) -> list[dict]:
    """
    CERT_EXPIRED    – expiry_year < current year  (WARNING)
    CERT_EXPIRY_SOON – expiry_year == current year or next year (WARNING)
    """
    flags = []
    this_year = _TODAY.year
    for idx, cert in enumerate(
        cv.get("personal_skills", {}).get("certifications") or []
    ):
        title = cert.get("title", f"cert #{idx + 1}")
        expiry = cert.get("expiry_year")
        if expiry is None:
            continue
        if expiry < this_year:
            flags.append(_app_warn(
                "personal_skills", "CERT_EXPIRED",
                f"Certification '{title}' expired in {expiry}.",
                entry=title, field="expiry_year",
            ))
        elif expiry <= this_year + 1:
            flags.append(_app_warn(
                "personal_skills", "CERT_EXPIRY_SOON",
                f"Certification '{title}' expires in {expiry} "
                "(within the next 12 months).",
                entry=title, field="expiry_year",
            ))
    return flags


def _check_project_bullets(cv: dict) -> list[dict]:
    """
    EMPTY_BULLETS  – project with zero bullets
    SINGLE_BULLET  – project with only one bullet
    """
    flags = []
    for idx, item in enumerate(cv.get("project_experience") or []):
        label = _entry_label(item, idx)
        bullets = item.get("bullets") or []
        if len(bullets) == 0:
            flags.append(_app_warn(
                "project_experience", "EMPTY_BULLETS",
                f"'{label}' has no bullet points.",
                entry=label, field="bullets",
            ))
        elif len(bullets) == 1:
            flags.append(_app_warn(
                "project_experience", "SINGLE_BULLET",
                f"'{label}' has only one bullet point — consider expanding.",
                entry=label, field="bullets",
            ))
    return flags


def _check_profile(cv: dict) -> list[dict]:
    """PROFILE_MISSING – profile key absent or empty."""
    profile = cv.get("profile", "")
    if not profile or not profile.strip():
        return [_app_warn(
            "profile", "PROFILE_MISSING",
            "The 'profile' section is absent or empty.",
        )]
    return []


def _check_no_project_experience(cv: dict) -> list[dict]:
    """NO_PROJECT_EXPERIENCE – project_experience array is empty."""
    if not cv.get("project_experience"):
        return [_app_warn(
            "project_experience", "NO_PROJECT_EXPERIENCE",
            "No project experience entries found. "
            "This is unusual for a proposal CV.",
        )]
    return []


def _check_missing_mandatory_fields(cv: dict) -> list[dict]:
    """
    MISSING_MANDATORY_FIELD – belt-and-suspenders for required fields
    that schema marks as required but may still be empty strings.
    """
    flags = []

    def _chk(section: str, item: dict, field: str, label: str):
        val = item.get(field)
        if val is None or (isinstance(val, str) and not val.strip()):
            flags.append(_app_error(
                section, "MISSING_MANDATORY_FIELD",
                f"Required field '{field}' is missing or empty on '{label}'.",
                entry=label, field=field,
            ))

    pi = cv.get("personal_info") or {}
    for f in ("first_name", "last_name", "nationality"):
        _chk("personal_info", pi, f, "personal_info")

    for idx, item in enumerate(cv.get("work_experience") or []):
        label = _entry_label(item, idx)
        for f in ("date_from", "date_to", "job_title", "organisation"):
            _chk("work_experience", item, f, label)

    for idx, item in enumerate(cv.get("education") or []):
        label = _entry_label(item, idx)
        for f in ("date_from", "date_to", "qualification_title", "institution"):
            _chk("education", item, f, label)

    for idx, item in enumerate(cv.get("experience_overview") or []):
        label = _entry_label(item, idx)
        for f in ("date_from", "date_to", "role", "responsibilities",
                  "duration_months", "relevant_months"):
            _chk("experience_overview", item, f, label)

    for idx, item in enumerate(cv.get("project_experience") or []):
        label = _entry_label(item, idx)
        for f in ("date_from", "date_to", "project_title", "role", "bullets"):
            _chk("project_experience", item, f, label)

    langs = cv.get("languages") or {}
    if not langs.get("mother_tongue"):
        flags.append(_app_error(
            "languages", "MISSING_MANDATORY_FIELD",
            "Required field 'mother_tongue' is missing or empty.",
            field="mother_tongue",
        ))

    return flags


# ---------------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------------

RULE_PIPELINE = [
    _check_missing_mandatory_fields,
    _check_date_formats,
    _check_date_logic,
    _check_employment_overlap,
    _check_work_experience_gaps,
    _check_duration,
    _check_overview_project_mismatch,
    _check_parallel_projects,
    _check_cefr,
    _check_certifications,
    _check_project_bullets,
    _check_profile,
    _check_no_project_experience,
]


def run_dq(cv_json: dict, schema: dict) -> dict:
    """
    Run the full DQ pipeline on a parsed CV dict.

    Returns a result dict:
    {
        "passed":    bool,          # False if any ERROR flag exists
        "flags":     [...],         # all flags (AI + APP), sorted by severity
        "cv":        {...}          # original cv_json unchanged
    }

    Hard-stops on JSON_SCHEMA_INVALID.
    """
    app_flags: list[dict] = []

    # ── Step 1: JSON Schema validation (hard stop) ──────────────────────────
    try:
        validator = jsonschema.Draft7Validator(schema)
        errors = sorted(validator.iter_errors(cv_json), key=str)
        if errors:
            for err in errors:
                path = " → ".join(str(p) for p in err.absolute_path) or "(root)"
                app_flags.append(_app_error(
                    path, "JSON_SCHEMA_INVALID",
                    f"{err.message} (path: {path})",
                ))
            return {
                "passed": False,
                "flags": app_flags,
                "cv": cv_json,
            }
    except Exception as exc:  # jsonschema itself broken
        app_flags.append(_app_error(
            "(schema)", "JSON_SCHEMA_INVALID",
            f"Schema validation could not run: {exc}",
        ))
        return {"passed": False, "flags": app_flags, "cv": cv_json}

    # ── Step 2: Deterministic rules ─────────────────────────────────────────
    for rule_fn in RULE_PIPELINE:
        app_flags.extend(rule_fn(cv_json))

    # ── Step 3: Merge with AI flags ─────────────────────────────────────────
    ai_flags = cv_json.get("dq_flags") or []
    all_flags = ai_flags + app_flags

    # Sort: ERRORs first, then WARNINGs, then INFO
    _order = {"ERROR": 0, "WARNING": 1, "INFO": 2}
    all_flags.sort(key=lambda f: _order.get(f.get("severity", "INFO"), 2))

    passed = not any(f.get("severity") == "ERROR" for f in all_flags)

    return {
        "passed": passed,
        "flags": all_flags,
        "cv": cv_json,
    }


# ---------------------------------------------------------------------------
# CLI convenience
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import sys

    if len(sys.argv) < 3:
        print("Usage: python dq_engine.py <cv.json> <schema.json>")
        sys.exit(1)

    with open(sys.argv[1], encoding="utf-8") as f:
        cv_data = json.load(f)
    with open(sys.argv[2], encoding="utf-8") as f:
        schema_data = json.load(f)

    result = run_dq(cv_data, schema_data)

    print(f"\n{'='*60}")
    print(f"  DQ Result: {'✅ PASSED' if result['passed'] else '❌ FAILED'}")
    print(f"  Total flags: {len(result['flags'])}")
    print(f"{'='*60}\n")

    for flag in result["flags"]:
        sev = flag["severity"]
        icon = {"ERROR": "🔴", "WARNING": "🟡", "INFO": "🔵"}.get(sev, "⚪")
        src = flag.get("source", "?")
        code = flag.get("rule_code", "?")
        section = flag.get("section", "?")
        entry = flag.get("entry")
        msg = flag.get("message", "")

        header = f"{icon} [{src}] {code}  |  {section}"
        if entry:
            header += f"  |  {entry}"
        print(header)
        print(f"   {msg}")
        print()