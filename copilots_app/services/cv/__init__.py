"""
CV Copilot Service Facade.
"""

from copilots_app.services.cv.generator import generate_cv
from copilots_app.services.cv.dq_engine import (
    _check_date_formats,
    _flag,
)

from typing import Dict, Any, List, Tuple


def run_dq_audit(cv_data: Dict[str, Any]) -> List[Dict[str, Any]]:
    """Run full deterministic Data Quality checks on CV JSON payload."""
    from copilots_app.services.cv import dq_engine

    flags = []

    # 1. Date format sanity
    try:
        flags.extend(dq_engine._check_date_formats(cv_data))
    except Exception as e:
        flags.append(_flag("DQ", "General", "DATE_CHECK_FAIL", "WARNING", str(e)))

    # 2. Check personal info fields
    pi = cv_data.get("personal_info", {})
    for field in ("first_name", "last_name", "email", "phone"):
        if not pi.get(field):
            flags.append(_flag("APP", "personal_info", "MISSING_MANDATORY_FIELD", "ERROR", f"Missing {field}", field=field))

    # 3. Check tender info
    ti = cv_data.get("tender_info", {})
    if not ti.get("proposed_role"):
        flags.append(_flag("APP", "tender_info", "MISSING_ROLE", "WARNING", "No proposed role specified"))

    # 4. Check work experience
    work_exp = cv_data.get("work_experience", [])
    if not work_exp:
        flags.append(_flag("APP", "work_experience", "NO_EXPERIENCE", "WARNING", "No work experience entries listed"))
    else:
        for idx, item in enumerate(work_exp):
            if not item.get("job_title"):
                flags.append(_flag("APP", "work_experience", "MISSING_JOB_TITLE", "ERROR", "Job title is empty", entry=f"#{idx+1}"))
            if not item.get("organisation"):
                flags.append(_flag("APP", "work_experience", "MISSING_ORG", "WARNING", "Organisation is empty", entry=f"#{idx+1}"))

    # 5. Check education
    edu = cv_data.get("education", [])
    if not edu:
        flags.append(_flag("APP", "education", "NO_EDUCATION", "WARNING", "No education entries listed"))

    # 6. Check skills
    skills = cv_data.get("personal_skills", {})
    if not skills.get("mother_tongues") and not skills.get("other_languages"):
        flags.append(_flag("APP", "personal_skills", "NO_LANGUAGES", "WARNING", "No language proficiency entered"))

    return flags
