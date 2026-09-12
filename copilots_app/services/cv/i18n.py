from __future__ import annotations

from typing import Any, Dict, Optional

DEFAULT_CV_LANGUAGE = "en"

DOCX_SECTION_DEFAULTS = {
    "personal_info": True,
    "proposed_role": True,
    "professional_experience": True,
    "work_experience": True,
    "profile": True,
    "requirements_matrix": True,
    "experience_overview": True,
    "project_experience": True,
    "education": True,
    "personal_skills": True,
}

PPTX_SECTION_DEFAULTS = {
    "header_contact": True,
    "profile": True,
    "skills": True,
    "relevant_experience": True,
}

CV_STRINGS = {
    "en": {
        "docx_months": ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
        "pptx_months": ["Jan.", "Feb.", "Mar.", "Apr.", "May", "Jun.", "Jul.", "Aug.", "Sep.", "Oct.", "Nov.", "Dec."],
        "present": "Present",
        "europass": "Europass",
        "curriculum_vitae": "Curriculum Vitae",
        "page": "Page ",
        "personal_information": "PERSONAL INFORMATION",
        "photo_placeholder": "[Photo 3.8×3.8cm]",
        "sex_nationality": "Sex {sex} | Nationality {nationality}",
        "proposed_role": "PROPOSED ROLE",
        "professional_experience": "PROFESSIONAL EXPERIENCE",
        "years_of_work_experience": "Number of years of work experience:",
        "experience_buckets": ["Less than 4", "4 – 9", "10 – 14", "15+"],
        "work_experience": "WORK EXPERIENCE",
        "date": "DATE",
        "title": "TITLE",
        "organisation": "ORGANISATION",
        "profile": "PROFILE",
        "requirements_matrix": "Requirements matrix",
        "requirement": "Requirement",
        "experience_overview": "Experience overview",
        "roles_and_responsibilities": "Roles and Responsibilities",
        "number_of_months": "Number of Months",
        "total_relevant_months": "Total Number of Months of Relevant Professional Experience",
        "role_prefix": "Role: ",
        "responsibilities_prefix": "\nResponsibilities: ",
        "project_experience": "Project Experience",
        "project_allocation": "Project Allocation in % (only specified in case of parallel projects)",
        "education_and_training": "Education and Training",
        "other_languages": "Other language(s)",
        "understanding": "UNDERSTANDING",
        "speaking": "SPEAKING",
        "writing": "WRITING",
        "listening": "Listening",
        "reading": "Reading",
        "spoken_interaction": "Spoken interaction",
        "spoken_production": "Spoken production",
        "cefr_levels": "Levels: A1/2: Basic user - B1/2: Independent user - C1/2 Proficient user",
        "cefr_reference": "Common European Framework of Reference for Languages",
        "personal_skills": "Personal Skills",
        "mother_tongues": "Mother tongue(s)",
        "communication_skills": "Communication skills",
        "organisational_managerial_skills": "Organisational / managerial skills",
        "computer_skills": "Computer skills",
        "certifications": "Certifications",
        "perpetual": "(perpetual)",
        "expired_in": "(expired in {year})",
        "expires_in": "(expires in {year})",
        "candidate_name": "Candidate Name",
        "mobile": "Mobile",
        "email": "Email",
        "project": "Project",
    },
    "fr": {
        "docx_months": ["Janv.", "Févr.", "Mars", "Avr.", "Mai", "Juin", "Juil.", "Août", "Sept.", "Oct.", "Nov.", "Déc."],
        "pptx_months": ["Janv.", "Févr.", "Mars", "Avr.", "Mai", "Juin", "Juil.", "Août", "Sept.", "Oct.", "Nov.", "Déc."],
        "present": "Présent",
        "europass": "Europass",
        "curriculum_vitae": "Curriculum Vitae",
        "page": "Page ",
        "personal_information": "INFORMATIONS PERSONNELLES",
        "photo_placeholder": "[Photo 3,8×3,8 cm]",
        "sex_nationality": "Sexe {sex} | Nationalité {nationality}",
        "proposed_role": "POSTE PROPOSÉ",
        "professional_experience": "EXPÉRIENCE\nPROFESSIONNELLE",
        "years_of_work_experience": "Nombre d’années d’expérience professionnelle :",
        "experience_buckets": ["Moins de 4", "4 – 9", "10 – 14", "15+"],
        "work_experience": "EXPÉRIENCE\nPROFESSIONNELLE",
        "date": "DATE",
        "title": "INTITULÉ",
        "organisation": "ORGANISATION",
        "profile": "PROFIL",
        "requirements_matrix": "Matrice des exigences",
        "requirement": "Exigence",
        "experience_overview": "Aperçu de l’expérience",
        "roles_and_responsibilities": "Rôles et responsabilités",
        "number_of_months": "Nombre de mois",
        "total_relevant_months": "Nombre total de mois d’expérience professionnelle pertinente",
        "role_prefix": "Rôle : ",
        "responsibilities_prefix": "\nResponsabilités : ",
        "project_experience": "Expérience projet",
        "project_allocation": "Affectation au projet en % (uniquement en cas de projets parallèles)",
        "education_and_training": "Enseignement et formation",
        "other_languages": "Autre(s) langue(s)",
        "understanding": "COMPRÉHENSION",
        "speaking": "EXPRESSION ORALE",
        "writing": "ÉCRITURE",
        "listening": "Écoute",
        "reading": "Lecture",
        "spoken_interaction": "Interaction orale",
        "spoken_production": "Production orale",
        "cefr_levels": "Niveaux : A1/2 : utilisateur élémentaire - B1/2 : utilisateur indépendant - C1/2 : utilisateur expérimenté",
        "cefr_reference": "Cadre européen commun de référence pour les langues",
        "personal_skills": "Compétences personnelles",
        "mother_tongues": "Langue(s) maternelle(s)",
        "communication_skills": "Compétences en communication",
        "organisational_managerial_skills": "Compétences organisationnelles\n/ managériales",
        "computer_skills": "Compétences informatiques",
        "certifications": "Certifications",
        "perpetual": "(perpétuelle)",
        "expired_in": "(expirée en {year})",
        "expires_in": "(expire en {year})",
        "candidate_name": "Nom du candidat",
        "mobile": "Mobile",
        "email": "E-mail",
        "project": "Projet",
    },
}


def normalize_language(language: Optional[str]) -> str:
    if isinstance(language, str) and language in CV_STRINGS:
        return language
    return DEFAULT_CV_LANGUAGE


def get_cv_strings(language: Optional[str] = None) -> Dict[str, Any]:
    return CV_STRINGS[normalize_language(language)]


def normalize_sections(
    sections: Optional[Dict[str, Any]], defaults: Dict[str, bool]
) -> Dict[str, bool]:
    normalized = defaults.copy()
    if not isinstance(sections, dict):
        return normalized
    for key in defaults:
        if key in sections:
            normalized[key] = bool(sections[key])
    return normalized
