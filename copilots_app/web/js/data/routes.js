/**
 * Route configuration — single source of truth for all copilot metadata.
 * Each key maps to a view panel id ("view-{key}") and nav item data-route attribute.
 */
export const ROUTES = {
  powerpoint: {
    title: "PowerPoint Copilot",
    subtitle: "Generate slides, shapes, rich text, tables, and icons from clean DSL",
    icon: "../../assets/icons/powerpoint.png",
    badge: "PowerPoint COM",
    badgeColor: "var(--brand-ppt)",
    badgeSubtleColor: "var(--brand-ppt-subtle)",
    badgeGlowColor: "var(--brand-ppt-glow)",
    actions: [
      { id: "action-prompt", text: "System Prompt", icon: "settings" },
      { id: "action-help", text: "", icon: "help-circle", title: "PowerPoint Copilot Guide", isIconOnly: true },
    ]
  },
  word: {
    title: "Word Copilot",
    subtitle: "Document generator, live cursor injector, and active document editor with DSL extraction",
    icon: "../../assets/icons/word.png",
    badge: "Word COM",
    badgeColor: "var(--brand-word)",
    badgeSubtleColor: "var(--brand-word-subtle)",
    badgeGlowColor: "var(--brand-word-glow)",
    actions: [
      { id: "action-prompt", text: "System Prompt", icon: "settings" },
      { id: "action-help", text: "", icon: "help-circle", title: "Word Copilot Guide", isIconOnly: true },
    ]
  },
  excel: {
    title: "Excel Copilot",
    subtitle: "Active workbook COM automation, real-time spreadsheet analysis, and deterministic JSON action execution",
    icon: "../../assets/icons/excel.png",
    badge: "Excel COM",
    badgeColor: "var(--brand-excel)",
    badgeSubtleColor: "var(--brand-excel-subtle)",
    badgeGlowColor: "var(--brand-excel-glow)",
    actions: [
      { id: "action-prompt", text: "System Prompt", icon: "settings" },
      { id: "action-help", text: "", icon: "help-circle", title: "Excel Copilot Guide", isIconOnly: true },
    ]
  },
  cv: {
    title: "CV Copilot",
    subtitle: "Deterministic Data Quality engine, Europass profile validation, and formatted Word .docx generator",
    icon: "../../assets/icons/cv.png",
    badge: "DQ Engine + docx",
    badgeColor: "var(--brand-cv)",
    badgeSubtleColor: "var(--brand-cv-subtle)",
    badgeGlowColor: "var(--brand-cv-glow)",
    actions: [
      { id: "action-prompt", text: "System Prompt", icon: "settings" },
      { id: "action-help", text: "", icon: "help-circle", title: "CV Copilot Guide", isIconOnly: true },
    ]
  },
  python: {
    title: "Python Copilot",
    subtitle: "Interactive execution sandbox, script runner, live stdout/stderr console, and directory context inspector",
    icon: "../../assets/icons/python.png",
    badge: "Python Runner Sandbox",
    badgeColor: "var(--brand-python)",
    badgeSubtleColor: "var(--brand-python-subtle)",
    badgeGlowColor: "var(--brand-python-glow)",
    actions: [
      { id: "action-prompt", text: "System Prompt", icon: "settings" },
      { id: "action-help", text: "", icon: "help-circle", title: "Python Copilot Guide", isIconOnly: true },
    ]
  },
  organizer: {
    title: "Folder/File Organizer Copilot",
    subtitle: "Analyze recursive folder context, prepare LLM-ready summaries, and execute safe copy-only reorganization plans",
    icon: "../../assets/icons/folder-organizer.png",
    badge: "Folder DSL + Safe Copy",
    badgeColor: "var(--brand-organizer)",
    badgeSubtleColor: "var(--brand-organizer-subtle)",
    badgeGlowColor: "var(--brand-organizer-glow)",
    actions: [
      { id: "action-prompt", text: "System Prompt", icon: "settings" },
      { id: "action-help", text: "", icon: "help-circle", title: "Folder Organizer Copilot Guide", isIconOnly: true },
    ]
  }
};
