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
    bgGlowColor: "var(--brand-ppt-bg-glow)",
    bgGlowSubtleColor: "var(--brand-ppt-bg-glow-subtle)",
    actions: [
      { id: "action-prompt", text: "System Prompt", icon: "settings" },
      { id: "action-help", text: "", icon: "help-circle", title: "PowerPoint Copilot Guide", isIconOnly: true },
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
    bgGlowColor: "var(--brand-cv-bg-glow)",
    bgGlowSubtleColor: "var(--brand-cv-bg-glow-subtle)",
    actions: [
      { id: "action-prompt", text: "System Prompt", icon: "settings" },
      { id: "action-help", text: "", icon: "help-circle", title: "CV Copilot Guide", isIconOnly: true },
    ]
  }
};
