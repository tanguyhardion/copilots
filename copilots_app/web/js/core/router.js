/**
 * Client-side router — sidebar setup, view panel switching, and header rendering.
 */

import { ROUTES } from "../data/routes.js";
import { openPromptModal, openCheatsheetModal, openHelpModal } from "../modals/modals.js";
import { excelConnectActive, excelOpenFile } from "../views/excel.js";
import { pythonOpenExplorer, pythonCopyFolderContext, pythonClearWorkspace } from "../views/python.js";

/** Currently active route key. */
export let currentRoute = "powerpoint";

/**
 * Wire up sidebar nav-item click listeners.
 */
export function setupSidebar() {
  document.querySelectorAll(".nav-item").forEach(item => {
    item.addEventListener("click", () => {
      const route = item.getAttribute("data-route");
      navigateTo(route);
    });
  });
}

/**
 * Navigate to a copilot view by route id.
 * Updates the active nav item, visible view panel, header content, and CSS theme variables.
 * @param {string} routeId - Key in ROUTES config.
 */
export function navigateTo(routeId) {
  if (!ROUTES[routeId]) return;
  currentRoute = routeId;

  // Update nav selection
  document.querySelectorAll(".nav-item").forEach(item => {
    item.classList.toggle("active", item.getAttribute("data-route") === routeId);
  });

  // Update view panels
  document.querySelectorAll(".view-panel").forEach(panel => {
    panel.classList.toggle("active", panel.id === `view-${routeId}`);
  });

  const routeMeta = ROUTES[routeId];
  document.getElementById("header-title").innerText = routeMeta.title;
  document.getElementById("header-subtitle").innerText = routeMeta.subtitle;
  document.getElementById("header-icon").src = routeMeta.icon;

  // Apply per-copilot accent CSS variables
  document.documentElement.style.setProperty("--copilot-accent", routeMeta.badgeColor);
  document.documentElement.style.setProperty("--copilot-accent-subtle", routeMeta.badgeSubtleColor || "var(--primary-subtle)");
  document.documentElement.style.setProperty("--copilot-accent-glow", routeMeta.badgeGlowColor || "var(--primary-glow)");

  const badge = document.getElementById("header-badge");
  badge.innerText = routeMeta.badge;
  badge.style.backgroundColor = routeMeta.badgeColor;

  // Render header action buttons dynamically
  const actionsContainer = document.getElementById("header-actions");
  actionsContainer.innerHTML = "";

  routeMeta.actions.forEach(act => {
    const btn = document.createElement("button");
    btn.className = act.isIconOnly ? "btn btn-secondary btn-icon-only" : "btn btn-secondary";
    btn.id = act.id;
    if (act.title) {
      btn.title = act.title;
      btn.setAttribute("aria-label", act.title);
    }
    btn.innerHTML = act.isIconOnly
      ? `<i data-lucide="${act.icon}" class="btn-icon"></i>`
      : `<i data-lucide="${act.icon}" class="btn-icon"></i> ${act.text}`;
    btn.addEventListener("click", () => handleHeaderAction(act.id, routeId));
    actionsContainer.appendChild(btn);
  });

  // Re-render Lucide icons for dynamically inserted buttons
  if (window.lucide) lucide.createIcons();
}

/**
 * Dispatch a header action button click to the appropriate handler.
 * @param {string} actionId - The action button id.
 * @param {string} routeId - The active route id.
 */
export function handleHeaderAction(actionId, routeId) {
  const handlers = {
    "action-prompt":              () => openPromptModal(routeId),
    "action-cheatsheet":          () => openCheatsheetModal(routeId),
    "action-help":                () => openHelpModal(routeId),
    "action-excel-active":        () => excelConnectActive(),
    "action-excel-open":          () => excelOpenFile(),
    "action-python-open-folder":  () => pythonOpenExplorer(),
    "action-python-copy-context": () => pythonCopyFolderContext(),
    "action-python-clear":        () => pythonClearWorkspace(),
  };
  handlers[actionId]?.();
}
