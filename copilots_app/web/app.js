/**
 * Copilots Suite — Application Entry Point
 *
 * Imports view controllers and wires them together in initApp().
 */

import { setupSidebar, navigateTo } from "./js/core/router.js";
import { setupModals }              from "./js/modals/modals.js";
import { setupPowerPointView }      from "./js/views/powerpoint.js";
import { setupCVView }              from "./js/views/cv.js";
import { initWrapToggles }          from "./js/components/wrapToggle.js";

// Wait for pywebview API to be ready
window.addEventListener("pywebviewready", () => {
  console.log("pywebview API ready!");
  initApp();
});

// Fallback in case opened in a regular browser during development
setTimeout(() => {
  if (!window.__initialized) initApp();
}, 600);

async function initApp() {
  if (window.__initialized) return;
  window.__initialized = true;

  setupSidebar();
  setupModals();
  loadVersion();

  await Promise.all([
    setupPowerPointView(),
    setupCVView(),
  ]);

  navigateTo("powerpoint");

  // Initialise wrap-toggle buttons on main textarea editors
  initWrapToggles([
    "ppt-editor",
    "cv-editor",
    "modal-prompt-editor",
  ]);

  // Render all static Lucide icons (must run after wrap-toggle buttons are injected)
  if (window.lucide) lucide.createIcons();
}

async function loadVersion() {
  const el = document.getElementById("brand-version");
  if (!el) return;
  try {
    const version = await window.pywebview?.api?.get_version?.();
    el.textContent = version ? `Suite v${version}` : "Suite";
  } catch {
    // Running in browser without pywebview — leave as-is
  }
}
