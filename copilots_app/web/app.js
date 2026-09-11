/**
 * Copilots Suite — Application Entry Point
 *
 * This file is intentionally minimal. Its only job is to import each module
 * and wire them together in initApp(). All business logic lives in the
 * dedicated modules under js/.
 */

import { setupSidebar, navigateTo } from "./js/core/router.js";
import { setupModals }              from "./js/modals/modals.js";
import { setupPowerPointView }      from "./js/views/powerpoint.js";
import { setupWordView }            from "./js/views/word.js";
import { setupExcelView }           from "./js/views/excel.js";
import { setupCVView }              from "./js/views/cv.js";
import { setupPythonView }          from "./js/views/python.js";
import { setupOrganizerView }       from "./js/views/organizer.js";

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
    setupWordView(),
    setupExcelView(),
    setupCVView(),
    setupPythonView(),
    setupOrganizerView(),
  ]);

  navigateTo("powerpoint");

  // Render all static Lucide icons
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
