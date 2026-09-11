/**
 * Core UI utility functions shared across all view modules.
 */

/**
 * Update the status bar for a given view.
 * @param {string} viewId - The view prefix (e.g. "ppt", "word", "excel").
 * @param {string} message - The status message to display.
 * @param {"info"|"success"|"warning"|"error"} level - Severity level, controls styling.
 * @param {boolean} loading - When true, shows the loading animation.
 */
export function setStatus(viewId, message, level = "info", loading = false) {
  const bar = document.getElementById(`${viewId}-status`);
  if (!bar) return;

  bar.className = `status-bar ${level} ${loading ? "loading" : ""}`;
  const textElem = bar.querySelector(".status-text");
  if (textElem) textElem.innerText = message;
}

/**
 * Enable or disable all action-bar and header-action buttons within a container.
 * @param {string} containerId - The id of the parent container element.
 * @param {boolean} disabled - True to disable, false to re-enable.
 */
export function setButtonsDisabled(containerId, disabled) {
  const container = document.getElementById(containerId);
  if (!container) return;
  container.querySelectorAll(".action-bar button, .header-actions button").forEach(btn => {
    btn.disabled = disabled;
    btn.style.opacity = disabled ? "0.6" : "1.0";
    btn.style.pointerEvents = disabled ? "none" : "auto";
  });
}

/**
 * Escape HTML special characters to prevent XSS when setting innerHTML.
 * @param {string} text - Raw text to escape.
 * @returns {string} HTML-safe string.
 */
export function escapeHtml(text) {
  const map = {
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#039;"
  };
  return text.replace(/[&<>"']/g, m => map[m]);
}
