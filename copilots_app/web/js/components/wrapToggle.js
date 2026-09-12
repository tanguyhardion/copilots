/**
 * Wrap Toggle Component
 *
 * Adds a small, discreet wrap/unwrap toggle button inside the top-right
 * corner of any `.code-area-wrapper` that contains a `.code-textarea`.
 *
 * Usage:
 *   import { initWrapToggle } from "./components/wrapToggle.js";
 *   initWrapToggle("my-textarea-id");        // default: wrapped = true
 *   initWrapToggle("my-textarea-id", false); // default: wrapped = false
 */

/**
 * Initialise a wrap toggle for the given textarea.
 *
 * @param {string} textareaId  - The id of the textarea element.
 * @param {boolean} [wrappedByDefault=true] - Whether text wrapping starts on.
 */
export function initWrapToggle(textareaId, wrappedByDefault = true) {
  const textarea = document.getElementById(textareaId);
  if (!textarea) return;

  const wrapper = textarea.closest(".code-area-wrapper");
  if (!wrapper) return;

  // Apply initial wrap state
  applyWrap(textarea, wrappedByDefault);

  // Build the toggle button
  const btn = document.createElement("button");
  btn.type = "button";
  btn.className = "wrap-toggle-btn";
  btn.title = wrappedByDefault ? "Unwrap text" : "Wrap text";
  btn.setAttribute("aria-label", btn.title);
  btn.dataset.wrapped = wrappedByDefault ? "true" : "false";

  btn.innerHTML = buildIcon(wrappedByDefault);

  btn.addEventListener("click", (e) => {
    e.preventDefault();
    e.stopPropagation();

    const isWrapped = btn.dataset.wrapped === "true";
    const nextWrapped = !isWrapped;

    btn.dataset.wrapped = nextWrapped ? "true" : "false";
    btn.title = nextWrapped ? "Unwrap text" : "Wrap text";
    btn.setAttribute("aria-label", btn.title);
    btn.innerHTML = buildIcon(nextWrapped);

    applyWrap(textarea, nextWrapped);

    // Re-render Lucide icon after innerHTML replacement
    if (window.lucide) lucide.createIcons({ nodes: [btn] });
  });

  // Ensure the wrapper can anchor the absolutely-positioned button.
  // `.code-area-wrapper` already sets position:relative via CSS, but set it
  // explicitly here as a safety net for any edge cases (e.g. hidden panels).
  wrapper.style.position = "relative";

  wrapper.appendChild(btn);

  // Re-render the Lucide icon if the library is present
  if (window.lucide) lucide.createIcons({ nodes: [btn] });
}

/**
 * Bulk-initialise wrap toggles for multiple textarea ids.
 * @param {string[]} ids
 * @param {boolean} [wrappedByDefault=true]
 */
export function initWrapToggles(ids, wrappedByDefault = true) {
  ids.forEach((id) => initWrapToggle(id, wrappedByDefault));
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function applyWrap(textarea, wrap) {
  if (wrap) {
    textarea.style.whiteSpace = "pre-wrap";
    textarea.style.overflowX  = "hidden";
    textarea.wrap             = "soft";
  } else {
    textarea.style.whiteSpace = "pre";
    textarea.style.overflowX  = "auto";
    textarea.wrap             = "off";
  }
}

/**
 * Returns an SVG icon string representing the current wrap state.
 * Uses Lucide-style strokes so it blends with the rest of the UI.
 *
 * wrap-on  → "wrap lines" icon (two horizontal lines with a curved arrow)
 * wrap-off → "no-wrap" icon (a single long straight line)
 */
function buildIcon(wrapped) {
  // We rely on Lucide icons already loaded on the page.
  // "wrap-text" is a Lucide icon (available since v0.263).
  // Fall back to a minimal inline SVG if not present.
  if (wrapped) {
    return `<i data-lucide="wrap-text" class="wrap-toggle-icon"></i>`;
  } else {
    // There is no canonical "no-wrap" Lucide icon, so we draw one inline.
    return `<svg class="wrap-toggle-icon" viewBox="0 0 24 24" fill="none"
      stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"
      xmlns="http://www.w3.org/2000/svg">
      <line x1="3" y1="8" x2="21" y2="8"/>
      <line x1="3" y1="12" x2="21" y2="12"/>
      <line x1="3" y1="16" x2="21" y2="16"/>
    </svg>`;
  }
}
