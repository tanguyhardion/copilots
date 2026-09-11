/**
 * PowerPoint Copilot view controller.
 */

import { setStatus, setButtonsDisabled } from "../core/utils.js";

/**
 * Bind all action buttons in the PowerPoint view panel.
 */
export async function setupPowerPointView() {
  const editor = document.getElementById("ppt-editor");

  // Copy DSL shapes to Windows clipboard
  document.getElementById("ppt-btn-copy").addEventListener("click", async () => {
    const dsl = editor.value;
    if (!dsl.trim()) {
      setStatus("ppt", "DSL is empty — nothing to copy", "warning");
      return;
    }
    setStatus("ppt", "Building shapes for clipboard…", "info", true);
    setButtonsDisabled("view-powerpoint", true);
    try {
      const res = await window.pywebview.api.ppt_copy_clipboard(dsl);
      setStatus("ppt", res.success ? res.message : res.error, res.success ? "success" : "error");
    } catch (err) {
      setStatus("ppt", `Error: ${err}`, "error");
    } finally {
      setButtonsDisabled("view-powerpoint", false);
    }
  });

  // Insert shapes onto the currently active slide
  document.getElementById("ppt-btn-insert").addEventListener("click", async () => {
    const dsl = editor.value;
    if (!dsl.trim()) {
      setStatus("ppt", "DSL is empty — nothing to insert", "warning");
      return;
    }
    setStatus("ppt", "Inserting shapes onto active slide…", "info", true);
    setButtonsDisabled("view-powerpoint", true);
    try {
      const res = await window.pywebview.api.ppt_insert_current_slide(dsl);
      setStatus("ppt", res.success ? res.message : res.error, res.success ? "success" : "error");
    } catch (err) {
      setStatus("ppt", `Error: ${err}`, "error");
    } finally {
      setButtonsDisabled("view-powerpoint", false);
    }
  });

  // Create full slide(s) at the end of the active presentation
  document.getElementById("ppt-btn-create").addEventListener("click", async () => {
    const dsl = editor.value;
    if (!dsl.trim()) {
      setStatus("ppt", "DSL is empty — nothing to build", "warning");
      return;
    }
    setStatus("ppt", "Creating new slide(s)…", "info", true);
    setButtonsDisabled("view-powerpoint", true);
    try {
      const res = await window.pywebview.api.ppt_create_full_slides(dsl);
      setStatus("ppt", res.success ? res.message : res.error, res.success ? "success" : "error");
    } catch (err) {
      setStatus("ppt", `Error: ${err}`, "error");
    } finally {
      setButtonsDisabled("view-powerpoint", false);
    }
  });
}
