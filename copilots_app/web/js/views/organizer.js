/**
 * Folder/File Organizer Copilot view controller.
 */

import { setStatus } from "../core/utils.js";

/**
 * Bind all action buttons in the Organizer view panel.
 * Also queries initial organizer status from the Python API.
 */
export async function setupOrganizerView() {
  const selectBtn    = document.getElementById("organizer-btn-select-folder");
  const copyBtn      = document.getElementById("organizer-btn-copy-context");
  const executeBtn   = document.getElementById("organizer-btn-execute-plan");
  const openFolderBtn = document.getElementById("organizer-btn-open-folder");
  const planTextarea = document.getElementById("organizer-plan-text");

  if (!selectBtn || !copyBtn || !executeBtn) return;

  selectBtn.addEventListener("click", organizerSelectFolder);
  copyBtn.addEventListener("click", organizerCopyContext);
  executeBtn.addEventListener("click", organizerExecutePlan);
  if (openFolderBtn) openFolderBtn.addEventListener("click", organizerOpenOutputFolder);

  // Auto-validate plan text on input (debounced)
  let _validateTimer = null;
  planTextarea.addEventListener("input", () => {
    clearTimeout(_validateTimer);
    _validateTimer = setTimeout(organizerValidatePlan, 500);
  });

  // Query existing organizer state
  try {
    const res = await window.pywebview?.api?.organizer_get_status?.();
    if (res?.success) {
      updateOrganizerLabels(res.source_root, res.last_output_root);
      const openBtn = document.getElementById("organizer-btn-open-folder");
      if (openBtn && res.last_output_root) openBtn.disabled = false;
    }
  } catch (e) {}
}

/**
 * Update the source and output path labels.
 * @param {string|null} sourceRoot - Source directory path.
 * @param {string|null} outputRoot - Output directory path.
 */
export function updateOrganizerLabels(sourceRoot, outputRoot) {
  const sourceLabel = document.getElementById("organizer-source-label");
  const outputLabel = document.getElementById("organizer-output-label");

  if (sourceLabel) {
    sourceLabel.innerText = sourceRoot ? `Source: ${sourceRoot}` : "No folder selected.";
    sourceLabel.title     = sourceRoot || "";
  }
  if (outputLabel) {
    outputLabel.innerText = outputRoot ? `Output: ${outputRoot}` : "Output: not generated yet";
    outputLabel.title     = outputRoot || "";
  }
}

/**
 * Open a folder picker dialog and set the source folder.
 */
export async function organizerSelectFolder() {
  try {
    if (window.pywebview?.api) {
      const res = await window.pywebview.api.organizer_select_folder();
      if (res.success) {
        updateOrganizerLabels(res.source_root, "");
        setStatus("organizer", "Folder selected. Click 'Copy for LLM' to generate recursive context.", "success");
      } else if (!res.cancelled) {
        setStatus("organizer", res.error || "Could not select folder.", "error");
      }
    }
  } catch (err) {
    setStatus("organizer", `Folder selection error: ${err}`, "error");
  }
}

/**
 * Build the recursive folder context and copy it to clipboard.
 */
export async function organizerCopyContext() {
  try {
    if (window.pywebview?.api) {
      const res = await window.pywebview.api.organizer_build_context();
      if (res.success) {
        document.getElementById("organizer-context-text").value = res.context_markdown;
        await navigator.clipboard.writeText(res.context_markdown);
        setStatus("organizer", `✓ Copied recursive context (${res.file_count} file(s), ${res.excerpted_count} excerpt(s)) to clipboard.`, "success");
      } else {
        setStatus("organizer", res.error || "Could not build folder context.", "error");
      }
    }
  } catch (err) {
    setStatus("organizer", `Context generation error: ${err}`, "error");
  }
}

/**
 * Parse and validate the plan DSL; display instruction count or errors.
 */
export async function organizerValidatePlan() {
  const planText = document.getElementById("organizer-plan-text").value;
  if (!planText.trim()) {
    setStatus("organizer", "Plan is empty. Add MOVE instructions first.", "warning");
    return;
  }

  try {
    if (window.pywebview?.api) {
      const res = await window.pywebview.api.organizer_parse_plan(planText);
      if (res.instruction_count > 0) {
        const suffix = (res.errors && res.errors.length) ? ` (${res.errors.length} invalid line(s) ignored)` : "";
        setStatus("organizer", `Plan parsed: ${res.instruction_count} valid MOVE instruction(s).${suffix}`, "success");
      } else {
        setStatus("organizer", (res.errors && res.errors[0]) || "No valid MOVE instructions parsed.", "warning");
      }
    }
  } catch (err) {
    setStatus("organizer", `Plan validation error: ${err}`, "error");
  }
}

/**
 * Execute the copy plan and update output labels.
 */
export async function organizerExecutePlan() {
  const planText = document.getElementById("organizer-plan-text").value;
  if (!planText.trim()) {
    setStatus("organizer", "Plan is empty. Add MOVE instructions first.", "warning");
    return;
  }

  try {
    if (window.pywebview?.api) {
      const res = await window.pywebview.api.organizer_execute_plan(planText);
      if (res.success) {
        const status = await window.pywebview.api.organizer_get_status();
        if (status?.success) {
          updateOrganizerLabels(status.source_root, res.output_root || status.last_output_root);
        } else {
          updateOrganizerLabels("", res.output_root);
        }
        const openFolderBtn = document.getElementById("organizer-btn-open-folder");
        if (openFolderBtn) openFolderBtn.disabled = false;
        setStatus("organizer", res.message, "success");
      } else {
        setStatus("organizer", res.error || res.message || "Plan execution failed.", "error");
      }
    }
  } catch (err) {
    setStatus("organizer", `Plan execution error: ${err}`, "error");
  }
}

/**
 * Open the last output folder in Windows Explorer.
 */
export async function organizerOpenOutputFolder() {
  try {
    if (window.pywebview?.api) {
      const res = await window.pywebview.api.organizer_open_output_folder();
      if (!res.success) {
        setStatus("organizer", res.error || "Could not open output folder.", "error");
      }
    }
  } catch (err) {
    setStatus("organizer", `Open folder error: ${err}`, "error");
  }
}
