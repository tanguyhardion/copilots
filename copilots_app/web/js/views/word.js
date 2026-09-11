/**
 * Word Copilot view controller.
 */

import { setStatus, setButtonsDisabled } from "../core/utils.js";

/**
 * Bind all action buttons in the Word view panel.
 */
export async function setupWordView() {
  const editor = document.getElementById("word-editor");

  // Build and open a standalone .docx document
  document.getElementById("word-btn-build").addEventListener("click", async () => {
    setStatus("word", "Compiling Word document (.docx)…", "info", true);
    setButtonsDisabled("view-word", true);
    try {
      const res = await window.pywebview.api.word_build_and_open(editor.value);
      setStatus("word", res.success ? res.message : res.error, res.success ? "success" : "error");
    } catch (err) {
      setStatus("word", `Error: ${err}`, "error");
    } finally {
      setButtonsDisabled("view-word", false);
    }
  });

  // Insert DSL content at the active Word cursor position
  document.getElementById("word-btn-insert").addEventListener("click", async () => {
    setStatus("word", "Injecting elements at active Word cursor…", "info", true);
    setButtonsDisabled("view-word", true);
    try {
      const res = await window.pywebview.api.word_insert_at_cursor(editor.value);
      setStatus("word", res.success ? res.message : res.error, res.success ? "success" : "error");
    } catch (err) {
      setStatus("word", `Error: ${err}`, "error");
    } finally {
      setButtonsDisabled("view-word", false);
    }
  });

  // Apply deterministic in-place edits to the active Word document
  document.getElementById("word-btn-apply").addEventListener("click", async () => {
    setStatus("word", "Executing edit plan in active Word document…", "info", true);
    setButtonsDisabled("view-word", true);
    try {
      const res = await window.pywebview.api.word_apply_edits(editor.value);
      setStatus("word", res.success ? res.message : res.error, res.success ? "success" : "error");
    } catch (err) {
      setStatus("word", `Error: ${err}`, "error");
    } finally {
      setButtonsDisabled("view-word", false);
    }
  });

  // Extract DSL from the currently active Word document
  document.getElementById("word-btn-extract").addEventListener("click", async () => {
    setStatus("word", "Extracting structure from active Word document…", "info", true);
    setButtonsDisabled("view-word", true);
    try {
      const res = await window.pywebview.api.word_extract_dsl();
      if (res.success) {
        editor.value = res.dsl;
        setStatus("word", res.message, "success");
      } else {
        setStatus("word", res.error, "error");
      }
    } catch (err) {
      setStatus("word", `Error: ${err}`, "error");
    } finally {
      setButtonsDisabled("view-word", false);
    }
  });
}
