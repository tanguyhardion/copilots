/**
 * Email Copilot view controller.
 */

import { setStatus, setButtonsDisabled } from "../core/utils.js";

/**
 * Bind all action buttons, input controls, and quick actions in the Email view panel.
 */
export async function setupEmailView() {
  const dslEditor = document.getElementById("email-dsl-editor");
  const resultsArea = document.getElementById("email-results-text");
  const limitInput = document.getElementById("email-cap-limit");
  const charsInput = document.getElementById("email-cap-chars");
  const folderInput = document.getElementById("email-folder-select");

  const btnExecute = document.getElementById("email-btn-execute");
  const btnUnread = document.getElementById("email-btn-unread");
  const btnCopyResults = document.getElementById("email-btn-copy-results");
  const btnClear = document.getElementById("email-btn-clear");

  if (!dslEditor || !btnExecute) return;

  // Execute DSL search / command
  btnExecute.addEventListener("click", async () => {
    const dslText = dslEditor.value.trim();
    if (!dslText) {
      setStatus("email", "Please enter an Email DSL query (e.g. SEARCH query=\"budget\" or SUMMARIZE_UNREAD).", "warning");
      return;
    }

    const defaultLimit = parseInt(limitInput?.value || "20", 10);
    const defaultChars = parseInt(charsInput?.value || "300", 10);

    setStatus("email", "Connecting to Outlook & querying emails…", "info", true);
    setButtonsDisabled("view-email", true);

    try {
      const res = await window.pywebview.api.email_execute_dsl(dslText, defaultLimit, defaultChars);
      if (res.success) {
        resultsArea.value = res.context_markdown;
        setStatus("email", `✓ ${res.message} (Markdown context ready for LLM)`, "success");
        // Automatically copy to clipboard for quick paste to LLM
        try {
          await navigator.clipboard.writeText(res.context_markdown);
        } catch (e) {}
      } else {
        setStatus("email", res.error || "Email query failed.", "error");
      }
    } catch (err) {
      setStatus("email", `Outlook error: ${err}`, "error");
    } finally {
      setButtonsDisabled("view-email", false);
    }
  });

  // Quick Summarize Unread action
  btnUnread.addEventListener("click", async () => {
    const folder = folderInput?.value || "inbox";
    const defaultLimit = parseInt(limitInput?.value || "15", 10);
    const defaultChars = parseInt(charsInput?.value || "250", 10);

    // Set sample DSL in the editor for transparency
    dslEditor.value = `SUMMARIZE_UNREAD folder=${folder} limit=${defaultLimit} max_chars=${defaultChars}`;

    setStatus("email", "Pulling unread emails from Outlook…", "info", true);
    setButtonsDisabled("view-email", true);

    try {
      const res = await window.pywebview.api.email_fetch_unread(folder, defaultLimit, defaultChars);
      if (res.success) {
        resultsArea.value = res.context_markdown;
        setStatus("email", `✓ ${res.message} (Copied to clipboard for LLM triage)`, "success");
        try {
          await navigator.clipboard.writeText(res.context_markdown);
        } catch (e) {}
      } else {
        setStatus("email", res.error || "Failed to fetch unread emails.", "error");
      }
    } catch (err) {
      setStatus("email", `Outlook error: ${err}`, "error");
    } finally {
      setButtonsDisabled("view-email", false);
    }
  });

  // Copy results markdown to clipboard
  btnCopyResults.addEventListener("click", async () => {
    const text = resultsArea.value;
    if (!text) {
      setStatus("email", "No results to copy. Run a query or pull unread emails first.", "warning");
      return;
    }
    try {
      await navigator.clipboard.writeText(text);
      setStatus("email", "✓ Email context copied to clipboard!", "success");
    } catch (e) {
      setStatus("email", `Copy error: ${e}`, "error");
    }
  });

  // Clear editor
  if (btnClear) {
    btnClear.addEventListener("click", () => {
      dslEditor.value = "";
      resultsArea.value = "";
      setStatus("email", "Ready — enter DSL search or click Summarize Unread.", "info");
    });
  }

  // Populate folder select or query status if available
  try {
    const accRes = await window.pywebview?.api?.email_get_accounts?.();
    if (accRes?.success && folderInput) {
      // populate or retain standard folders
    }
  } catch (e) {}
}
