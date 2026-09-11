/**
 * Excel Copilot view controller.
 */

import { setStatus, setButtonsDisabled } from "../core/utils.js";

/**
 * Bind all action buttons and tab switchers in the Excel view panel.
 */
export async function setupExcelView() {
  const protocolEditor = document.getElementById("excel-protocol-editor");

  // Tab switcher
  const btnExplorer = document.getElementById("tab-btn-explorer");
  const btnContext  = document.getElementById("tab-btn-context");
  const pageExplorer = document.getElementById("tab-page-explorer");
  const pageContext  = document.getElementById("tab-page-context");

  btnExplorer.addEventListener("click", () => {
    btnExplorer.classList.add("active");
    btnContext.classList.remove("active");
    pageExplorer.classList.add("active");
    pageContext.classList.remove("active");
  });

  btnContext.addEventListener("click", () => {
    btnContext.classList.add("active");
    btnExplorer.classList.remove("active");
    pageContext.classList.add("active");
    pageExplorer.classList.remove("active");
  });

  // Copy the LLM prompt context to clipboard
  document.getElementById("excel-btn-copy-context").addEventListener("click", () => {
    const text = document.getElementById("excel-context-text").value;
    navigator.clipboard.writeText(text).then(() => {
      setStatus("excel", "Copied LLM prompt context to clipboard!", "success");
    });
  });

  // Connect to the active Excel workbook / open a file
  document.getElementById("excel-btn-open").addEventListener("click", excelOpenFile);
  document.getElementById("excel-btn-active").addEventListener("click", excelConnectActive);

  // Validate the JSON Action Protocol in the editor
  document.getElementById("excel-btn-validate").addEventListener("click", () => {
    const logElem = document.getElementById("excel-log-console");
    try {
      const parsed = JSON.parse(protocolEditor.value);
      if (!parsed.intent || !Array.isArray(parsed.actions)) {
        throw new Error("Missing 'intent' or 'actions' array in JSON.");
      }
      logElem.innerText = `[VALIDATION OK] Valid JSON Action Protocol (${parsed.actions.length} actions defined). Intent: ${parsed.intent}`;
      setStatus("excel", "Protocol JSON syntax is valid.", "success");
    } catch (e) {
      logElem.innerText = `[VALIDATION ERROR] ${e.message}`;
      setStatus("excel", `Validation error: ${e.message}`, "error");
    }
  });

  // Execute the action protocol against the connected workbook
  document.getElementById("excel-btn-execute").addEventListener("click", async () => {
    const jsonStr = protocolEditor.value;
    const backup  = document.getElementById("excel-backup-checkbox").checked;
    const logElem = document.getElementById("excel-log-console");

    setStatus("excel", "Executing action protocol…", "info", true);
    setButtonsDisabled("view-excel", true);
    try {
      const res = await window.pywebview.api.excel_execute_protocol(jsonStr, backup);
      if (res.logs && res.logs.length > 0) {
        logElem.innerText = res.logs.join("\n");
      }
      setStatus("excel", res.success ? res.message : (res.error || res.message), res.success ? "success" : "error");
    } catch (e) {
      setStatus("excel", `Execution error: ${e}`, "error");
      logElem.innerText = `[ERROR] ${e}`;
    } finally {
      setButtonsDisabled("view-excel", false);
    }
  });
}

/**
 * Attach to the currently open Excel workbook via COM.
 */
export async function excelConnectActive() {
  setStatus("excel", "Connecting to active Excel workbook…", "info", true);
  try {
    const res = await window.pywebview.api.excel_connect_active();
    handleExcelLoaded(res);
  } catch (err) {
    setStatus("excel", `Connection failed: ${err}`, "error");
  }
}

/**
 * Open a workbook from a file-picker dialog.
 */
export async function excelOpenFile() {
  setStatus("excel", "Selecting workbook…", "info", true);
  try {
    const res = await window.pywebview.api.excel_open_file();
    handleExcelLoaded(res);
  } catch (err) {
    setStatus("excel", `Open failed: ${err}`, "error");
  }
}

/**
 * Populate the Explorer and Context tabs after a workbook is loaded.
 * @param {object} res - API response containing sheet metadata and context text.
 */
function handleExcelLoaded(res) {
  if (res.cancelled) {
    setStatus("excel", "File open cancelled.", "info");
    return;
  }
  if (!res.success) {
    setStatus("excel", res.error, "error");
    return;
  }

  // Update metric cards
  document.getElementById("excel-metric-file").innerText     = res.file_name;
  document.getElementById("excel-metric-sheets").innerText   = res.sheet_count;
  document.getElementById("excel-metric-tables").innerText   = res.table_count;
  document.getElementById("excel-metric-formulas").innerText = res.formula_count;

  // Render sheet cards
  const list = document.getElementById("excel-sheets-list");
  list.innerHTML = "";

  res.sheets.forEach(sheet => {
    const card = document.createElement("div");
    card.className = "sheet-item";
    card.innerHTML = `
      <div class="sheet-item-header">
        <span><i data-lucide="table-2" style="width:13px;height:13px;vertical-align:middle;margin-right:4px;"></i>${sheet.name}</span>
        <span>${sheet.row_count} rows × ${sheet.col_count} cols</span>
      </div>
      <div class="sheet-item-meta">
        ${sheet.tables_count} tables · ${sheet.formulas_count} formulas
      </div>
      <div class="sheet-headers-tag">
        ${sheet.headers.map(h => `<span class="header-tag">${h}</span>`).join("")}
      </div>
    `;
    list.appendChild(card);
  });

  if (window.lucide) lucide.createIcons();

  document.getElementById("excel-context-text").value = res.context_text || "";
  setStatus("excel", res.message, "success");
}
