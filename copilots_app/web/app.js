/**
 * Frontend Application Controller for Copilots Suite (pywebview bridge)
 */

// Route configurations
const ROUTES = {
  powerpoint: {
    title: "PowerPoint Copilot",
    subtitle: "Generate slides, shapes, rich text, tables, and icons from clean DSL",
    icon: "../../assets/icons/powerpoint.png",
    badge: "PowerPoint COM",
    badgeColor: "var(--brand-ppt)",
    actions: [
      { id: "action-prompt", text: "System Prompt", icon: "settings" },
      { id: "action-cheatsheet", text: "DSL Cheatsheet", icon: "book-open" },
    ]
  },
  word: {
    title: "Word Copilot",
    subtitle: "Document generator, live cursor injector, and active document editor with DSL extraction",
    icon: "../../assets/icons/word.png",
    badge: "Word COM",
    badgeColor: "var(--brand-word)",
    actions: [
      { id: "action-prompt", text: "System Prompt", icon: "settings" },
      { id: "action-cheatsheet", text: "Word DSL Cheatsheet", icon: "book-open" },
    ]
  },
  excel: {
    title: "Excel Copilot",
    subtitle: "Active workbook COM automation, real-time spreadsheet analysis, and deterministic JSON action execution",
    icon: "../../assets/icons/excel.png",
    badge: "Excel COM (pywin32)",
    badgeColor: "var(--brand-excel)",
    actions: [
      { id: "action-prompt", text: "System Prompt", icon: "settings" },
      { id: "action-excel-active", text: "Connect Active", icon: "link" },
      { id: "action-excel-open", text: "Open Workbook", icon: "folder-open" },
      { id: "action-excel-demo", text: "Open Demo", icon: "flask-conical" },
    ]
  },
  cv: {
    title: "CV Copilot",
    subtitle: "Deterministic Data Quality engine, Europass profile validation, and formatted Word .docx generator",
    icon: "../../assets/icons/cv.png",
    badge: "DQ Engine + docx",
    badgeColor: "var(--brand-cv)",
    actions: [
      { id: "action-prompt", text: "System Prompt", icon: "settings" },
      { id: "action-cv-reset", text: "Reset Sample CV", icon: "rotate-ccw" },
    ]
  }
};

let currentRoute = "powerpoint";
let activeModalPromptKey = "powerpoint";

// Wait for pywebview API to be ready
window.addEventListener("pywebviewready", () => {
  console.log("pywebview API ready!");
  initApp();
});

// Fallback in case opened in regular browser during development
setTimeout(() => {
  if (!window.__initialized) initApp();
}, 600);

function initApp() {
  if (window.__initialized) return;
  window.__initialized = true;

  setupSidebar();
  setupPowerPointView();
  setupWordView();
  setupExcelView();
  setupCVView();
  setupModals();

  navigateTo("powerpoint");

  // Render all static Lucide icons
  if (window.lucide) lucide.createIcons();
}

/* =========================================================================
   Sidebar & Navigation
   ========================================================================= */
function setupSidebar() {
  document.querySelectorAll(".nav-item").forEach(item => {
    item.addEventListener("click", () => {
      const route = item.getAttribute("data-route");
      navigateTo(route);
    });
  });
}

function navigateTo(routeId) {
  if (!ROUTES[routeId]) return;
  currentRoute = routeId;

  // Update Nav selection
  document.querySelectorAll(".nav-item").forEach(item => {
    if (item.getAttribute("data-route") === routeId) {
      item.classList.add("active");
    } else {
      item.classList.remove("active");
    }
  });

  // Update View Panels
  document.querySelectorAll(".view-panel").forEach(panel => {
    if (panel.id === `view-${routeId}`) {
      panel.classList.add("active");
    } else {
      panel.classList.remove("active");
    }
  });

  const routeMeta = ROUTES[routeId];
  document.getElementById("header-title").innerText = routeMeta.title;
  document.getElementById("header-subtitle").innerText = routeMeta.subtitle;
  document.getElementById("header-icon").src = routeMeta.icon;

  // Set active copilot theme variables
  document.documentElement.style.setProperty("--copilot-accent", routeMeta.badgeColor);
  document.documentElement.style.setProperty("--copilot-accent-subtle", `var(--brand-${routeId}-subtle)`);

  const badge = document.getElementById("header-badge");
  badge.innerText = routeMeta.badge;
  badge.style.backgroundColor = routeMeta.badgeColor;

  // Render header actions
  const actionsContainer = document.getElementById("header-actions");
  actionsContainer.innerHTML = "";

  routeMeta.actions.forEach(act => {
    const btn = document.createElement("button");
    btn.className = "btn btn-secondary";
    btn.id = act.id;
    btn.innerHTML = `<i data-lucide="${act.icon}" class="btn-icon"></i> ${act.text}`;
    btn.addEventListener("click", () => handleHeaderAction(act.id, routeId));
    actionsContainer.appendChild(btn);
  });

  // Re-render Lucide icons for dynamically inserted buttons
  if (window.lucide) lucide.createIcons();
}

function handleHeaderAction(actionId, routeId) {
  if (actionId === "action-prompt") {
    openPromptModal(routeId);
  } else if (actionId === "action-cheatsheet") {
    openCheatsheetModal(routeId);
  } else if (actionId === "action-excel-active") {
    excelConnectActive();
  } else if (actionId === "action-excel-open") {
    excelOpenFile();
  } else if (actionId === "action-excel-demo") {
    excelOpenDemo();
  } else if (actionId === "action-cv-reset") {
    cvResetSample();
  }
}

/* =========================================================================
   Status Bar Helper
   ========================================================================= */
function setStatus(viewId, message, level = "info", loading = false) {
  const bar = document.getElementById(`${viewId}-status`);
  if (!bar) return;

  bar.className = `status-bar ${level} ${loading ? 'loading' : ''}`;
  const textElem = bar.querySelector(".status-text");
  if (textElem) textElem.innerText = message;
}

/* =========================================================================
   PowerPoint View
   ========================================================================= */
async function setupPowerPointView() {
  const select = document.getElementById("ppt-samples-select");
  const editor = document.getElementById("ppt-editor");

  // Load samples
  if (window.pywebview?.api) {
    try {
      const samples = await window.pywebview.api.get_ppt_samples();
      select.innerHTML = "";
      for (const [name, dsl] of Object.entries(samples)) {
        const opt = document.createElement("option");
        opt.value = name;
        opt.textContent = name;
        select.appendChild(opt);
      }

      if (samples["Overview Demo"]) {
        editor.value = samples["Overview Demo"];
      } else if (Object.values(samples).length > 0) {
        editor.value = Object.values(samples)[0];
      }

      select.addEventListener("change", () => {
        const val = select.value;
        if (samples[val]) {
          editor.value = samples[val];
          setStatus("ppt", `Loaded sample template: '${val}'`, "info");
        }
      });
    } catch (e) {
      console.warn("Could not load ppt samples:", e);
    }
  }

  // Copy to clipboard button
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
      if (res.success) {
        setStatus("ppt", res.message, "success");
      } else {
        setStatus("ppt", res.error, "error");
      }
    } catch (err) {
      setStatus("ppt", `Error: ${err}`, "error");
    } finally {
      setButtonsDisabled("view-powerpoint", false);
    }
  });

  // Insert on current slide
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
      if (res.success) {
        setStatus("ppt", res.message, "success");
      } else {
        setStatus("ppt", res.error, "error");
      }
    } catch (err) {
      setStatus("ppt", `Error: ${err}`, "error");
    } finally {
      setButtonsDisabled("view-powerpoint", false);
    }
  });

  // Create full slide(s)
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
      if (res.success) {
        setStatus("ppt", res.message, "success");
      } else {
        setStatus("ppt", res.error, "error");
      }
    } catch (err) {
      setStatus("ppt", `Error: ${err}`, "error");
    } finally {
      setButtonsDisabled("view-powerpoint", false);
    }
  });
}

/* =========================================================================
   Word View
   ========================================================================= */
async function setupWordView() {
  const select = document.getElementById("word-samples-select");
  const editor = document.getElementById("word-editor");

  if (window.pywebview?.api) {
    try {
      const samples = await window.pywebview.api.get_word_samples();
      select.innerHTML = "";
      for (const [name, dsl] of Object.entries(samples)) {
        const opt = document.createElement("option");
        opt.value = name;
        opt.textContent = name;
        select.appendChild(opt);
      }

      if (samples["Complete Demo"]) {
        editor.value = samples["Complete Demo"];
      } else if (Object.values(samples).length > 0) {
        editor.value = Object.values(samples)[0];
      }

      select.addEventListener("change", () => {
        const val = select.value;
        if (samples[val]) {
          editor.value = samples[val];
          setStatus("word", `Loaded template: '${val}'`, "info");
        }
      });
    } catch (e) {
      console.warn("Could not load word samples:", e);
    }
  }

  // Build & Open
  document.getElementById("word-btn-build").addEventListener("click", async () => {
    const dsl = editor.value;
    setStatus("word", "Compiling Word document (.docx)…", "info", true);
    setButtonsDisabled("view-word", true);
    try {
      const res = await window.pywebview.api.word_build_and_open(dsl);
      if (res.success) {
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

  // Insert at cursor
  document.getElementById("word-btn-insert").addEventListener("click", async () => {
    const dsl = editor.value;
    setStatus("word", "Injecting elements at active Word cursor…", "info", true);
    setButtonsDisabled("view-word", true);
    try {
      const res = await window.pywebview.api.word_insert_at_cursor(dsl);
      if (res.success) {
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

  // Apply edits
  document.getElementById("word-btn-apply").addEventListener("click", async () => {
    const dsl = editor.value;
    setStatus("word", "Executing edit plan in active Word document…", "info", true);
    setButtonsDisabled("view-word", true);
    try {
      const res = await window.pywebview.api.word_apply_edits(dsl);
      if (res.success) {
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

  // Extract DSL
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

/* =========================================================================
   Excel View
   ========================================================================= */
async function setupExcelView() {
  const protocolEditor = document.getElementById("excel-protocol-editor");
  if (window.pywebview?.api) {
    try {
      const sampleJson = await window.pywebview.api.get_excel_sample_json();
      protocolEditor.value = sampleJson;
    } catch (e) {}
  }

  // Tabs
  const btnExplorer = document.getElementById("tab-btn-explorer");
  const btnContext = document.getElementById("tab-btn-context");
  const pageExplorer = document.getElementById("tab-page-explorer");
  const pageContext = document.getElementById("tab-page-context");

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

  // Copy context
  document.getElementById("excel-btn-copy-context").addEventListener("click", () => {
    const text = document.getElementById("excel-context-text").value;
    navigator.clipboard.writeText(text).then(() => {
      setStatus("excel", "Copied LLM prompt context to clipboard!", "success");
    });
  });

  // Validate JSON
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

  // Execute protocol
  document.getElementById("excel-btn-execute").addEventListener("click", async () => {
    const jsonStr = protocolEditor.value;
    const backup = document.getElementById("excel-backup-checkbox").checked;
    const logElem = document.getElementById("excel-log-console");

    setStatus("excel", "Executing action protocol…", "info", true);
    setButtonsDisabled("view-excel", true);
    try {
      const res = await window.pywebview.api.excel_execute_protocol(jsonStr, backup);
      if (res.logs && res.logs.length > 0) {
        logElem.innerText = res.logs.join("\n");
      }
      if (res.success) {
        setStatus("excel", res.message, "success");
      } else {
        setStatus("excel", res.error || res.message, "error");
      }
    } catch (e) {
      setStatus("excel", `Execution error: ${e}`, "error");
      logElem.innerText = `[ERROR] ${e}`;
    } finally {
      setButtonsDisabled("view-excel", false);
    }
  });
}

async function excelConnectActive() {
  setStatus("excel", "Connecting to active Excel workbook…", "info", true);
  try {
    const res = await window.pywebview.api.excel_connect_active();
    handleExcelLoaded(res);
  } catch (err) {
    setStatus("excel", `Connection failed: ${err}`, "error");
  }
}

async function excelOpenFile() {
  setStatus("excel", "Selecting workbook…", "info", true);
  try {
    const res = await window.pywebview.api.excel_open_file();
    handleExcelLoaded(res);
  } catch (err) {
    setStatus("excel", `Open failed: ${err}`, "error");
  }
}

async function excelOpenDemo() {
  setStatus("excel", "Loading demo portfolio…", "info", true);
  try {
    const res = await window.pywebview.api.excel_open_demo();
    handleExcelLoaded(res);
  } catch (err) {
    setStatus("excel", `Demo failed: ${err}`, "error");
  }
}

function handleExcelLoaded(res) {
  if (res.cancelled) {
    setStatus("excel", "File open cancelled.", "info");
    return;
  }
  if (!res.success) {
    setStatus("excel", res.error, "error");
    return;
  }

  // Update Metrics
  document.getElementById("excel-metric-file").innerText = res.file_name;
  document.getElementById("excel-metric-sheets").innerText = res.sheet_count;
  document.getElementById("excel-metric-tables").innerText = res.table_count;
  document.getElementById("excel-metric-formulas").innerText = res.formula_count;

  // Render Sheets
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

  // Prompt Context
  document.getElementById("excel-context-text").value = res.context_text || "";
  setStatus("excel", res.message, "success");
}

/* =========================================================================
   CV View
   ========================================================================= */
async function setupCVView() {
  const cvEditor = document.getElementById("cv-editor");

  // Load sample data
  if (window.pywebview?.api) {
    try {
      const res = await window.pywebview.api.cv_get_sample_data();
      if (res.success) {
        cvEditor.value = JSON.stringify(res.cv_data, null, 2);
        updateCVMetrics(res.cv_data);
      }
    } catch (e) {}
  }

  // Format JSON
  document.getElementById("cv-btn-format").addEventListener("click", () => {
    try {
      const parsed = JSON.parse(cvEditor.value);
      cvEditor.value = JSON.stringify(parsed, null, 2);
      updateCVMetrics(parsed);
      setStatus("cv", "JSON formatted successfully.", "info");
    } catch (e) {
      setStatus("cv", `Invalid JSON: ${e.message}`, "error");
    }
  });

  // Audit
  document.getElementById("cv-btn-audit").addEventListener("click", async () => {
    setStatus("cv", "Running deterministic Data Quality audit…", "info", true);
    setButtonsDisabled("view-cv", true);
    try {
      const res = await window.pywebview.api.cv_run_audit(cvEditor.value);
      if (res.success) {
        renderCVAuditResults(res.audit);
        setStatus("cv", `✓ Audit complete: DQ Score ${res.audit.dq_score}%`, "success");
      } else {
        setStatus("cv", res.error, "error");
      }
    } catch (err) {
      setStatus("cv", `Audit error: ${err}`, "error");
    } finally {
      setButtonsDisabled("view-cv", false);
    }
  });

  // Generate docx
  document.getElementById("cv-btn-generate").addEventListener("click", async () => {
    setStatus("cv", "Generating Europass Word (.docx) document…", "info", true);
    setButtonsDisabled("view-cv", true);
    try {
      const res = await window.pywebview.api.cv_generate_docx(cvEditor.value);
      if (res.success) {
        setStatus("cv", res.message, "success");
      } else {
        setStatus("cv", res.error, "error");
      }
    } catch (err) {
      setStatus("cv", `Generation error: ${err}`, "error");
    } finally {
      setButtonsDisabled("view-cv", false);
    }
  });
}

async function cvResetSample() {
  setStatus("cv", "Resetting to default Europass profile…", "info");
  try {
    const res = await window.pywebview.api.cv_get_sample_data();
    if (res.success) {
      document.getElementById("cv-editor").value = JSON.stringify(res.cv_data, null, 2);
      updateCVMetrics(res.cv_data);
      setStatus("cv", "Sample CV profile reset.", "success");
    }
  } catch (err) {
    setStatus("cv", `Reset error: ${err}`, "error");
  }
}

function updateCVMetrics(cvData) {
  const p = cvData.personal_information || {};
  const name = `${p.first_name || ""} ${p.last_name || ""}`.trim() || "Candidate";
  document.getElementById("cv-metric-name").innerText = name;
  document.getElementById("cv-metric-exp").innerText = `${(cvData.work_experience || []).length} Positions`;
  document.getElementById("cv-metric-skills").innerText = `${(cvData.skills || []).length} Skills`;
}

function renderCVAuditResults(audit) {
  // Update metric card score
  document.getElementById("cv-metric-score").innerText = `${audit.dq_score}%`;
  document.getElementById("cv-metric-name").innerText = audit.candidate_name;
  document.getElementById("cv-metric-exp").innerText = `${audit.experience_count} Positions`;
  document.getElementById("cv-metric-skills").innerText = `${audit.skills_count} Skills`;

  const container = document.getElementById("cv-audit-results");
  container.innerHTML = "";

  // Summary badge
  const summaryBadge = document.createElement("div");
  summaryBadge.className = `audit-summary-badge ${audit.passed ? 'passed' : 'warning'}`;
  summaryBadge.innerHTML = `
    <span>${audit.passed ? '✓ PASSED' : '⚠ ATTENTION REQUIRED'}</span>
    <span>Score: ${audit.dq_score}% (${audit.passed_checks}/${audit.total_checks} checks passed)</span>
  `;
  container.appendChild(summaryBadge);

  if (audit.issues.length === 0) {
    const cleanMsg = document.createElement("div");
    cleanMsg.className = "empty-state";
    cleanMsg.innerText = "All deterministic DQ checks passed without warnings or errors!";
    container.appendChild(cleanMsg);
    return;
  }

  audit.issues.forEach(issue => {
    const card = document.createElement("div");
    card.className = `audit-issue-card ${issue.severity.toLowerCase()}`;
    card.innerHTML = `
      <div style="display:flex; justify-content:space-between; font-weight:600; margin-bottom:2px;">
        <span>${issue.field}</span>
        <span style="font-size:10px; text-transform:uppercase;">${issue.severity}</span>
      </div>
      <div>${issue.message}</div>
    `;
    container.appendChild(card);
  });
}

/* =========================================================================
   Modals (System Prompts & Cheatsheets)
   ========================================================================= */
function setupModals() {
  // Prompt modal close buttons
  document.getElementById("modal-prompt-close").addEventListener("click", closePromptModal);
  document.getElementById("modal-prompt-btn-cancel").addEventListener("click", closePromptModal);

  // Click outside (backdrop) to close prompt modal
  document.getElementById("prompt-modal").addEventListener("click", (e) => {
    if (e.target === document.getElementById("prompt-modal")) closePromptModal();
  });

  // Copy prompt
  document.getElementById("modal-prompt-btn-copy").addEventListener("click", () => {
    const text = document.getElementById("modal-prompt-editor").value;
    navigator.clipboard.writeText(text).then(() => {
      setStatus(currentRoute, "System prompt copied to clipboard!", "success");
      closePromptModal();
    });
  });

  // Save prompt override
  document.getElementById("modal-prompt-btn-save").addEventListener("click", async () => {
    const text = document.getElementById("modal-prompt-editor").value;
    try {
      const res = await window.pywebview.api.save_prompt_override(activeModalPromptKey, text);
      if (res.success) {
        setStatus(currentRoute, res.message, "success");
        closePromptModal();
      } else {
        alert("Error saving prompt: " + res.error);
      }
    } catch (e) {
      alert("Save error: " + e);
    }
  });

  // Reset prompt default
  document.getElementById("modal-prompt-btn-reset").addEventListener("click", async () => {
    if (!confirm("Are you sure you want to reset this system prompt to the bundled default?")) return;
    try {
      const res = await window.pywebview.api.reset_prompt_default(activeModalPromptKey);
      if (res.success) {
        document.getElementById("modal-prompt-editor").value = res.content;
        document.getElementById("modal-prompt-badge").innerText = "Bundled Default";
        setStatus(currentRoute, res.message, "success");
      }
    } catch (e) {
      alert("Reset error: " + e);
    }
  });

  // Cheatsheet modal close buttons
  document.getElementById("cheatsheet-close").addEventListener("click", closeCheatsheetModal);
  document.getElementById("cheatsheet-btn-ok").addEventListener("click", closeCheatsheetModal);

  // Click outside (backdrop) to close cheatsheet modal
  document.getElementById("cheatsheet-modal").addEventListener("click", (e) => {
    if (e.target === document.getElementById("cheatsheet-modal")) closeCheatsheetModal();
  });

  // Escape key closes any open modal
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape") {
      closePromptModal();
      closeCheatsheetModal();
    }
  });
}

async function openPromptModal(copilotKey) {
  activeModalPromptKey = copilotKey;
  const modal = document.getElementById("prompt-modal");
  const editor = document.getElementById("modal-prompt-editor");
  const title = document.getElementById("modal-prompt-title");
  const badge = document.getElementById("modal-prompt-badge");

  try {
    const info = await window.pywebview.api.get_prompt_info(copilotKey);
    if (info.success) {
      title.innerText = info.title;
      editor.value = info.content;
      badge.innerText = info.is_custom ? "User Custom Override" : "Bundled Default";
      modal.classList.add("open");
    }
  } catch (err) {
    console.error("Failed to load prompt:", err);
  }
}

function closePromptModal() {
  document.getElementById("prompt-modal").classList.remove("open");
}

function closeCheatsheetModal() {
  document.getElementById("cheatsheet-modal").classList.remove("open");
}

function openCheatsheetModal(routeId) {
  const modal = document.getElementById("cheatsheet-modal");
  const title = document.getElementById("cheatsheet-title");
  const content = document.getElementById("cheatsheet-content");

  if (routeId === "powerpoint") {
    title.innerText = "PowerPoint DSL Cheatsheet & Syntax";
    content.innerHTML = `
      <h3>Slide Delimiter</h3>
      <pre>--- slide [title="Optional Slide Title"]</pre>

      <h3>Core Shapes</h3>
      <pre>rect x=1.0 y=1.5 w=4.0 h=2.0 fill=primary stroke=border_light text="Box Title"
rounded_rect x=5.5 y=1.5 w=4.0 h=2.0 rx=0.2 fill=bg_card text="Rounded Box"
chevron x=1.0 y=4.0 w=3.0 h=1.0 fill=brand_ppt text="Step 1" direction=right</pre>

      <h3>Rich Typography &amp; Paragraphs</h3>
      <pre>text x=1.0 y=1.0 w=10.0 h=1.0 | "Main Heading" size=24 bold=true color=primary
p size=14 color=text_secondary | "Multi-line descriptive body text goes here..."</pre>

      <h3>Tables</h3>
      <pre>table x=1.0 y=2.0 w=11.0 h=3.0 cols=30%,40%,30%
header="Phase","Milestone","Status"
row="1. Inception","Architecture Alignment","Complete"
row="2. Implementation","pywebview Migration","In Progress"</pre>
    `;
  } else if (routeId === "word") {
    title.innerText = "Word DSL &amp; Edit Engine Syntax";
    content.innerHTML = `
      <h3>Document Generation Mode</h3>
      <pre>page size=a4 orientation=portrait margin=54,54,54,54

h1 align=center color=a4 | "Document Title"
h3 align=center color=a2 | "Subtitle"
hr color=a3 weight=2

p spacing_after=8 | "Paragraph text with standard typography."

table width=100% header_fill=a4 header_text_color=#FFFFFF text_color=t1
cols=30%,70%
header="Property","Description"
row="Engine","pywebview Desktop"
row="Status","Production Ready"</pre>

      <h3>Active Document Edit Mode</h3>
      <pre>edit target=active
replace find="Old Text" replace="New Approved Text"
insert_before find="Summary" | "Note: Preceding disclaimer here."
insert_after find="Summary" | "Appendix follow-up."</pre>
    `;
  }

  modal.classList.add("open");
}

function setButtonsDisabled(containerId, disabled) {
  const container = document.getElementById(containerId);
  if (!container) return;
  container.querySelectorAll(".action-bar button, .header-actions button").forEach(btn => {
    btn.disabled = disabled;
    btn.style.opacity = disabled ? "0.6" : "1.0";
    btn.style.pointerEvents = disabled ? "none" : "auto";
  });
}
