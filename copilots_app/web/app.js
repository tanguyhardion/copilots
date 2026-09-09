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
    badgeSubtleColor: "var(--brand-ppt-subtle)",
    badgeGlowColor: "var(--brand-ppt-glow)",
    actions: [
      { id: "action-prompt", text: "System Prompt", icon: "settings" },
      { id: "action-help", text: "", icon: "help-circle", title: "PowerPoint Copilot Guide", isIconOnly: true },
    ]
  },
  word: {
    title: "Word Copilot",
    subtitle: "Document generator, live cursor injector, and active document editor with DSL extraction",
    icon: "../../assets/icons/word.png",
    badge: "Word COM",
    badgeColor: "var(--brand-word)",
    badgeSubtleColor: "var(--brand-word-subtle)",
    badgeGlowColor: "var(--brand-word-glow)",
    actions: [
      { id: "action-prompt", text: "System Prompt", icon: "settings" },
      { id: "action-help", text: "", icon: "help-circle", title: "Word Copilot Guide", isIconOnly: true },
    ]
  },
  excel: {
    title: "Excel Copilot",
    subtitle: "Active workbook COM automation, real-time spreadsheet analysis, and deterministic JSON action execution",
    icon: "../../assets/icons/excel.png",
    badge: "Excel COM",
    badgeColor: "var(--brand-excel)",
    badgeSubtleColor: "var(--brand-excel-subtle)",
    badgeGlowColor: "var(--brand-excel-glow)",
    actions: [
      { id: "action-prompt", text: "System Prompt", icon: "settings" },
      { id: "action-help", text: "", icon: "help-circle", title: "Excel Copilot Guide", isIconOnly: true },
    ]
  },
  cv: {
    title: "CV Copilot",
    subtitle: "Deterministic Data Quality engine, Europass profile validation, and formatted Word .docx generator",
    icon: "../../assets/icons/cv.png",
    badge: "DQ Engine + docx",
    badgeColor: "var(--brand-cv)",
    badgeSubtleColor: "var(--brand-cv-subtle)",
    badgeGlowColor: "var(--brand-cv-glow)",
    actions: [
      { id: "action-prompt", text: "System Prompt", icon: "settings" },
      { id: "action-help", text: "", icon: "help-circle", title: "CV Copilot Guide", isIconOnly: true },
    ]
  },
  python: {
    title: "Python Copilot",
    subtitle: "Interactive execution sandbox, script runner, live stdout/stderr console, and directory context inspector",
    icon: "../../assets/icons/python.png",
    badge: "Python Runner Sandbox",
    badgeColor: "var(--brand-python)",
    badgeSubtleColor: "var(--brand-python-subtle)",
    badgeGlowColor: "var(--brand-python-glow)",
    actions: [
      { id: "action-prompt", text: "System Prompt", icon: "settings" },
      { id: "action-help", text: "", icon: "help-circle", title: "Python Copilot Guide", isIconOnly: true },
    ]
  },
  organizer: {
    title: "Folder/File Organizer Copilot",
    subtitle: "Analyze recursive folder context, prepare LLM-ready summaries, and execute safe copy-only reorganization plans",
    icon: "../../assets/icons/folder-organizer.png",
    badge: "Folder DSL + Safe Copy",
    badgeColor: "var(--brand-organizer)",
    badgeSubtleColor: "var(--brand-organizer-subtle)",
    badgeGlowColor: "var(--brand-organizer-glow)",
    actions: [
      { id: "action-prompt", text: "System Prompt", icon: "settings" },
      { id: "action-help", text: "", icon: "help-circle", title: "Folder Organizer Copilot Guide", isIconOnly: true },
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
  setupPythonView();
  setupOrganizerView();
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
  document.documentElement.style.setProperty("--copilot-accent-subtle", routeMeta.badgeSubtleColor || "var(--primary-subtle)");
  document.documentElement.style.setProperty("--copilot-accent-glow", routeMeta.badgeGlowColor || "var(--primary-glow)");

  const badge = document.getElementById("header-badge");
  badge.innerText = routeMeta.badge;
  badge.style.backgroundColor = routeMeta.badgeColor;

  // Render header actions
  const actionsContainer = document.getElementById("header-actions");
  actionsContainer.innerHTML = "";

  routeMeta.actions.forEach(act => {
    const btn = document.createElement("button");
    btn.className = act.isIconOnly ? "btn btn-secondary btn-icon-only" : "btn btn-secondary";
    btn.id = act.id;
    if (act.title) {
      btn.title = act.title;
      btn.setAttribute("aria-label", act.title);
    }
    if (act.isIconOnly) {
      btn.innerHTML = `<i data-lucide="${act.icon}" class="btn-icon"></i>`;
    } else {
      btn.innerHTML = `<i data-lucide="${act.icon}" class="btn-icon"></i> ${act.text}`;
    }
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
  } else if (actionId === "action-help") {
    openHelpModal(routeId);
  } else if (actionId === "action-excel-active") {
    excelConnectActive();
  } else if (actionId === "action-excel-open") {
    excelOpenFile();
  } else if (actionId === "action-python-open-folder") {
    pythonOpenExplorer();
  } else if (actionId === "action-python-copy-context") {
    pythonCopyFolderContext();
  } else if (actionId === "action-python-clear") {
    pythonClearWorkspace();
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
  const editor = document.getElementById("ppt-editor");

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
  const editor = document.getElementById("word-editor");

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

  // Connect Active & Open Workbook
  document.getElementById("excel-btn-open").addEventListener("click", excelOpenFile);
  document.getElementById("excel-btn-active").addEventListener("click", excelConnectActive);

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
      setStatus("cv", `Word generation error: ${err}`, "error");
    } finally {
      setButtonsDisabled("view-cv", false);
    }
  });

  // Generate pptx (1-Slide Executive Proposal)
  const btnPptx = document.getElementById("cv-btn-generate-pptx");
  if (btnPptx) {
    btnPptx.addEventListener("click", async () => {
      setStatus("cv", "Generating 1-Slide Executive PowerPoint (.pptx) document…", "info", true);
      setButtonsDisabled("view-cv", true);
      try {
        const res = await window.pywebview.api.cv_generate_pptx(cvEditor.value);
        if (res.success) {
          setStatus("cv", res.message, "success");
        } else {
          setStatus("cv", res.error, "error");
        }
      } catch (err) {
        setStatus("cv", `PowerPoint generation error: ${err}`, "error");
      } finally {
        setButtonsDisabled("view-cv", false);
      }
    });
  }
}

function updateCVMetrics(cvData) {
  const p = cvData.personal_info || cvData.personal_information || {};
  const name = `${p.first_name || ""} ${p.last_name || ""}`.trim() || "Candidate";
  const expCount = (cvData.project_experience || cvData.work_experience || []).length;
  let skillsCount = 0;
  if (cvData.skills && Array.isArray(cvData.skills)) {
    skillsCount = cvData.skills.length;
  } else if (cvData.personal_skills && typeof cvData.personal_skills === "object") {
    Object.values(cvData.personal_skills).forEach(val => {
      if (Array.isArray(val)) skillsCount += val.length;
    });
  }
  document.getElementById("cv-metric-name").innerText = name;
  document.getElementById("cv-metric-exp").innerText = `${expCount} Positions`;
  document.getElementById("cv-metric-skills").innerText = `${skillsCount} Skills`;
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
   Python Copilot View
   ========================================================================= */
let pythonLastFiles = [];

async function setupPythonView() {
  const editor = document.getElementById("python-editor");
  const persistCheck = document.getElementById("python-persist-checkbox");
  const runBtn = document.getElementById("python-btn-run");
  const clearCodeBtn = document.getElementById("python-btn-clear-code");

  // Tabs
  const tabBtnConsole = document.getElementById("tab-btn-py-console");
  const tabBtnFiles = document.getElementById("tab-btn-py-files");
  const tabBtnContext = document.getElementById("tab-btn-py-context");
  const tabPageConsole = document.getElementById("tab-page-py-console");
  const tabPageFiles = document.getElementById("tab-page-py-files");
  const tabPageContext = document.getElementById("tab-page-py-context");

  function switchPyTab(activeBtn, activePage) {
    [tabBtnConsole, tabBtnFiles, tabBtnContext].forEach(b => b.classList.remove("active"));
    [tabPageConsole, tabPageFiles, tabPageContext].forEach(p => p.classList.remove("active"));
    activeBtn.classList.add("active");
    activePage.classList.add("active");
  }

  tabBtnConsole.addEventListener("click", () => switchPyTab(tabBtnConsole, tabPageConsole));
  tabBtnFiles.addEventListener("click", () => switchPyTab(tabBtnFiles, tabPageFiles));
  tabBtnContext.addEventListener("click", () => switchPyTab(tabBtnContext, tabPageContext));

  // Clear code button
  clearCodeBtn.addEventListener("click", () => {
    editor.value = "";
    editor.focus();
  });

  // Persistence toggle listener
  persistCheck.addEventListener("change", async (e) => {
    const persist = e.target.checked;
    try {
      if (window.pywebview?.api) {
        const res = await window.pywebview.api.python_set_persistence(persist);
        if (res.success) {
          updatePythonPersistUI(persist, res.folder_path);
          renderPythonFiles(res.files || []);
          setStatus("python", persist ? "Persistence ENABLED (files stored in AppData sandbox)" : "Persistence DISABLED (temporary sandbox auto-cleaned)", "info");
        }
      }
    } catch (err) {
      console.error("Failed to toggle persistence:", err);
    }
  });

  // Run button
  runBtn.addEventListener("click", pythonRunCode);

  // Explorer buttons
  document.getElementById("python-btn-open-explorer").addEventListener("click", pythonOpenExplorer);
  document.getElementById("python-btn-open-explorer-mini").addEventListener("click", pythonOpenExplorer);

  // Clear sandbox button
  document.getElementById("python-btn-clear-sandbox").addEventListener("click", pythonClearWorkspace);

  // Copy Context buttons
  document.getElementById("python-btn-copy-context").addEventListener("click", pythonCopyFolderContext);
  document.getElementById("python-btn-copy-context-mini").addEventListener("click", pythonCopyFolderContext);
  document.getElementById("python-btn-copy-context-full").addEventListener("click", pythonCopyFolderContext);

  // Initial status query
  if (window.pywebview?.api) {
    try {
      const status = await window.pywebview.api.python_get_status();
      if (status.success) {
        persistCheck.checked = status.persist;
        updatePythonPersistUI(status.persist, status.folder_path);
        renderPythonFiles(status.files || []);
      }
    } catch (e) {}
  }
}

function updatePythonPersistUI(persist, folderPath) {
  const persistMetric = document.getElementById("python-metric-persist");
  const folderLabel = document.getElementById("python-folder-label");

  if (persist) {
    persistMetric.innerText = "Persistent (Saved)";
    persistMetric.style.color = "var(--brand-excel)";
    folderLabel.innerText = `Persistent: ${folderPath}`;
    folderLabel.title = folderPath;
  } else {
    persistMetric.innerText = "Temp (Auto-Clean)";
    persistMetric.style.color = "var(--text-secondary)";
    folderLabel.innerText = `Sandbox: ${folderPath}`;
    folderLabel.title = folderPath;
  }
}

async function pythonRunCode() {
  const editor = document.getElementById("python-editor");
  const code = editor.value;
  const consoleElem = document.getElementById("python-log-console");
  const persist = document.getElementById("python-persist-checkbox").checked;

  if (!code.trim()) {
    setStatus("python", "Code editor is empty. Paste or write a Python script to run.", "warning");
    return;
  }

  setStatus("python", "Executing Python script in sandbox…", "info", true);
  setButtonsDisabled("view-python", true);

  consoleElem.innerHTML = `<span class="info-line"># Running script (sys.executable)...</span>\n`;

  try {
    const res = await window.pywebview.api.python_run_code(code, persist);

    // Render console lines
    let outHtml = "";
    if (res.stdout) {
      outHtml += `<span class="stdout-line">${escapeHtml(res.stdout)}</span>\n`;
    }
    if (res.stderr) {
      outHtml += `<span class="stderr-line">${escapeHtml(res.stderr)}</span>\n`;
    }
    if (!res.stdout && !res.stderr) {
      outHtml += `<span class="info-line"># Script executed with no terminal output.</span>\n`;
    }

    outHtml += `\n<span class="info-line"># Process finished with exit code ${res.exit_code} (${res.duration_ms} ms)</span>`;
    consoleElem.innerHTML = outHtml;
    consoleElem.scrollTop = consoleElem.scrollHeight;

    // Update metrics
    const statusMetric = document.getElementById("python-metric-status");
    if (res.exit_code === 0) {
      statusMetric.innerText = "✓ Success (0)";
      statusMetric.style.color = "var(--brand-excel)";
      setStatus("python", `Execution finished successfully in ${res.duration_ms} ms.`, "success");
    } else {
      statusMetric.innerText = `✗ Exit (${res.exit_code})`;
      statusMetric.style.color = "var(--error)";
      setStatus("python", `Execution exited with error code ${res.exit_code}.`, "error");
    }

    document.getElementById("python-metric-duration").innerText = `${res.duration_ms} ms`;

    // Render files
    renderPythonFiles(res.all_files || []);

    // Refresh context
    refreshPythonContext();

    // If new files were produced, switch briefly or notify
    if (res.produced_files && res.produced_files.length > 0) {
      setStatus("python", `✓ Success! Produced ${res.produced_files.length} file(s) in sandbox folder.`, "success");
    }
  } catch (err) {
    consoleElem.innerHTML += `\n<span class="stderr-line">Bridge execution error: ${escapeHtml(String(err))}</span>`;
    setStatus("python", `Bridge error: ${err}`, "error");
  } finally {
    setButtonsDisabled("view-python", false);
    if (window.lucide) lucide.createIcons();
  }
}

function renderPythonFiles(files) {
  pythonLastFiles = files;
  const list = document.getElementById("python-file-list");
  const countSpan = document.getElementById("tab-files-count");
  const metricFiles = document.getElementById("python-metric-files");

  countSpan.innerText = files.length;
  metricFiles.innerText = `${files.length} File${files.length === 1 ? '' : 's'}`;

  list.innerHTML = "";

  if (!files || files.length === 0) {
    list.innerHTML = `<div class="empty-state">No generated files in the workspace. Any documents created by your script will appear here.</div>`;
    return;
  }

  files.forEach(f => {
    const item = document.createElement("div");
    item.className = "python-file-item";

    // Choose icon and color
    let iconName = "file";
    let iconClass = "";
    const ext = (f.ext || "").toLowerCase();

    if (f.is_dir) {
      iconName = "folder";
      iconClass = "folder";
    } else if (["py", "json", "js", "html", "css", "sql"].includes(ext)) {
      iconName = "file-code";
      iconClass = "code";
    } else if (["doc", "docx", "txt", "md", "pdf"].includes(ext)) {
      iconName = "file-text";
      iconClass = "doc";
    } else if (["xls", "xlsx", "csv"].includes(ext)) {
      iconName = "table";
      iconClass = "sheet";
    }

    item.innerHTML = `
      <div class="python-file-info">
        <div class="python-file-icon ${iconClass}">
          <i data-lucide="${iconName}" style="width:15px;height:15px;"></i>
        </div>
        <div style="overflow:hidden;">
          <div class="python-file-name" title="${f.name}">${f.name}</div>
          <div class="python-file-meta">${f.size_formatted}</div>
        </div>
      </div>
      <div class="python-file-actions">
        <button class="btn-tiny" title="Open file in default app" data-action="open" data-name="${f.name}">
          <i data-lucide="external-link" style="width:11px;height:11px;"></i> Open
        </button>
        <button class="btn-tiny" title="Copy filename" data-action="copy-name" data-name="${f.name}">
          <i data-lucide="copy" style="width:11px;height:11px;"></i>
        </button>
      </div>
    `;

    // Action handlers
    item.querySelector('[data-action="open"]').addEventListener("click", () => {
      pythonOpenFile(f.name);
    });

    item.querySelector('[data-action="copy-name"]').addEventListener("click", () => {
      navigator.clipboard.writeText(f.name).then(() => {
        setStatus("python", `Copied '${f.name}' to clipboard`, "info");
      });
    });

    list.appendChild(item);
  });

  if (window.lucide) lucide.createIcons();
}

async function refreshPythonContext() {
  try {
    if (window.pywebview?.api) {
      const res = await window.pywebview.api.python_get_folder_context();
      if (res.success) {
        document.getElementById("python-context-text").value = res.context_markdown;
      }
    }
  } catch (e) {}
}

async function pythonOpenExplorer() {
  try {
    if (window.pywebview?.api) {
      const res = await window.pywebview.api.python_open_folder();
      if (res.success) {
        setStatus("python", `Opened sandbox directory in Windows Explorer.`, "success");
      } else {
        setStatus("python", `Could not open folder: ${res.error}`, "error");
      }
    }
  } catch (err) {
    setStatus("python", `Explorer error: ${err}`, "error");
  }
}

async function pythonOpenFile(filename) {
  try {
    if (window.pywebview?.api) {
      const res = await window.pywebview.api.python_open_file(filename);
      if (res.success) {
        setStatus("python", `Opened ${filename}`, "success");
      } else {
        setStatus("python", res.error, "error");
      }
    }
  } catch (err) {
    setStatus("python", `Open file error: ${err}`, "error");
  }
}

async function pythonCopyFolderContext() {
  try {
    if (window.pywebview?.api) {
      const res = await window.pywebview.api.python_get_folder_context();
      if (res.success && res.context_markdown) {
        await navigator.clipboard.writeText(res.context_markdown);
        document.getElementById("python-context-text").value = res.context_markdown;
        setStatus("python", `✓ Folder context with ${res.file_count} file(s) copied to clipboard! Ready to paste into ChatGPT/LLM.`, "success");
      } else {
        setStatus("python", "Failed to retrieve folder context.", "error");
      }
    }
  } catch (err) {
    setStatus("python", `Copy context error: ${err}`, "error");
  }
}

async function pythonClearWorkspace() {
  if (!confirm("Are you sure you want to clear all files in the current Python sandbox?")) return;
  try {
    if (window.pywebview?.api) {
      const res = await window.pywebview.api.python_clear_sandbox();
      if (res.success) {
        renderPythonFiles([]);
        document.getElementById("python-log-console").innerHTML = `<span class="info-line"># Workspace cleared.</span>\n`;
        refreshPythonContext();
        setStatus("python", "Sandbox workspace cleared.", "success");
      } else {
        setStatus("python", res.error, "error");
      }
    }
  } catch (err) {
    setStatus("python", `Clear error: ${err}`, "error");
  }
}

/* =========================================================================
   Folder/File Organizer Copilot View
   ========================================================================= */
async function setupOrganizerView() {
  const selectBtn = document.getElementById("organizer-btn-select-folder");
  const copyBtn = document.getElementById("organizer-btn-copy-context");
  const executeBtn = document.getElementById("organizer-btn-execute-plan");
  const openFolderBtn = document.getElementById("organizer-btn-open-folder");
  const planTextarea = document.getElementById("organizer-plan-text");

  if (!selectBtn || !copyBtn || !executeBtn) return;

  selectBtn.addEventListener("click", organizerSelectFolder);
  copyBtn.addEventListener("click", organizerCopyContext);
  executeBtn.addEventListener("click", organizerExecutePlan);
  if (openFolderBtn) openFolderBtn.addEventListener("click", organizerOpenOutputFolder);

  // Auto-validate on paste or typing (debounced)
  let _validateTimer = null;
  planTextarea.addEventListener("input", () => {
    clearTimeout(_validateTimer);
    _validateTimer = setTimeout(organizerValidatePlan, 500);
  });

  try {
    const res = await window.pywebview?.api?.organizer_get_status?.();
    if (res?.success) {
      updateOrganizerLabels(res.source_root, res.last_output_root);
      // Enable the open-folder button if there's an existing output
      const openFolderBtn = document.getElementById("organizer-btn-open-folder");
      if (openFolderBtn && res.last_output_root) openFolderBtn.disabled = false;
    }
  } catch (e) {}
}

function updateOrganizerLabels(sourceRoot, outputRoot) {
  const sourceLabel = document.getElementById("organizer-source-label");
  const outputLabel = document.getElementById("organizer-output-label");

  if (sourceLabel) {
    sourceLabel.innerText = sourceRoot ? `Source: ${sourceRoot}` : "No folder selected.";
    sourceLabel.title = sourceRoot || "";
  }
  if (outputLabel) {
    outputLabel.innerText = outputRoot ? `Output: ${outputRoot}` : "Output: not generated yet";
    outputLabel.title = outputRoot || "";
  }
}

async function organizerSelectFolder() {
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

async function organizerCopyContext() {
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

async function organizerValidatePlan() {
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

async function organizerExecutePlan() {
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
        // Enable the open-folder button now that we have an output
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

async function organizerOpenOutputFolder() {
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

function escapeHtml(text) {
  const map = {
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#039;'
  };
  return text.replace(/[&<>"']/g, m => map[m]);
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

  // Help modal close buttons
  document.getElementById("help-close").addEventListener("click", closeHelpModal);
  document.getElementById("help-btn-ok").addEventListener("click", closeHelpModal);

  // Click outside (backdrop) to close cheatsheet modal
  document.getElementById("cheatsheet-modal").addEventListener("click", (e) => {
    if (e.target === document.getElementById("cheatsheet-modal")) closeCheatsheetModal();
  });

  // Click outside (backdrop) to close help modal
  document.getElementById("help-modal").addEventListener("click", (e) => {
    if (e.target === document.getElementById("help-modal")) closeHelpModal();
  });

  // Escape key closes any open modal
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape") {
      closePromptModal();
      closeCheatsheetModal();
      closeHelpModal();
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

function closeHelpModal() {
  document.getElementById("help-modal").classList.remove("open");
}

function openHelpModal(routeId) {
  const modal = document.getElementById("help-modal");
  const title = document.getElementById("help-title");
  const badge = document.getElementById("help-badge");
  const icon = document.getElementById("help-modal-icon");
  const content = document.getElementById("help-content");

  const routeMeta = ROUTES[routeId] || { title: "Copilot" };
  title.innerText = `How to Use ${routeMeta.title}`;
  badge.innerText = `${routeMeta.badge || "Copilot"} · Workflow & Features`;
  if (icon) {
    icon.style.color = routeMeta.badgeColor || "var(--primary)";
  }

  if (routeId === "powerpoint") {
    content.innerHTML = `
      <div class="help-guide-section">
        <div class="help-section-title"><i data-lucide="compass"></i> End-to-End Workflow</div>
        <div class="help-steps-grid">
          <div class="help-step-card">
            <div class="help-step-header"><span class="help-step-num">1</span> System Prompt</div>
            <p>Click <strong>System Prompt</strong> on the top right. Copy the prompt to clipboard and paste it into your LLM (ChatGPT, Claude, Gemini, etc.) as your custom instructions or system prompt.</p>
          </div>
          <div class="help-step-card">
            <div class="help-step-header"><span class="help-step-num">2</span> Ask Your LLM</div>
            <p>Ask the LLM to design a slide deck, card grid, process chevron, architecture diagram, or comparison table. The LLM will output pure PowerPoint DSL.</p>
          </div>
          <div class="help-step-card">
            <div class="help-step-header"><span class="help-step-num">3</span> Paste &amp; Execute</div>
            <p>Paste the generated DSL into the code editor. Open Microsoft PowerPoint, select the active presentation/slide, and click one of the action buttons below.</p>
          </div>
        </div>
      </div>

      <div class="help-guide-section">
        <div class="help-section-title"><i data-lucide="layers"></i> Selecting the Active Document</div>
        <div class="help-feature-item">
          <div class="help-feature-desc">
            PowerPoint Copilot attaches live to your running <strong>Microsoft PowerPoint</strong> application via Windows COM. Whatever presentation and slide is currently active (or open in the forefront of PowerPoint) is where elements will be inserted. If PowerPoint is not open, you will receive a status notification prompting you to start PowerPoint.
          </div>
        </div>
      </div>

      <div class="help-guide-section">
        <div class="help-section-title"><i data-lucide="mouse-pointer"></i> Functions &amp; Buttons Explained</div>
        <div class="help-feature-list">
          <div class="help-feature-item">
            <span class="help-feature-btn-badge"><i data-lucide="settings"></i> System Prompt</span>
            <p class="help-feature-desc">Opens the system prompt modal. View the default prompt instructions that teach LLMs how to construct valid DSL, or customize and save persistent overrides.</p>
          </div>
          <div class="help-feature-item">
            <span class="help-feature-btn-badge"><i data-lucide="clipboard"></i> Copy to Clipboard</span>
            <p class="help-feature-desc">Compiles the DSL shapes into native PowerPoint drawing objects and places them onto the Windows Clipboard. You can then switch to any slide in PowerPoint and press <code>Ctrl+V</code>.</p>
          </div>
          <div class="help-feature-item">
            <span class="help-feature-btn-badge"><i data-lucide="plus-circle"></i> Insert on Current Slide</span>
            <p class="help-feature-desc">Directly injects the rendered shapes, text boxes, and tables onto whatever slide is currently active and visible in PowerPoint without replacing the slide background.</p>
          </div>
          <div class="help-feature-item">
            <span class="help-feature-btn-badge"><i data-lucide="presentation"></i> Create Full Slide(s)</span>
            <p class="help-feature-desc">Creates brand new blank 16:9 slides at the end of the active presentation and renders all shapes, cards, and headers across the new slides.</p>
          </div>
        </div>
      </div>
    `;
  } else if (routeId === "word") {
    content.innerHTML = `
      <div class="help-guide-section">
        <div class="help-section-title"><i data-lucide="compass"></i> End-to-End Workflow</div>
        <div class="help-steps-grid">
          <div class="help-step-card">
            <div class="help-step-header"><span class="help-step-num">1</span> System Prompt</div>
            <p>Click <strong>System Prompt</strong> at top right. Copy the instructions to your LLM. The prompt trains the model on Word DSL creation, paragraph styling, and precise edit operations.</p>
          </div>
          <div class="help-step-card">
            <div class="help-step-header"><span class="help-step-num">2</span> Select Document &amp; Plan</div>
            <p>To edit an existing document, click <strong>Extract DSL from Word</strong>. Feed the extracted structure to your LLM and ask for modifications (using <code>edit target=active</code> commands).</p>
          </div>
          <div class="help-step-card">
            <div class="help-step-header"><span class="help-step-num">3</span> Apply or Build</div>
            <p>Paste the LLM's response. Click <strong>Apply Edits to Doc</strong> to modify the open document, or click <strong>Build &amp; Open Document</strong> to create a fresh <code>.docx</code>.</p>
          </div>
        </div>
      </div>

      <div class="help-guide-section">
        <div class="help-section-title"><i data-lucide="file-text"></i> Selecting the Active Document</div>
        <div class="help-feature-item">
          <div class="help-feature-desc">
            Word Copilot connects directly to running <strong>Microsoft Word</strong> instances. When using <em>Extract DSL</em>, <em>Insert at Cursor</em>, or <em>Apply Edits</em>, the Copilot automatically targets the document in Word's currently active top window.
          </div>
        </div>
      </div>

      <div class="help-guide-section">
        <div class="help-section-title"><i data-lucide="mouse-pointer"></i> Functions &amp; Buttons Explained</div>
        <div class="help-feature-list">
          <div class="help-feature-item">
            <span class="help-feature-btn-badge"><i data-lucide="download"></i> Extract DSL from Word</span>
            <p class="help-feature-desc">Reads the active Microsoft Word document via COM, parses paragraphs, headings, tables, and lists, and outputs a clean structured DSL representation into the editor for LLM prompting.</p>
          </div>
          <div class="help-feature-item">
            <span class="help-feature-btn-badge"><i data-lucide="pen-tool"></i> Insert at Cursor</span>
            <p class="help-feature-desc">Renders the DSL in the editor and inserts it at the exact blinking cursor selection in your currently active Word window.</p>
          </div>
          <div class="help-feature-item">
            <span class="help-feature-btn-badge"><i data-lucide="sparkles"></i> Apply Edits to Doc</span>
            <p class="help-feature-desc">Executes deterministic in-place edits (like <code>replace</code>, <code>insert_before</code>, <code>insert_after</code>, <code>delete</code>) targeting paragraphs in the open Word document without losing styles.</p>
          </div>
          <div class="help-feature-item">
            <span class="help-feature-btn-badge"><i data-lucide="file-text"></i> Build &amp; Open Document</span>
            <p class="help-feature-desc">Compiles full standalone documents from the editor's DSL into an executive styled <code>.docx</code> file and opens it in Word.</p>
          </div>
        </div>
      </div>
    `;
  } else if (routeId === "excel") {
    content.innerHTML = `
      <div class="help-guide-section">
        <div class="help-section-title"><i data-lucide="compass"></i> End-to-End Workflow</div>
        <div class="help-steps-grid">
          <div class="help-step-card">
            <div class="help-step-header"><span class="help-step-num">1</span> Connect &amp; Analyze</div>
            <p>Click <strong>Connect Active</strong> to attach to an open Excel workbook, or <strong>Open Workbook</strong> to select a file. The analyzer inspects sheet schemas, tables, and formulas.</p>
          </div>
          <div class="help-step-card">
            <div class="help-step-header"><span class="help-step-num">2</span> Copy Context to LLM</div>
            <p>Switch to the <strong>LLM Prompt Context</strong> tab and click <em>Copy Context</em>. Paste this alongside your prompt into your LLM to request data manipulations.</p>
          </div>
          <div class="help-step-card">
            <div class="help-step-header"><span class="help-step-num">3</span> Validate &amp; Execute</div>
            <p>Paste the LLM's JSON Action Protocol. Validate syntax, ensure the backup checkbox is checked, and click <strong>Execute Protocol</strong>.</p>
          </div>
        </div>
      </div>

      <div class="help-guide-section">
        <div class="help-section-title"><i data-lucide="table"></i> Selecting the Active Document</div>
        <div class="help-feature-item">
          <div class="help-feature-desc">
            Use <strong>Connect Active</strong> in the bottom action bar to target the workbook that is currently open in Microsoft Excel. Alternatively, use <strong>Open Workbook</strong> to browse for any <code>.xlsx</code> or <code>.xlsm</code> file on your computer, which will launch and attach to it automatically.
          </div>
        </div>
      </div>

      <div class="help-guide-section">
        <div class="help-section-title"><i data-lucide="mouse-pointer"></i> Functions &amp; Buttons Explained</div>
        <div class="help-feature-list">
          <div class="help-feature-item">
            <span class="help-feature-btn-badge"><i data-lucide="link"></i> Connect Active</span>
            <p class="help-feature-desc">Attaches to the active Excel window, extracts live table headers, sheets, and formula counts, and populates the Explorer.</p>
          </div>
          <div class="help-feature-item">
            <span class="help-feature-btn-badge"><i data-lucide="folder-open"></i> Open Workbook</span>
            <p class="help-feature-desc">Presents a Windows file picker to open an existing spreadsheet file directly in Excel and generate semantic schema context.</p>
          </div>
          <div class="help-feature-item">
            <span class="help-feature-btn-badge"><i data-lucide="clipboard"></i> Copy Context</span>
            <p class="help-feature-desc">Copies the full markdown schema (sheet names, column headers, data types, sample values) formatted specifically for LLM understanding.</p>
          </div>
          <div class="help-feature-item">
            <span class="help-feature-btn-badge"><i data-lucide="check-circle"></i> Validate Protocol</span>
            <p class="help-feature-desc">Verifies that the JSON in the editor matches the expected Action Protocol schema with valid intents, targets, and operations.</p>
          </div>
          <div class="help-feature-item">
            <span class="help-feature-btn-badge"><i data-lucide="play"></i> Execute Protocol</span>
            <p class="help-feature-desc">Runs the deterministic batch actions (writing values, creating tables, formatting ranges, setting formulas) against Excel, creating a safety backup first if selected.</p>
          </div>
        </div>
      </div>
    `;
  } else if (routeId === "cv") {
    content.innerHTML = `
      <div class="help-guide-section">
        <div class="help-section-title"><i data-lucide="compass"></i> End-to-End Workflow</div>
        <div class="help-steps-grid">
          <div class="help-step-card">
            <div class="help-step-header"><span class="help-step-num">1</span> System Prompt</div>
            <p>Click <strong>System Prompt</strong>. Copy the CV prompt into your LLM. Provide raw candidate notes, text resumes, or LinkedIn profiles to have the LLM format standard structured JSON.</p>
          </div>
          <div class="help-step-card">
            <div class="help-step-header"><span class="help-step-num">2</span> Audit Data Quality</div>
            <p>Paste the JSON into the profile editor and click <strong>Audit Data Quality</strong>. The engine evaluates 12+ deterministic rules (dates, emails, skill tags, descriptions, languages).</p>
          </div>
          <div class="help-step-card">
            <div class="help-step-header"><span class="help-step-num">3</span> Compile Word or PPTX</div>
            <p>Generate either a <strong>1-Slide Executive PowerPoint (.pptx)</strong> proposal or a comprehensive multi-page <strong>Europass Word (.docx)</strong> document.</p>
          </div>
        </div>
      </div>

      <div class="help-guide-section">
        <div class="help-section-title"><i data-lucide="file-check"></i> Output Document Formats</div>
        <div class="help-feature-list">
          <div class="help-feature-item">
            <div class="help-feature-desc">
              <strong>1-Slide Executive Proposal (.pptx)</strong>: Designed specifically for RFP bids, client proposals, and tender decks based on <code>template.pptx</code>. Condenses the profile into a high-impact single slide with Contact &amp; Role header, executive summary, 6 core subject matter skills, and top 5 relevant projects.
            </div>
          </div>
          <div class="help-feature-item">
            <div class="help-feature-desc">
              <strong>Europass Word Document (.docx)</strong>: Generates an exhaustive, multi-page curriculum vitae fully compliant with European Commission standards, complete with full employment history, project breakdowns, education, languages, and certifications.
            </div>
          </div>
        </div>
      </div>

      <div class="help-guide-section">
        <div class="help-section-title"><i data-lucide="mouse-pointer"></i> Functions &amp; Buttons Explained</div>
        <div class="help-feature-list">
          <div class="help-feature-item">
            <span class="help-feature-btn-badge">Format JSON</span>
            <p class="help-feature-desc">Pretty-prints and validates the JSON document in the editor, updating the top candidate metrics.</p>
          </div>
          <div class="help-feature-item">
            <span class="help-feature-btn-badge"><i data-lucide="shield-check"></i> Audit Data Quality</span>
            <p class="help-feature-desc">Executes the Data Quality (DQ) rulebook, outputting a composite compliance score, pass/fail state, and actionable issues for missing fields or malformed data.</p>
          </div>
          <div class="help-feature-item">
            <span class="help-feature-btn-badge"><i data-lucide="presentation"></i> Generate 1-Slide PowerPoint (.pptx)</span>
            <p class="help-feature-desc">Populates <code>template.pptx</code> in-place with candidate data, preserving exact geometry, coordinates, and typography, then opens the slide in PowerPoint.</p>
          </div>
          <div class="help-feature-item">
            <span class="help-feature-btn-badge"><i data-lucide="file-check"></i> Generate Europass Word (.docx)</span>
            <p class="help-feature-desc">Generates a Microsoft Word document matching the standard Europass layout with clean margins, timeline tables, and skills matrix.</p>
          </div>
        </div>
      </div>
    `;
  } else if (routeId === "python") {
    content.innerHTML = `
      <div class="help-guide-section">
        <div class="help-section-title"><i data-lucide="compass"></i> End-to-End Workflow</div>
        <div class="help-steps-grid">
          <div class="help-step-card">
            <div class="help-step-header"><span class="help-step-num">1</span> Prompt Your LLM</div>
            <p>Click <strong>System Prompt</strong> to copy Python Copilot's prompt. Ask your LLM to write scripts to process data, generate charts, convert files, or manipulate documents.</p>
          </div>
          <div class="help-step-card">
            <div class="help-step-header"><span class="help-step-num">2</span> Run Code in Sandbox</div>
            <p>Paste the Python code into the runner editor and click <strong>Run Python Code</strong>. Execution runs asynchronously with live stdout/stderr streams.</p>
          </div>
          <div class="help-step-card">
            <div class="help-step-header"><span class="help-step-num">3</span> Inspect Artifacts</div>
            <p>View produced files in the <strong>Generated Files</strong> tab. Click <em>Open</em> to view any generated file, or copy folder context back to your LLM for multi-step tasks.</p>
          </div>
        </div>
      </div>

      <div class="help-guide-section">
        <div class="help-section-title"><i data-lucide="folder"></i> Sandbox Directory &amp; File Persistence</div>
        <div class="help-feature-item">
          <div class="help-feature-desc">
            By default, Python Copilot runs scripts in a sandboxed session. Checking <strong>Persist generated files &amp; scripts</strong> saves files in your user AppData directory so work is retained between sessions.
          </div>
        </div>
      </div>

      <div class="help-guide-section">
        <div class="help-section-title"><i data-lucide="mouse-pointer"></i> Functions &amp; Buttons Explained</div>
        <div class="help-feature-list">
          <div class="help-feature-item">
            <span class="help-feature-btn-badge"><i data-lucide="play"></i> Run Python Code</span>
            <p class="help-feature-desc">Executes the code editor's script inside the current sandbox environment, recording stdout, stderr, execution time, and new files.</p>
          </div>
          <div class="help-feature-item">
            <span class="help-feature-btn-badge"><i data-lucide="rotate-ccw"></i> Clear Editor</span>
            <p class="help-feature-desc">Quickly resets the Python script editor to write or paste a new script.</p>
          </div>
          <div class="help-feature-item">
            <span class="help-feature-btn-badge"><i data-lucide="folder-open"></i> Open Folder / Explorer</span>
            <p class="help-feature-desc">Opens the active sandbox working directory in Windows File Explorer so you can view all outputs or add input files.</p>
          </div>
          <div class="help-feature-item">
            <span class="help-feature-btn-badge"><i data-lucide="copy"></i> Copy Folder Context</span>
            <p class="help-feature-desc">Generates and copies a structured markdown directory listing to clipboard, making it effortless to prompt an LLM about files in your workspace.</p>
          </div>
          <div class="help-feature-item">
            <span class="help-feature-btn-badge"><i data-lucide="trash-2"></i> Clear Sandbox</span>
            <p class="help-feature-desc">Purges all generated files and artifacts in the sandbox workspace to start fresh.</p>
          </div>
        </div>
      </div>
    `;
  } else if (routeId === "organizer") {
    content.innerHTML = `
      <div class="help-guide-section">
        <div class="help-section-title"><i data-lucide="compass"></i> End-to-End Workflow</div>
        <div class="help-steps-grid">
          <div class="help-step-card">
            <div class="help-step-header"><span class="help-step-num">1</span> Select Source Folder</div>
            <p>Click <strong>Select Folder</strong> to choose the source tree to analyze. The original folder is always read-only from this copilot.</p>
          </div>
          <div class="help-step-card">
            <div class="help-step-header"><span class="help-step-num">2</span> Copy LLM Context</div>
            <p>Use <strong>Copy for LLM</strong> to generate a full recursive tree plus non-image file excerpts. Paste this into your LLM to request a reorganization plan.</p>
          </div>
          <div class="help-step-card">
            <div class="help-step-header"><span class="help-step-num">3</span> Execute Copy Plan</div>
            <p>Paste DSL instructions and click <strong>Execute Copy Plan</strong>. Files are copied into a separate output folder; source files are never moved or deleted.</p>
          </div>
        </div>
      </div>
      <div class="help-guide-section">
        <div class="help-section-title"><i data-lucide="code"></i> DSL</div>
        <div class="help-feature-item">
          <div class="help-feature-desc">
            One instruction per line: <code>MOVE "source/relative/path.ext" -> "dest/relative/path.ext"</code><br/>
            Comments supported with <code>#</code> or <code>//</code>. Invalid lines are reported and skipped.
          </div>
        </div>
      </div>
    `;
  }

  modal.classList.add("open");
  if (window.lucide) lucide.createIcons();
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
