/**
 * Python Copilot view controller.
 */

import { setStatus, setButtonsDisabled, escapeHtml } from "../core/utils.js";

/** Tracks the last known file list for the sandbox. */
let pythonLastFiles = [];

/**
 * Bind all action buttons and tab switchers in the Python view panel.
 * Also queries initial sandbox status from the Python API.
 */
export async function setupPythonView() {
  const editor       = document.getElementById("python-editor");
  const persistCheck = document.getElementById("python-persist-checkbox");
  const runBtn       = document.getElementById("python-btn-run");
  const clearCodeBtn = document.getElementById("python-btn-clear-code");

  // Tab switcher
  const tabBtnConsole  = document.getElementById("tab-btn-py-console");
  const tabBtnFiles    = document.getElementById("tab-btn-py-files");
  const tabBtnContext  = document.getElementById("tab-btn-py-context");
  const tabPageConsole = document.getElementById("tab-page-py-console");
  const tabPageFiles   = document.getElementById("tab-page-py-files");
  const tabPageContext = document.getElementById("tab-page-py-context");

  const switchPyTab = (activeBtn, activePage) => {
    [tabBtnConsole, tabBtnFiles, tabBtnContext].forEach(b => b.classList.remove("active"));
    [tabPageConsole, tabPageFiles, tabPageContext].forEach(p => p.classList.remove("active"));
    activeBtn.classList.add("active");
    activePage.classList.add("active");
  };

  tabBtnConsole.addEventListener("click", () => switchPyTab(tabBtnConsole, tabPageConsole));
  tabBtnFiles.addEventListener("click",   () => switchPyTab(tabBtnFiles,   tabPageFiles));
  tabBtnContext.addEventListener("click", () => switchPyTab(tabBtnContext,  tabPageContext));

  // Clear editor
  clearCodeBtn.addEventListener("click", () => {
    editor.value = "";
    editor.focus();
  });

  // Persistence toggle
  persistCheck.addEventListener("change", async (e) => {
    const persist = e.target.checked;
    try {
      if (window.pywebview?.api) {
        const res = await window.pywebview.api.python_set_persistence(persist);
        if (res.success) {
          updatePythonPersistUI(persist, res.folder_path);
          renderPythonFiles(res.files || []);
          setStatus(
            "python",
            persist
              ? "Persistence ENABLED (files stored in AppData sandbox)"
              : "Persistence DISABLED (temporary sandbox auto-cleaned)",
            "info"
          );
        }
      }
    } catch (err) {
      console.error("Failed to toggle persistence:", err);
    }
  });

  // Run button
  runBtn.addEventListener("click", pythonRunCode);

  // Explorer buttons (multiple entry points for the same action)
  document.getElementById("python-btn-open-explorer").addEventListener("click", pythonOpenExplorer);
  document.getElementById("python-btn-open-explorer-mini").addEventListener("click", pythonOpenExplorer);

  // Clear sandbox
  document.getElementById("python-btn-clear-sandbox").addEventListener("click", pythonClearWorkspace);

  // Copy context buttons (multiple entry points)
  document.getElementById("python-btn-copy-context").addEventListener("click", pythonCopyFolderContext);
  document.getElementById("python-btn-copy-context-mini").addEventListener("click", pythonCopyFolderContext);
  document.getElementById("python-btn-copy-context-full").addEventListener("click", pythonCopyFolderContext);

  // Query initial sandbox status
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

/**
 * Reflect persistence mode in the metric label and folder path label.
 * @param {boolean} persist - Whether persistence is enabled.
 * @param {string} folderPath - The current sandbox or persistent folder path.
 */
export function updatePythonPersistUI(persist, folderPath) {
  const persistMetric = document.getElementById("python-metric-persist");
  const folderLabel   = document.getElementById("python-folder-label");

  if (persist) {
    persistMetric.innerText   = "Persistent (Saved)";
    persistMetric.style.color = "var(--brand-excel)";
    folderLabel.innerText     = `Persistent: ${folderPath}`;
    folderLabel.title         = folderPath;
  } else {
    persistMetric.innerText   = "Temp (Auto-Clean)";
    persistMetric.style.color = "var(--text-secondary)";
    folderLabel.innerText     = `Sandbox: ${folderPath}`;
    folderLabel.title         = folderPath;
  }
}

/**
 * Execute the Python script in the editor inside the sandbox.
 */
export async function pythonRunCode() {
  const editor      = document.getElementById("python-editor");
  const code        = editor.value;
  const consoleElem = document.getElementById("python-log-console");
  const persist     = document.getElementById("python-persist-checkbox").checked;

  if (!code.trim()) {
    setStatus("python", "Code editor is empty. Paste or write a Python script to run.", "warning");
    return;
  }

  setStatus("python", "Executing Python script in sandbox…", "info", true);
  setButtonsDisabled("view-python", true);
  consoleElem.innerHTML = `<span class="info-line"># Running script (sys.executable)...</span>\n`;

  try {
    const res = await window.pywebview.api.python_run_code(code, persist);

    let outHtml = "";
    if (res.stdout) outHtml += `<span class="stdout-line">${escapeHtml(res.stdout)}</span>\n`;
    if (res.stderr)  outHtml += `<span class="stderr-line">${escapeHtml(res.stderr)}</span>\n`;
    if (!res.stdout && !res.stderr) {
      outHtml += `<span class="info-line"># Script executed with no terminal output.</span>\n`;
    }
    outHtml += `\n<span class="info-line"># Process finished with exit code ${res.exit_code} (${res.duration_ms} ms)</span>`;
    consoleElem.innerHTML = outHtml;
    consoleElem.scrollTop = consoleElem.scrollHeight;

    // Update exit code metric
    const statusMetric = document.getElementById("python-metric-status");
    if (res.exit_code === 0) {
      statusMetric.innerText   = "✓ Success (0)";
      statusMetric.style.color = "var(--brand-excel)";
      setStatus("python", `Execution finished successfully in ${res.duration_ms} ms.`, "success");
    } else {
      statusMetric.innerText   = `✗ Exit (${res.exit_code})`;
      statusMetric.style.color = "var(--error)";
      setStatus("python", `Execution exited with error code ${res.exit_code}.`, "error");
    }

    document.getElementById("python-metric-duration").innerText = `${res.duration_ms} ms`;

    renderPythonFiles(res.all_files || []);
    refreshPythonContext();

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

/**
 * Render the sandbox file list in the Generated Files tab.
 * @param {Array<object>} files - File descriptors from the API.
 */
export function renderPythonFiles(files) {
  pythonLastFiles = files;
  const list        = document.getElementById("python-file-list");
  const countSpan   = document.getElementById("tab-files-count");
  const metricFiles = document.getElementById("python-metric-files");

  countSpan.innerText   = files.length;
  metricFiles.innerText = `${files.length} File${files.length === 1 ? "" : "s"}`;
  list.innerHTML        = "";

  if (!files || files.length === 0) {
    list.innerHTML = `<div class="empty-state">No generated files in the workspace. Any documents created by your script will appear here.</div>`;
    return;
  }

  files.forEach(f => {
    const item = document.createElement("div");
    item.className = "python-file-item";

    // Determine icon by file type
    let iconName  = "file";
    let iconClass = "";
    const ext = (f.ext || "").toLowerCase();

    if (f.is_dir) {
      iconName = "folder"; iconClass = "folder";
    } else if (["py", "json", "js", "html", "css", "sql"].includes(ext)) {
      iconName = "file-code"; iconClass = "code";
    } else if (["doc", "docx", "txt", "md", "pdf"].includes(ext)) {
      iconName = "file-text"; iconClass = "doc";
    } else if (["xls", "xlsx", "csv"].includes(ext)) {
      iconName = "table"; iconClass = "sheet";
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

    item.querySelector('[data-action="open"]').addEventListener("click", () => pythonOpenFile(f.name));
    item.querySelector('[data-action="copy-name"]').addEventListener("click", () => {
      navigator.clipboard.writeText(f.name).then(() => {
        setStatus("python", `Copied '${f.name}' to clipboard`, "info");
      });
    });

    list.appendChild(item);
  });

  if (window.lucide) lucide.createIcons();
}

/**
 * Refresh the folder context textarea from the API.
 */
export async function refreshPythonContext() {
  try {
    if (window.pywebview?.api) {
      const res = await window.pywebview.api.python_get_folder_context();
      if (res.success) {
        document.getElementById("python-context-text").value = res.context_markdown;
      }
    }
  } catch (e) {}
}

/**
 * Open the sandbox directory in Windows Explorer.
 */
export async function pythonOpenExplorer() {
  try {
    if (window.pywebview?.api) {
      const res = await window.pywebview.api.python_open_folder();
      setStatus(
        "python",
        res.success ? "Opened sandbox directory in Windows Explorer." : `Could not open folder: ${res.error}`,
        res.success ? "success" : "error"
      );
    }
  } catch (err) {
    setStatus("python", `Explorer error: ${err}`, "error");
  }
}

/**
 * Open a specific file from the sandbox in its default application.
 * @param {string} filename - Filename relative to the sandbox root.
 */
export async function pythonOpenFile(filename) {
  try {
    if (window.pywebview?.api) {
      const res = await window.pywebview.api.python_open_file(filename);
      setStatus("python", res.success ? `Opened ${filename}` : res.error, res.success ? "success" : "error");
    }
  } catch (err) {
    setStatus("python", `Open file error: ${err}`, "error");
  }
}

/**
 * Generate and copy the folder context markdown to clipboard.
 */
export async function pythonCopyFolderContext() {
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

/**
 * Clear all files in the current Python sandbox after confirmation.
 */
export async function pythonClearWorkspace() {
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
