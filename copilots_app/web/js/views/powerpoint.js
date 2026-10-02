/**
 * PowerPoint Copilot view controller.
 */

import { setStatus, setButtonsDisabled } from "../core/utils.js";

/**
 * Bind all action buttons in the PowerPoint view panel.
 */
export async function setupPowerPointView() {
  const editor = document.getElementById("ppt-editor");

  // Read / Inspect the active slide, or every slide of the presentation
  async function readSlides(readAll) {
    setStatus(
      "ppt",
      readAll
        ? "Reading shapes and backgrounds from all PowerPoint slides…"
        : "Reading shapes and background from active PowerPoint slide…",
      "info",
      true,
    );
    setButtonsDisabled("view-powerpoint", true);
    try {
      const api = window.pywebview.api;
      const res = readAll ? await api.ppt_read_all_slides() : await api.ppt_read_active_slide();
      if (res.success) {
        editor.value = res.dsl;
        setStatus("ppt", res.message, "success");
      } else {
        setStatus("ppt", res.error, "error");
      }
    } catch (err) {
      setStatus("ppt", `Read error: ${err}`, "error");
    } finally {
      setButtonsDisabled("view-powerpoint", false);
    }
  }

  // Split button: main part reads the active slide, the menu offers active / all slides
  const readMenu = document.getElementById("ppt-read-menu");
  const readMenuToggle = document.getElementById("ppt-btn-read-menu");

  function setReadMenuOpen(open) {
    readMenu.classList.toggle("open", open);
    readMenuToggle.setAttribute("aria-expanded", String(open));
  }

  readMenuToggle.addEventListener("click", (e) => {
    e.stopPropagation();
    setReadMenuOpen(!readMenu.classList.contains("open"));
  });
  document.addEventListener("click", (e) => {
    if (!readMenu.contains(e.target)) setReadMenuOpen(false);
  });
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape") setReadMenuOpen(false);
  });

  // The main button remembers the last-used read mode; the menu only offers the other one
  const READ_MODE_KEY = "ppt-read-mode";
  const readMainBtn = document.getElementById("ppt-btn-read");
  const readActiveItem = document.getElementById("ppt-btn-read-active");
  const readAllItem = document.getElementById("ppt-btn-read-all");
  const readModes = {
    active: { readAll: false, icon: "scan-line", label: "Read Active Slide", item: readActiveItem },
    all: { readAll: true, icon: "layers", label: "Read All Slides", item: readAllItem },
  };

  let readMode = "active";
  try {
    const saved = localStorage.getItem(READ_MODE_KEY);
    if (saved in readModes) readMode = saved;
  } catch (_) {
    // localStorage unavailable — fall back to the default mode
  }

  function applyReadMode() {
    const { icon, label, item } = readModes[readMode];
    readMainBtn.title = item.title;
    readMainBtn.innerHTML = `<i data-lucide="${icon}" class="btn-icon"></i> ${label}`;
    if (window.lucide) lucide.createIcons({ nodes: [readMainBtn] });
    Object.values(readModes).forEach((m) => {
      m.item.hidden = m.item === item;
    });
  }

  function setReadMode(mode) {
    readMode = mode;
    try {
      localStorage.setItem(READ_MODE_KEY, mode);
    } catch (_) {
      // ignore — the choice just won't persist
    }
    applyReadMode();
  }

  applyReadMode();

  readMainBtn.addEventListener("click", () => readSlides(readModes[readMode].readAll));
  Object.entries(readModes).forEach(([mode, { readAll, item }]) => {
    item.addEventListener("click", () => {
      setReadMenuOpen(false);
      setReadMode(mode);
      readSlides(readAll);
    });
  });

  // Apply surgical edits to the active slide, or to the slides targeted by `edit slide=N`
  document.getElementById("ppt-btn-edit").addEventListener("click", async () => {
    const dsl = editor.value;
    if (!dsl.trim()) {
      setStatus("ppt", "Editor is empty — nothing to edit", "warning");
      return;
    }
    setStatus("ppt", "Applying edits…", "info", true);
    setButtonsDisabled("view-powerpoint", true);
    try {
      const res = await window.pywebview.api.ppt_apply_edits(dsl);
      setStatus("ppt", res.success ? res.message : res.error, res.success ? "success" : "error");
    } catch (err) {
      setStatus("ppt", `Edit error: ${err}`, "error");
    } finally {
      setButtonsDisabled("view-powerpoint", false);
    }
  });

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
