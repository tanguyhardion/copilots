/**
 * Modal controllers — System Prompt, Cheatsheet, and Help modals.
 */

import { ROUTES } from "../data/routes.js";
import { setStatus } from "../core/utils.js";
import { currentRoute } from "../core/router.js";

/* =========================================================================
   Setup
   ========================================================================= */

/**
 * Wire up all modal close buttons, backdrop click, and keyboard handlers.
 * Must be called once on app init.
 */
export function setupModals() {
  // Prompt modal
  document.getElementById("modal-prompt-close").addEventListener("click", closePromptModal);
  document.getElementById("modal-prompt-btn-cancel").addEventListener("click", closePromptModal);
  document.getElementById("prompt-modal").addEventListener("click", e => {
    if (e.target === document.getElementById("prompt-modal")) closePromptModal();
  });

  // Copy prompt to clipboard
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
      const res = await window.pywebview.api.save_prompt_override(_activeModalPromptKey, text);
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

  // Reset prompt to bundled default
  document.getElementById("modal-prompt-btn-reset").addEventListener("click", async () => {
    if (!confirm("Are you sure you want to reset this system prompt to the bundled default?")) return;
    try {
      const res = await window.pywebview.api.reset_prompt_default(_activeModalPromptKey);
      if (res.success) {
        document.getElementById("modal-prompt-editor").value = res.content;
        document.getElementById("modal-prompt-badge").innerText = "Bundled Default";
        setStatus(currentRoute, res.message, "success");
      }
    } catch (e) {
      alert("Reset error: " + e);
    }
  });

  // Cheatsheet modal
  document.getElementById("cheatsheet-close").addEventListener("click", closeCheatsheetModal);
  document.getElementById("cheatsheet-btn-ok").addEventListener("click", closeCheatsheetModal);
  document.getElementById("cheatsheet-modal").addEventListener("click", e => {
    if (e.target === document.getElementById("cheatsheet-modal")) closeCheatsheetModal();
  });

  // Help modal
  document.getElementById("help-close").addEventListener("click", closeHelpModal);
  document.getElementById("help-btn-ok").addEventListener("click", closeHelpModal);
  document.getElementById("help-modal").addEventListener("click", e => {
    if (e.target === document.getElementById("help-modal")) closeHelpModal();
  });

  // Escape key closes any open modal
  document.addEventListener("keydown", e => {
    if (e.key === "Escape") {
      closePromptModal();
      closeCheatsheetModal();
      closeHelpModal();
    }
  });
}

/* =========================================================================
   Prompt Modal
   ========================================================================= */

/** Tracks which copilot's prompt is currently being viewed/edited. */
let _activeModalPromptKey = "powerpoint";

/**
 * Open the system prompt modal for a given copilot.
 * @param {string} copilotKey - Route key identifying the copilot.
 */
export async function openPromptModal(copilotKey) {
  _activeModalPromptKey = copilotKey;
  const modal  = document.getElementById("prompt-modal");
  const editor = document.getElementById("modal-prompt-editor");
  const title  = document.getElementById("modal-prompt-title");
  const badge  = document.getElementById("modal-prompt-badge");

  // Tint modal primary buttons with this copilot's brand color
  const accentColor = ROUTES[copilotKey]?.badgeColor || "var(--primary)";
  modal.style.setProperty("--modal-accent", accentColor);
  modal.setAttribute("data-copilot", copilotKey);

  try {
    const info = await window.pywebview.api.get_prompt_info(copilotKey);
    if (info.success) {
      title.innerText  = info.title;
      editor.value     = info.content;
      badge.innerText  = info.is_custom ? "User Custom Override" : "Bundled Default";
      modal.classList.add("open");
    }
  } catch (err) {
    console.error("Failed to load prompt:", err);
  }
}

/** Close the system prompt modal. */
export function closePromptModal() {
  document.getElementById("prompt-modal").classList.remove("open");
}

/* =========================================================================
   Cheatsheet Modal
   ========================================================================= */

/**
 * Open the DSL cheatsheet modal for a given copilot.
 * @param {string} routeId - Route key.
 */
export function openCheatsheetModal(routeId) {
  const modal   = document.getElementById("cheatsheet-modal");
  const title   = document.getElementById("cheatsheet-title");
  const content = document.getElementById("cheatsheet-content");

  // Tint modal primary buttons with this copilot's brand color
  const accentColor = ROUTES[routeId]?.badgeColor || "var(--primary)";
  modal.style.setProperty("--modal-accent", accentColor);
  modal.setAttribute("data-copilot", routeId);

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

/** Close the cheatsheet modal. */
export function closeCheatsheetModal() {
  document.getElementById("cheatsheet-modal").classList.remove("open");
}

/* =========================================================================
   Help Modal
   ========================================================================= */

/**
 * Open the help guide modal for a given copilot.
 * @param {string} routeId - Route key.
 */
export function openHelpModal(routeId) {
  const modal   = document.getElementById("help-modal");
  const title   = document.getElementById("help-title");
  const badge   = document.getElementById("help-badge");
  const icon    = document.getElementById("help-modal-icon");
  const content = document.getElementById("help-content");

  const routeMeta = ROUTES[routeId] || { title: "Copilot" };
  title.innerText = `How to Use ${routeMeta.title}`;
  badge.innerText = `${routeMeta.badge || "Copilot"} · Workflow & Features`;
  if (icon) icon.style.color = routeMeta.badgeColor || "var(--primary)";

  // Tint modal primary buttons with this copilot's brand color
  modal.style.setProperty("--modal-accent", routeMeta.badgeColor || "var(--primary)");
  modal.setAttribute("data-copilot", routeId);

  content.innerHTML = HELP_CONTENT[routeId] || "";

  modal.classList.add("open");
  if (window.lucide) lucide.createIcons();
}

/** Close the help modal. */
export function closeHelpModal() {
  document.getElementById("help-modal").classList.remove("open");
}

/* =========================================================================
   Help Content — static HTML strings keyed by routeId
   ========================================================================= */

const HELP_CONTENT = {
  powerpoint: `
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
  `,

  word: `
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
  `,

  excel: `
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
  `,

  cv: `
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
          <p>Choose the output language, sections to include, and destination folder, then generate either a <strong>1-Slide Executive PowerPoint (.pptx)</strong> proposal or a comprehensive multi-page <strong>Europass Word (.docx)</strong> document.</p>
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
          <p class="help-feature-desc">Populates <code>template.pptx</code> in-place with candidate data, preserving exact geometry, coordinates, and typography, while honoring the selected language, included sections, and output folder.</p>
        </div>
        <div class="help-feature-item">
          <span class="help-feature-btn-badge"><i data-lucide="file-check"></i> Generate Europass Word (.docx)</span>
          <p class="help-feature-desc">Generates a Microsoft Word document matching the standard Europass layout with clean margins, timeline tables, and skills matrix, using the selected language, included sections, and output folder.</p>
        </div>
      </div>
    </div>
  `,

  python: `
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
  `,

  organizer: `
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
  `,
};
