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
          <span class="help-feature-btn-badge"><i data-lucide="clipboard"></i> Copy Shapes to Clipboard</span>
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
};
