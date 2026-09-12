/**
 * CV Copilot view controller.
 */

import { setStatus, setButtonsDisabled } from "../core/utils.js";

/**
 * Bind all action buttons in the CV view panel.
 */
export async function setupCVView() {
  const cvEditor          = document.getElementById("cv-editor");
  const cvLanguageSelect  = document.getElementById("cv-language-select");
  const cvOutputPathInput = document.getElementById("cv-output-path");
  const cvSelectOutputBtn = document.getElementById("cv-btn-select-output");
  const cvOptionsPanel    = document.getElementById("cv-options-panel");
  const cvOptionsToggle   = document.getElementById("cv-options-toggle");

  // Collapsible Generation Options panel
  if (cvOptionsToggle && cvOptionsPanel) {
    cvOptionsToggle.addEventListener("click", () => {
      const isCollapsed = cvOptionsPanel.classList.toggle("collapsed");
      cvOptionsToggle.setAttribute("aria-expanded", String(!isCollapsed));
      const hint = cvOptionsToggle.querySelector(".cv-options-hint");
      if (hint) {
        hint.textContent = isCollapsed ? "Click to expand" : "Click to collapse";
      }
    });
  }

  /**
   * Collect the section checkboxes for a given output format.
   * @param {"docx"|"pptx"} format
   * @returns {Record<string, boolean>}
   */
  const collectCVSections = (format) => {
    const sections = {};
    document.querySelectorAll(`#view-cv input[data-cv-format="${format}"]`).forEach(checkbox => {
      sections[checkbox.dataset.section] = checkbox.checked;
    });
    return sections;
  };

  // Output folder picker
  if (cvSelectOutputBtn) {
    cvSelectOutputBtn.addEventListener("click", async () => {
      try {
        const res = await window.pywebview.api.cv_select_output_folder();
        if (res.success) {
          cvOutputPathInput.value = res.path || "";
          cvOutputPathInput.title = res.path || "";
          setStatus("cv", "Output folder selected.", "success");
        } else if (!res.cancelled) {
          setStatus("cv", res.error || "Could not select output folder.", "error");
        }
      } catch (err) {
        setStatus("cv", `Folder selection error: ${err}`, "error");
      }
    });
  }

  // Format JSON and refresh metrics
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

  // Auto-audit on paste
  cvEditor.addEventListener("paste", async () => {
    await new Promise(r => setTimeout(r, 0));
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

  // Generate Europass Word (.docx)
  document.getElementById("cv-btn-generate").addEventListener("click", async () => {
    setStatus("cv", "Generating Europass Word (.docx) document…", "info", true);
    setButtonsDisabled("view-cv", true);
    try {
      const res = await window.pywebview.api.cv_generate_docx(
        cvEditor.value,
        cvLanguageSelect.value,
        collectCVSections("docx"),
        cvOutputPathInput.value.trim()
      );
      setStatus("cv", res.success ? res.message : res.error, res.success ? "success" : "error");
    } catch (err) {
      setStatus("cv", `Word generation error: ${err}`, "error");
    } finally {
      setButtonsDisabled("view-cv", false);
    }
  });

  // Generate 1-Slide Executive PowerPoint (.pptx)
  const btnPptx = document.getElementById("cv-btn-generate-pptx");
  if (btnPptx) {
    btnPptx.addEventListener("click", async () => {
      setStatus("cv", "Generating 1-Slide Executive PowerPoint (.pptx) document…", "info", true);
      setButtonsDisabled("view-cv", true);
      try {
        const res = await window.pywebview.api.cv_generate_pptx(
          cvEditor.value,
          cvLanguageSelect.value,
          collectCVSections("pptx"),
          cvOutputPathInput.value.trim()
        );
        setStatus("cv", res.success ? res.message : res.error, res.success ? "success" : "error");
      } catch (err) {
        setStatus("cv", `PowerPoint generation error: ${err}`, "error");
      } finally {
        setButtonsDisabled("view-cv", false);
      }
    });
  }
}

/**
 * Populate the top metric cards from parsed CV JSON.
 * @param {object} cvData - Parsed CV JSON object.
 */
export function updateCVMetrics(cvData) {
  const p    = cvData.personal_info || cvData.personal_information || {};
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

  document.getElementById("cv-metric-name").innerText   = name;
  document.getElementById("cv-metric-exp").innerText    = `${expCount} Positions`;
  document.getElementById("cv-metric-skills").innerText = `${skillsCount} Skills`;
}

/**
 * Render DQ audit results into the audit-results container.
 * @param {object} audit - Audit result from the Python DQ engine.
 */
export function renderCVAuditResults(audit) {
  document.getElementById("cv-metric-score").innerText  = `${audit.dq_score}%`;
  document.getElementById("cv-metric-name").innerText   = audit.candidate_name;
  document.getElementById("cv-metric-exp").innerText    = `${audit.experience_count} Positions`;
  document.getElementById("cv-metric-skills").innerText = `${audit.skills_count} Skills`;

  const container = document.getElementById("cv-audit-results");
  container.innerHTML = "";

  // Summary badge
  const summaryBadge = document.createElement("div");
  summaryBadge.className = `audit-summary-badge ${audit.passed ? "passed" : "warning"}`;
  summaryBadge.innerHTML = `
    <span>${audit.passed ? "✓ PASSED" : "⚠ ATTENTION REQUIRED"}</span>
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
