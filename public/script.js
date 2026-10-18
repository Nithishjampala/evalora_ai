
const form = document.getElementById("evaluationForm");
const submitButton = document.getElementById("evaluateButton");
const loadingPanel = document.getElementById("loadingPanel");
const resultsPanel = document.getElementById("resultsPanel");
const errorPanel = document.getElementById("errorPanel");

const ratingPoints = {
  "Outstanding": 100,
  "Very Good": 85,
  "Good": 70,
  "Needs Improvement": 40,
  "Incomplete": 0
};

const el = (id) => document.getElementById(id);

el("submission_file").addEventListener("change", (event) => {
  el("submissionFileName").textContent =
    event.target.files[0]?.name || "PDF, DOCX, PPTX, TXT or RTF · Max 4 MB";
});

el("guideline_file").addEventListener("change", (event) => {
  el("guidelineFileName").textContent =
    event.target.files[0]?.name || "PDF, DOCX, PPTX, TXT or RTF · Max 4 MB";
});

function escapeHTML(value) {
  return String(value ?? "").replace(/[&<>"']/g, (char) => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#39;"
  })[char]);
}

function showError(message) {
  errorPanel.textContent = message;
  errorPanel.classList.remove("hidden");
}

function ratingScore(rating) {
  return ratingPoints[rating] ?? 0;
}

function renderResults(data) {
  const result = data.result || data;
  const questions = Array.isArray(result.questions) ? result.questions : [];
  const compliance = Array.isArray(result.guideline_compliance)
    ? result.guideline_compliance : [];

  el("resultProjectName").textContent =
    result.project_name || form.elements.project_name.value;
  el("resultMeta").textContent =
    `${result.phase || form.elements.phase.value} phase · Project ID: ${result.project_id || form.elements.project_id.value}`;

  const calculated = questions
    .filter((q) => /^Q[1-5]$/i.test(String(q.id || "")))
    .map((q) => Number.isFinite(Number(q.score)) ? Number(q.score) : ratingScore(q.rating));

  const average = calculated.length
    ? Math.round(calculated.reduce((sum, score) => sum + score, 0) / calculated.length)
    : Number(result.overall_score || 0);

  el("overallScore").textContent = Math.max(0, Math.min(100, average));
  el("scoreLabel").textContent = result.overall_feedback || "Evaluation completed.";
  const confidence = Number(result.ai_confidence_score ?? 0);
  el("confidenceScore").textContent =
    `${Math.max(0, Math.min(100, Math.round(confidence)))}%`;

  el("questionResults").innerHTML = questions.map((q) => `
    <div class="question-card">
      <div class="question-top">
        <strong>${escapeHTML(q.id || "Question")} · ${escapeHTML(q.question || "")}</strong>
        <span class="rating-pill">${escapeHTML(q.rating || "Not rated")} · ${escapeHTML(q.score ?? ratingScore(q.rating))}</span>
      </div>
      <p>${escapeHTML(q.feedback || q.evidence || "No explanation provided.")}</p>
    </div>
  `).join("") || "<p class='muted'>No question results were returned.</p>";

  el("complianceResults").innerHTML = compliance.map((item) => `
    <div class="compliance-card">
      <h4>${escapeHTML(item.requirement || "Guideline requirement")}</h4>
      <span class="compliance-status">${escapeHTML(item.status || "Not Applicable")}</span>
      <p>${escapeHTML(item.evidence || item.explanation || "")}</p>
    </div>
  `).join("") || "<p class='muted'>No guideline compliance results were returned. Upload or enter guidelines for this check.</p>";

  el("overallFeedback").textContent = result.overall_feedback || "No overall feedback provided.";
  el("additionalFeedback").textContent = result.additional_feedback || "No additional feedback provided.";

  const suggestions = Array.isArray(result.suggestions) ? result.suggestions : [];
  el("suggestionResults").innerHTML = suggestions.map((item) =>
    `<li>${escapeHTML(typeof item === "string" ? item : item.suggestion || item.text || JSON.stringify(item))}</li>`
  ).join("") || "<li>No suggestions provided.</li>";

  resultsPanel.dataset.summary = [
    `Project: ${result.project_name || ""}`,
    `Phase: ${result.phase || ""}`,
    `Overall score: ${average}/100`,
    `Feedback: ${result.overall_feedback || ""}`,
    ...suggestions.map((s) => `Suggestion: ${typeof s === "string" ? s : s.suggestion || s.text || ""}`)
  ].join("\n");

  resultsPanel.classList.remove("hidden");
  resultsPanel.scrollIntoView({ behavior: "smooth", block: "start" });
}

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  errorPanel.classList.add("hidden");
  resultsPanel.classList.add("hidden");

  const submission = el("submission_file").files[0];
  const guidelinesText = form.elements.guidelines_text.value.trim();
  const guidelineFile = el("guideline_file").files[0];

  if (!submission) {
    showError("Please select a project submission file.");
    return;
  }
  if (!guidelinesText && !guidelineFile) {
    showError("Please enter evaluation guidelines or upload a guideline file.");
    return;
  }

  const formData = new FormData(form);
  submitButton.disabled = true;
  submitButton.textContent = "Evaluating...";
  loadingPanel.classList.remove("hidden");
  loadingPanel.scrollIntoView({ behavior: "smooth", block: "center" });

  try {
    const response = await fetch("/api/evaluate", {
      method: "POST",
      body: formData
    });

    const payload = await response.json().catch(() => ({}));
    if (!response.ok) {
      throw new Error(payload.error || `Request failed (${response.status}). Check the VS Code terminal.`);
    }

    renderResults(payload);
  } catch (error) {
    showError(error.message || "An unexpected error occurred.");
  } finally {
    loadingPanel.classList.add("hidden");
    submitButton.disabled = false;
    submitButton.innerHTML = "<span>Evaluate project with AI</span><span>→</span>";
  }
});

el("copyButton").addEventListener("click", async () => {
  try {
    await navigator.clipboard.writeText(resultsPanel.dataset.summary || "");
    el("copyButton").textContent = "Copied!";
    setTimeout(() => { el("copyButton").textContent = "Copy summary"; }, 1600);
  } catch {
    showError("Clipboard access failed. Select and copy the result text manually.");
  }
});

el("newEvaluationButton").addEventListener("click", () => {
  resultsPanel.classList.add("hidden");
  form.reset();
  el("submissionFileName").textContent = "PDF, DOCX, PPTX, TXT or RTF · Max 4 MB";
  el("guidelineFileName").textContent = "PDF, DOCX, PPTX, TXT or RTF · Max 4 MB";
  window.scrollTo({ top: 0, behavior: "smooth" });
});
