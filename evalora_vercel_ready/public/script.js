const form = document.querySelector("#form");
const button = document.querySelector("#submit");
form.addEventListener("submit", async (event) => {
  event.preventDefault();
  button.disabled = true;
  button.textContent = "Evaluating…";
  const status = document.querySelector("#status");
  const output = document.querySelector("#result");
  output.hidden = true;
  status.textContent = "Evaluating in three stages. This may take a minute; please keep this tab open.";
  try {
    const response = await fetch("/api/evaluate", { method: "POST", body: new FormData(form) });
    const data = await response.json().catch(() => ({}));
    if (!response.ok) throw new Error(data.detail || `Request failed (HTTP ${response.status})`);
    document.querySelector("#score").textContent = `Score: ${data.score}/100`;
    document.querySelector("#report").textContent = data.report || "No report returned.";
    document.querySelector("#sheet").textContent = data.sheet_status || "Google Sheets status unavailable.";
    document.querySelector("#models").textContent = "Models used: " + (data.models_used || []).join(" → ");
    output.hidden = false;
    status.textContent = "Evaluation finished.";
  } catch (error) {
    status.textContent = "Error: " + (error?.message || "Unknown error");
  } finally {
    button.disabled = false;
    button.textContent = "Evaluate project";
  }
});
