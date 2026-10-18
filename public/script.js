
const RENDER_URL = "https://evalora-ai-7jr2.onrender.com";

async function uploadFile(file) {
  const formData = new FormData();
  formData.append("files", file);

  const response = await fetch(
    `${RENDER_URL}/gradio_api/upload`,
    {
      method: "POST",
      body: formData
    }
  );

  if (!response.ok) {
    throw new Error(`File upload failed: HTTP ${response.status}`);
  }

  const paths = await response.json();
  return paths[0];
}

async function runEvaluation({
  name,
  projectId,
  projectName,
  phase,
  guidelineFile,
  submissionFile
}) {
  const guidelinePath = await uploadFile(guidelineFile);
  const submissionPath = await uploadFile(submissionFile);

  const response = await fetch(
    `${RENDER_URL}/gradio_api/call/v2/evaluate`,
    {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        data: [
          name,
          projectId,
          projectName,
          phase,
          "",
          {
            path: guidelinePath,
            meta: { _type: "gradio.FileData" }
          },
          {
            path: submissionPath,
            meta: { _type: "gradio.FileData" }
          }
        ]
      })
    }
  );

  if (!response.ok) {
    throw new Error(`Evaluation request failed: HTTP ${response.status}`);
  }

  const { event_id } = await response.json();

  const resultResponse = await fetch(
    `${RENDER_URL}/gradio_api/call/evaluate/${event_id}`
  );

  if (!resultResponse.ok) {
    throw new Error(`Result retrieval failed: HTTP ${resultResponse.status}`);
  }

  const stream = resultResponse.body.getReader();
  const decoder = new TextDecoder();
  let buffer = "";

  while (true) {
    const { value, done } = await stream.read();
    if (done) break;

    buffer += decoder.decode(value, { stream: true });

    const events = buffer.split("\n\n");
    buffer = events.pop() || "";

    for (const event of events) {
      const eventName = event.match(/^event:\s*(.+)$/m)?.[1];
      const dataText = event.match(/^data:\s*(.+)$/m)?.[1];

      if (!dataText) continue;

      if (eventName === "complete") {
        return JSON.parse(dataText);
      }

      if (eventName === "error") {
        throw new Error(dataText);
      }
    }
  }

  throw new Error("Evaluation ended without a complete result.");
}
