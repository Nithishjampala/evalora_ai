'use strict';

const form = document.querySelector('#form');
const button = document.querySelector('#submit');
const status = document.querySelector('#status');
const results = document.querySelector('#results');

const descriptions = {
  Ideation: [
    'Ideation review',
    'Problem clarity, user need, research, feasibility and originality.'
  ],
  Design: [
    'Design review',
    'Solution design, problem fit, core features and technical feasibility.'
  ],
  Build: [
    'Build review',
    'Implementation quality, problem fit, working features and feasibility.'
  ],
  Demo: [
    'Demo review',
    'Final solution impact, feature demonstration, validation and sustainability.'
  ]
};

const phaseInput = document.querySelector('[name="phase"]');

phaseInput?.addEventListener('change', (event) => {
  const description = descriptions[event.target.value];

  if (!description) return;

  document.querySelector('#phaseTitle').textContent = description[0];
  document.querySelector('#phaseDesc').textContent = description[1];
});

[
  ['#guideline', '#guidelineName'],
  ['#submission', '#submissionName']
].forEach(([inputSelector, labelSelector]) => {
  const input = document.querySelector(inputSelector);
  const label = document.querySelector(labelSelector);

  input?.addEventListener('change', (event) => {
    const file = event.target.files[0];

    if (label) {
      label.textContent = file
        ? `${file.name} · ${(file.size / 1048576).toFixed(2)} MB`
        : 'Choose a file';
    }
  });
});

form.addEventListener('submit', async (event) => {
  event.preventDefault();

  button.disabled = true;
  button.querySelector('span').textContent = 'Evaluating project…';
  status.textContent =
    'Extracting document text and running the Gemini evaluation…';

  results.hidden = true;

  try {
    const data = new FormData(form);

    // Validate the uploaded documents.
    for (const key of ['guideline_file', 'submission_file']) {
      const file = data.get(key);

      if (!(file instanceof File) || !file.name) {
        throw new Error(
          'Please upload both the guidelines and student submission.'
        );
      }

      if (file.size > 8 * 1024 * 1024) {
        throw new Error(`${file.name} exceeds the 8 MB limit.`);
      }
    }

    const response = await fetch('/api/evaluate', {
      method: 'POST',
      headers: {
        Accept: 'application/json'
      },
      body: data
    });

    // Read as text first to handle server errors that are not JSON.
    const raw = await response.text();

    const requestId =
      response.headers.get('x-vercel-id') ||
      response.headers.get('x-vercel-error') ||
      '';

    let output;

    try {
      output = JSON.parse(raw);
    } catch (_parseError) {
      const preview = (raw || '')
        .replace(/\s+/g, ' ')
        .trim()
        .slice(0, 1000);

      const details =
        preview || 'The server returned an empty response.';

      const trace = requestId
        ? ` Vercel request ID: ${requestId}.`
        : '';

      throw new Error(
        `Server returned a non-JSON response (HTTP ${response.status}). ` +
        `${details}${trace}`
      );
    }

    // Show the actual error returned by the backend.
    if (!response.ok) {
      const detail =
        output.detail ||
        output.message ||
        output.error;

      const trace = requestId
        ? ` (Vercel request ID: ${requestId})`
        : '';

      const message =
        typeof detail === 'string'
          ? detail
          : `Request failed with HTTP ${response.status}.`;

      throw new Error(`${message}${trace}`);
    }

    if (output.success !== true) {
      throw new Error(
        output.message ||
        'The server did not confirm a successful evaluation.'
      );
    }

    // Display the evaluation results.
    document.querySelector('#score').textContent =
      `${output.overall_score}/100`;

    document.querySelector('#confidence').textContent =
      `${output.confidence}/100`;

    document.querySelector('#model').textContent =
      output.model_used || 'Not reported';

    document.querySelector('#report').textContent =
      output.report || 'No report returned.';

    document.querySelector('#saveStatus').textContent =
      output.status || '';

    results.hidden = false;

    status.textContent = 'Evaluation completed successfully.';

    results.scrollIntoView({
      behavior: 'smooth',
      block: 'start'
    });

  } catch (error) {
    const message =
      error && error.message
        ? error.message
        : String(error);

    status.textContent = `Evaluation failed: ${message}`;

    console.error('Evalora evaluation request failed:', error);

  } finally {
    button.disabled = false;
    button.querySelector('span').textContent = 'Evaluate project';
  }
});

// Copy the evaluation report.
document.querySelector('#copy').addEventListener('click', async () => {
  try {
    await navigator.clipboard.writeText(
      document.querySelector('#report').textContent
    );

    document.querySelector('#copy').textContent = 'Copied!';

    setTimeout(() => {
      document.querySelector('#copy').textContent = 'Copy report';
    }, 1500);

  } catch (_error) {
    document.querySelector('#copy').textContent = 'Copy unavailable';
  }
});
