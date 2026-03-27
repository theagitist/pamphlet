const PHASES = ['Parsing', 'Extraction', 'Formatting', 'Writing'];
const POLL_INTERVAL = 500;

// DOM elements
const dropZone = document.getElementById('drop-zone');
const fileInput = document.getElementById('file-input');
const fileInfo = document.getElementById('file-info');
const fileName = document.getElementById('file-name');
const clearFileBtn = document.getElementById('clear-file');
const optionsSection = document.getElementById('options-section');
const generateSection = document.getElementById('generate-section');
const generateBtn = document.getElementById('generate-btn');
const progressSection = document.getElementById('progress-section');
const progressPhase = document.getElementById('progress-phase');
const progressStep = document.getElementById('progress-step');
const progressFill = document.getElementById('progress-fill');
const resultSection = document.getElementById('result-section');
const resultSuccess = document.getElementById('result-success');
const resultError = document.getElementById('result-error');
const resultMeta = document.getElementById('result-meta');
const downloadLink = document.getElementById('download-link');
const errorMessage = document.getElementById('error-message');
const unsupportedNotice = document.getElementById('unsupported-notice');
const startOverBtn = document.getElementById('start-over');
const optPageNumbers = document.getElementById('opt-page-numbers');

let selectedFile = null;
let sessionId = null;
let pollTimer = null;

// ── File selection ─────────────────────────────────────────────

dropZone.addEventListener('click', () => fileInput.click());

dropZone.addEventListener('dragover', (e) => {
  e.preventDefault();
  dropZone.classList.add('drag-over');
});

dropZone.addEventListener('dragleave', () => {
  dropZone.classList.remove('drag-over');
});

dropZone.addEventListener('drop', (e) => {
  e.preventDefault();
  dropZone.classList.remove('drag-over');
  const file = e.dataTransfer.files[0];
  if (file) selectFile(file);
});

fileInput.addEventListener('change', () => {
  if (fileInput.files[0]) selectFile(fileInput.files[0]);
});

clearFileBtn.addEventListener('click', () => {
  clearFile();
});

function selectFile(file) {
  if (!file.name.toLowerCase().endsWith('.pptx')) {
    alert('Please select a .pptx file.');
    return;
  }
  if (file.size > 50 * 1024 * 1024) {
    alert('File is too large. Maximum size is 50 MB.');
    return;
  }
  selectedFile = file;
  fileName.textContent = file.name;
  fileInfo.hidden = false;
  dropZone.hidden = true;
  optionsSection.hidden = false;
  generateSection.hidden = false;
}

function clearFile() {
  selectedFile = null;
  fileInput.value = '';
  fileInfo.hidden = true;
  dropZone.hidden = false;
  optionsSection.hidden = true;
  generateSection.hidden = true;
}

// ── Generate ───────────────────────────────────────────────────

generateBtn.addEventListener('click', async () => {
  if (!selectedFile) return;

  generateBtn.disabled = true;
  generateBtn.textContent = 'Uploading...';

  try {
    // Upload
    const formData = new FormData();
    formData.append('file', selectedFile);

    const uploadRes = await fetch('/api/upload', { method: 'POST', body: formData });
    if (!uploadRes.ok) {
      const err = await uploadRes.json();
      throw new Error(err.error || 'Upload failed');
    }

    const { id } = await uploadRes.json();
    sessionId = id;

    // Start conversion
    const genRes = await fetch(`/api/generate/${id}`, { method: 'POST' });
    const genData = await genRes.json();
    if (!genRes.ok) {
      throw new Error(genData.error || 'Generation failed');
    }

    // Show progress (may start as 'queued')
    showProgress();
    if (genData.status === 'queued' && genData.queuePosition > 0) {
      progressPhase.textContent = 'Waiting for space to process...';
      progressStep.textContent = `Position ${genData.queuePosition}`;
      progressFill.style.width = '5%';
    }
    startPolling();

  } catch (err) {
    showError(err.message);
  }
});

// ── Progress polling ───────────────────────────────────────────

function showProgress() {
  generateSection.hidden = true;
  optionsSection.hidden = true;
  document.getElementById('upload-section').hidden = true;
  progressSection.hidden = false;
  resultSection.hidden = true;
  updateProgressUI('Preparing...', 0, PHASES.length);
}

function updateProgressUI(phase, phaseIndex, total) {
  progressPhase.textContent = phase;
  progressStep.textContent = `${phaseIndex + 1} / ${total}`;
  const pct = ((phaseIndex + 1) / total) * 100;
  progressFill.style.width = pct + '%';
}

function startPolling() {
  if (pollTimer) clearInterval(pollTimer);
  pollTimer = setInterval(pollStatus, POLL_INTERVAL);
}

function stopPolling() {
  if (pollTimer) {
    clearInterval(pollTimer);
    pollTimer = null;
  }
}

async function pollStatus() {
  if (!sessionId) return;

  try {
    const res = await fetch(`/api/status/${sessionId}`);
    if (!res.ok) {
      stopPolling();
      showError('Lost connection to server.');
      return;
    }

    const data = await res.json();

    if (data.status === 'queued') {
      progressPhase.textContent = 'Waiting for space to process...';
      progressStep.textContent = data.queue ? `${data.queue.waiting} ahead` : '';
      progressFill.style.width = '5%';
    } else if (data.status === 'processing') {
      updateProgressUI(data.phase || 'Processing...', data.phaseIndex || 0, data.totalPhases || 4);
    } else if (data.status === 'ready') {
      stopPolling();
      progressFill.style.width = '100%';
      progressPhase.textContent = 'Complete';
      setTimeout(() => showResult(data), 400);
    } else if (data.status === 'failed') {
      stopPolling();
      showError(data.error || 'Conversion failed.');
    }
  } catch {
    stopPolling();
    showError('Lost connection to server.');
  }
}

// ── Results ────────────────────────────────────────────────────

function showResult(data) {
  progressSection.hidden = true;
  resultSection.hidden = false;
  resultSuccess.hidden = false;
  resultError.hidden = true;

  resultMeta.textContent = `${data.slideCount} slides converted`;
  downloadLink.href = `/api/download/${sessionId}`;

  // Unsupported objects notice
  if (data.unsupportedObjects && data.unsupportedObjects.length > 0) {
    unsupportedNotice.hidden = false;
    const items = data.unsupportedObjects
      .map(o => `Slide ${o.slide}: ${o.description}`)
      .join('; ');
    unsupportedNotice.textContent = `Some content could not be converted: ${items}`;
  } else {
    unsupportedNotice.hidden = true;
  }
}

function showError(message) {
  progressSection.hidden = true;
  generateSection.hidden = true;
  resultSection.hidden = false;
  resultSuccess.hidden = true;
  resultError.hidden = false;
  errorMessage.textContent = message;
}

// ── Start over ─────────────────────────────────────────────────

startOverBtn.addEventListener('click', () => {
  stopPolling();
  sessionId = null;
  selectedFile = null;
  fileInput.value = '';

  // Reset UI
  document.getElementById('upload-section').hidden = false;
  dropZone.hidden = false;
  fileInfo.hidden = true;
  optionsSection.hidden = true;
  generateSection.hidden = true;
  progressSection.hidden = true;
  resultSection.hidden = true;
  generateBtn.disabled = false;
  generateBtn.textContent = 'Generate Handout';
  progressFill.style.width = '0%';
});
