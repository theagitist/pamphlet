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
    showError('Please select a valid .pptx PowerPoint file.');
    document.getElementById('upload-section').hidden = true;
    return;
  }
  if (file.size > 50 * 1024 * 1024) {
    showError('File is too large (max 50 MB). Please choose a smaller file.');
    document.getElementById('upload-section').hidden = true;
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
  generateBtn.textContent = 'Preparing...';

  try {
    // Show progress UI immediately for Upload phase
    showProgress();
    updateProgressUI('Uploading...', -1, PHASES.length); // -1 to indicate pre-processing
    progressFill.style.width = '2%';

    // Upload with progress
    const sessionIdResult = await uploadWithProgress(selectedFile);
    sessionId = sessionIdResult;

    // Start conversion
    updateProgressUI('Starting conversion...', 0, PHASES.length);
    const genRes = await fetch(`/api/generate/${sessionId}`, { method: 'POST' });
    const genData = await genRes.json();
    if (!genRes.ok) {
      throw new Error(genData.error || 'Generation failed');
    }

    // Handle queued state
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

function uploadWithProgress(file) {
  return new Promise((resolve, reject) => {
    const xhr = new XMLHttpRequest();
    const formData = new FormData();
    formData.append('file', file);

    xhr.upload.addEventListener('progress', (e) => {
      if (e.lengthComputable) {
        const percent = Math.round((e.loaded / e.total) * 90); // Cap at 90% until server responds
        progressFill.style.width = percent + '%';
        progressStep.textContent = `${Math.round(e.loaded / 1024 / 1024)}MB / ${Math.round(e.total / 1024 / 1024)}MB`;
      }
    });

    xhr.addEventListener('load', () => {
      if (xhr.status >= 200 && xhr.status < 300) {
        try {
          const data = JSON.parse(xhr.responseText);
          resolve(data.id);
        } catch (e) {
          reject(new Error('Invalid server response'));
        }
      } else {
        try {
          const err = JSON.parse(xhr.responseText);
          reject(new Error(err.error || 'Upload failed'));
        } catch (e) {
          reject(new Error(`Upload failed with status ${xhr.status}`));
        }
      }
    });

    xhr.addEventListener('error', () => reject(new Error('Network error during upload')));
    xhr.addEventListener('abort', () => reject(new Error('Upload aborted')));

    xhr.open('POST', '/api/upload');
    xhr.send(formData);
  });
}

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
  if (phaseIndex === -1) {
    // Already set by upload event listener, but fallback just in case
    if (!progressStep.textContent.includes('MB')) {
      progressStep.textContent = 'Uploading...';
    }
    return;
  }
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

function resetUI() {
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
}

startOverBtn.addEventListener('click', resetUI);
