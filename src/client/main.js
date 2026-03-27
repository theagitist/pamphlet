const PHASES = ['Parsing', 'Extraction', 'Formatting', 'Writing'];
const POLL_INTERVAL = 500;

// DOM elements
const dropZone = document.getElementById('drop-zone');
const fileInput = document.getElementById('file-input');
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

let sessionId = null;
let pollTimer = null;

// ── File selection → immediate conversion ─────────────────────

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
  if (file) handleFile(file);
});

fileInput.addEventListener('change', () => {
  if (fileInput.files[0]) handleFile(fileInput.files[0]);
});

function handleFile(file) {
  if (!file.name.toLowerCase().endsWith('.pptx')) {
    showError('Please select a valid .pptx PowerPoint file.');
    return;
  }
  if (file.size > 50 * 1024 * 1024) {
    showError('File is too large (max 50 MB). Please choose a smaller file.');
    return;
  }
  startConversion(file);
}

// ── Conversion pipeline ───────────────────────────────────────

async function startConversion(file) {
  try {
    showProgress();
    updateProgressUI('Uploading...', -1, PHASES.length);
    progressFill.style.width = '2%';

    // Upload
    const id = await uploadWithProgress(file);
    sessionId = id;

    // Trigger conversion
    updateProgressUI('Starting conversion...', 0, PHASES.length);
    const genRes = await fetch(`/api/generate/${sessionId}`, { method: 'POST' });
    const genData = await genRes.json();
    if (!genRes.ok) {
      throw new Error(genData.error || 'Generation failed');
    }

    if (genData.status === 'queued' && genData.queuePosition > 0) {
      progressPhase.textContent = 'Waiting for space to process...';
      progressStep.textContent = `Position ${genData.queuePosition}`;
      progressFill.style.width = '5%';
    }
    startPolling();

  } catch (err) {
    showError(err.message);
  }
}

function uploadWithProgress(file) {
  return new Promise((resolve, reject) => {
    const xhr = new XMLHttpRequest();
    const formData = new FormData();
    formData.append('file', file);

    xhr.upload.addEventListener('progress', (e) => {
      if (e.lengthComputable) {
        const percent = Math.round((e.loaded / e.total) * 90);
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
  document.getElementById('upload-section').hidden = true;
  progressSection.hidden = false;
  resultSection.hidden = true;
  updateProgressUI('Preparing...', 0, PHASES.length);
}

function updateProgressUI(phase, phaseIndex, total) {
  progressPhase.textContent = phase;
  if (phaseIndex === -1) {
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
      setTimeout(() => autoDownload(data), 400);
    } else if (data.status === 'failed') {
      stopPolling();
      showError(data.error || 'Conversion failed.');
    }
  } catch {
    stopPolling();
    showError('Lost connection to server.');
  }
}

// ── Auto-download on completion ───────────────────────────────

function autoDownload(data) {
  const url = `/api/download/${sessionId}`;

  // Trigger download via hidden link
  const a = document.createElement('a');
  a.href = url;
  a.download = 'handout.docx';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);

  // Show result screen
  progressSection.hidden = true;
  resultSection.hidden = false;
  resultSuccess.hidden = false;
  resultError.hidden = true;

  resultMeta.textContent = `${data.slideCount} slides converted`;
  downloadLink.href = url;

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
  resultSection.hidden = false;
  resultSuccess.hidden = true;
  resultError.hidden = false;
  errorMessage.textContent = message;
}

// ── Start over ─────────────────────────────────────────────────

function resetUI() {
  stopPolling();
  sessionId = null;
  fileInput.value = '';

  document.getElementById('upload-section').hidden = false;
  dropZone.hidden = false;
  progressSection.hidden = true;
  resultSection.hidden = true;
  progressFill.style.width = '0%';
}

startOverBtn.addEventListener('click', resetUI);
