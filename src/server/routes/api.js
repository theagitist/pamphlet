import { Router } from 'express';
import fs from 'node:fs';
import path from 'node:path';
import { upload } from '../middleware/upload.js';
import { uploadLimiter, downloadLimiter } from '../middleware/rateLimiter.js';
import {
  createSession, getSession, setUploaded,
  updateProgress, setStatus, startExpiryTimer,
} from '../services/sessionManager.js';
import { enqueue, getQueueInfo } from '../services/queueManager.js';
import { convert } from '../services/converter.js';

const router = Router();

// POST /api/upload — upload a .pptx file, returns session id
router.post('/upload', uploadLimiter, upload.single('file'), (req, res) => {
  if (!req.file) {
    return res.status(400).json({ error: 'No file uploaded' });
  }

  const session = createSession();
  setUploaded(session.id, req.file.path);

  res.json({
    id: session.id,
    status: session.status,
  });
});

// Handle multer errors (file too large, wrong type)
router.use((err, _req, res, _next) => {
  if (err.code === 'LIMIT_FILE_SIZE') {
    return res.status(413).json({ error: 'File too large. Maximum size is 50MB.' });
  }
  if (err.message?.includes('.pptx')) {
    return res.status(400).json({ error: err.message });
  }
  return res.status(500).json({ error: 'Upload failed' });
});

// POST /api/generate/:id — trigger conversion (queued)
router.post('/generate/:id', async (req, res) => {
  const session = getSession(req.params.id);
  if (!session) {
    return res.status(404).json({ error: 'Session not found' });
  }
  if (session.status !== 'uploaded') {
    return res.status(409).json({ error: `Cannot generate: status is "${session.status}"` });
  }

  // Enqueue the conversion — returns null if at capacity
  const queued = enqueue(async () => {
    setStatus(session.id, 'processing');
    try {
      const result = await convert(session.uploadPath, session.downloadPath, (progress) => {
        updateProgress(session.id, progress);
      });
      setStatus(session.id, 'ready', {
        unsupportedObjects: result.unsupportedObjects,
        slideCount: result.slides,
      });
      startExpiryTimer(session.id, parseInt(process.env.CLEANUP_MINUTES, 10) || 10);
    } catch (err) {
      console.error('Conversion failed:', err);
      setStatus(session.id, 'failed', { error: err.message });
      startExpiryTimer(session.id, 1);
    }
  });

  if (queued === null) {
    return res.status(503).json({
      error: 'Server is at capacity. Please try again shortly.',
      queue: getQueueInfo(),
    });
  }

  setStatus(session.id, 'queued');
  const queue = getQueueInfo();
  res.json({ id: session.id, status: 'queued', queuePosition: queue.waiting });
});

// GET /api/status/:id — poll conversion progress
router.get('/status/:id', (req, res) => {
  const session = getSession(req.params.id);
  if (!session) {
    return res.status(404).json({ error: 'Session not found' });
  }

  const response = {
    id: session.id,
    status: session.status,
    phase: session.phase,
    phaseIndex: session.phaseIndex,
    totalPhases: session.totalPhases,
  };

  if (session.status === 'queued') {
    response.queue = getQueueInfo();
  }

  if (session.status === 'ready') {
    response.slideCount = session.slideCount;
    response.unsupportedObjects = session.unsupportedObjects;
  }

  if (session.status === 'failed') {
    response.error = session.error;
  }

  res.json(response);
});

// GET /api/download/:id — download converted docx
router.get('/download/:id', downloadLimiter, (req, res) => {
  const session = getSession(req.params.id);
  if (!session) {
    return res.status(404).json({ error: 'Session not found' });
  }
  if (session.status !== 'ready') {
    return res.status(409).json({ error: `Cannot download: status is "${session.status}"` });
  }

  if (!fs.existsSync(session.downloadPath)) {
    return res.status(410).json({ error: 'File expired or not found' });
  }

  const filename = 'handout.docx';
  res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');

  const stream = fs.createReadStream(session.downloadPath);
  stream.pipe(res);
});

export default router;
