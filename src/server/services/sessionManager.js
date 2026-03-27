import { v4 as uuidv4 } from 'uuid';
import path from 'node:path';
import { UPLOADS_DIR, DOWNLOADS_DIR, removeFile } from './storage.js';

const sessions = new Map();

export function createSession() {
  const id = uuidv4();
  const session = {
    id,
    uploadPath: null,
    downloadPath: path.join(DOWNLOADS_DIR, `${id}.docx`),
    status: 'created',    // created → uploaded → queued → processing → ready → failed → expired
    phase: null,           // current conversion phase
    phaseIndex: 0,
    totalPhases: 4,
    unsupportedObjects: [],
    createdAt: Date.now(),
    expiryTimer: null,
  };
  sessions.set(id, session);
  return session;
}

export function getSession(id) {
  return sessions.get(id) || null;
}

export function setUploaded(id, uploadPath) {
  const session = sessions.get(id);
  if (!session) return null;
  session.uploadPath = uploadPath;
  session.status = 'uploaded';
  return session;
}

export function updateProgress(id, { phase, phaseIndex, total }) {
  const session = sessions.get(id);
  if (!session) return;
  session.phase = phase;
  session.phaseIndex = phaseIndex;
  session.totalPhases = total;
}

export function setStatus(id, status, extra = {}) {
  const session = sessions.get(id);
  if (!session) return null;
  session.status = status;
  Object.assign(session, extra);
  return session;
}

export function startExpiryTimer(id, minutes = 10) {
  const session = sessions.get(id);
  if (!session) return;

  if (session.expiryTimer) clearTimeout(session.expiryTimer);

  session.expiryTimer = setTimeout(() => {
    purgeSession(id);
  }, minutes * 60 * 1000);
}

export function purgeSession(id) {
  const session = sessions.get(id);
  if (!session) return;

  if (session.expiryTimer) clearTimeout(session.expiryTimer);
  if (session.uploadPath) removeFile(session.uploadPath);
  if (session.downloadPath) removeFile(session.downloadPath);

  session.status = 'expired';
  sessions.delete(id);
}

export function purgeAll() {
  for (const id of sessions.keys()) {
    purgeSession(id);
  }
}

export function sessionCount() {
  return sessions.size;
}
