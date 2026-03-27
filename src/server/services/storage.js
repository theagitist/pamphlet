import fs from 'node:fs';
import path from 'node:path';

const STORAGE_ROOT = process.env.PAMPHLET_STORAGE_ROOT || '/mnt/polivoxiadata/pamphlet.polivoxia.ca';

export const UPLOADS_DIR = path.join(STORAGE_ROOT, 'uploads');
export const DOWNLOADS_DIR = path.join(STORAGE_ROOT, 'downloads');
export const WORK_DIR = path.join(STORAGE_ROOT, 'work');

export function initStorage() {
  ensureDirs();
  console.log(`Storage initialized at ${STORAGE_ROOT}`);
}

// Ensure storage dirs exist — safe to call multiple times
export function ensureDirs() {
  fs.mkdirSync(UPLOADS_DIR, { recursive: true });
  fs.mkdirSync(DOWNLOADS_DIR, { recursive: true });
  fs.mkdirSync(WORK_DIR, { recursive: true });
}

// Create a per-session work directory for intermediate files (extracted images, etc.)
export function createWorkDir(sessionId) {
  const dir = path.join(WORK_DIR, sessionId);
  fs.mkdirSync(dir, { recursive: true });
  return dir;
}

// Remove an entire directory tree (for session work dirs)
export function removeDir(dirPath) {
  if (dirPath && fs.existsSync(dirPath)) {
    fs.rmSync(dirPath, { recursive: true, force: true });
  }
}

export function scrubAll() {
  for (const dir of [UPLOADS_DIR, DOWNLOADS_DIR, WORK_DIR]) {
    if (fs.existsSync(dir)) {
      fs.rmSync(dir, { recursive: true, force: true });
    }
  }
  console.log('Storage scrubbed.');
}

export function removeFile(filePath) {
  if (filePath && fs.existsSync(filePath)) {
    fs.unlinkSync(filePath);
  }
}
