import fs from 'node:fs';
import path from 'node:path';

const STORAGE_ROOT = process.env.PAMPHLET_STORAGE_ROOT || '/dev/shm/pamphlet';

export const UPLOADS_DIR = path.join(STORAGE_ROOT, 'uploads');
export const DOWNLOADS_DIR = path.join(STORAGE_ROOT, 'downloads');

export function initStorage() {
  ensureDirs();
  console.log(`Storage initialized at ${STORAGE_ROOT}`);
}

// Ensure storage dirs exist — safe to call multiple times
export function ensureDirs() {
  fs.mkdirSync(UPLOADS_DIR, { recursive: true });
  fs.mkdirSync(DOWNLOADS_DIR, { recursive: true });
}

export function scrubAll() {
  for (const dir of [UPLOADS_DIR, DOWNLOADS_DIR]) {
    if (fs.existsSync(dir)) {
      fs.rmSync(dir, { recursive: true, force: true });
    }
  }
  console.log('Storage scrubbed.');
}

export function removeFile(filePath) {
  if (fs.existsSync(filePath)) {
    fs.unlinkSync(filePath);
  }
}
