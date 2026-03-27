import multer from 'multer';
import path from 'node:path';
import fs from 'node:fs';
import { UPLOADS_DIR } from '../services/storage.js';

const MAX_FILE_SIZE = (parseInt(process.env.MAX_FILE_SIZE_MB, 10) || 50) * 1024 * 1024;

const storage = multer.diskStorage({
  destination: (_req, _file, cb) => {
    fs.mkdirSync(UPLOADS_DIR, { recursive: true });
    cb(null, UPLOADS_DIR);
  },
  filename: (_req, file, cb) => {
    const unique = Date.now() + '-' + Math.round(Math.random() * 1e6);
    cb(null, unique + path.extname(file.originalname));
  },
});

function fileFilter(_req, file, cb) {
  const ext = path.extname(file.originalname).toLowerCase();
  if (ext !== '.pptx') {
    return cb(new Error('Only .pptx files are accepted'), false);
  }
  cb(null, true);
}

export const upload = multer({
  storage,
  fileFilter,
  limits: { fileSize: MAX_FILE_SIZE },
});
