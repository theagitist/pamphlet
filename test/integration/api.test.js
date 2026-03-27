import { describe, it, expect, beforeAll, afterAll } from 'vitest';
import { resolve, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import supertest from 'supertest';
import { createApp } from '../../src/server/index.js';
import { initStorage, scrubAll } from '../../src/server/services/storage.js';

const __dirname = dirname(fileURLToPath(import.meta.url));
const fixturePath = resolve(__dirname, '../fixtures/basic.pptx');

let app;
let request;

beforeAll(() => {
  process.env.PAMPHLET_STORAGE_ROOT = '/tmp/pamphlet-test-' + Date.now();
  initStorage();
  app = createApp();
  request = supertest(app);
});

afterAll(() => {
  scrubAll();
});

describe('API routes', () => {
  let sessionId;

  describe('POST /api/upload', () => {
    it('accepts a .pptx file and returns session id', async () => {
      const res = await request
        .post('/api/upload')
        .attach('file', fixturePath)
        .expect(200);

      expect(res.body.id).toBeDefined();
      expect(res.body.status).toBe('uploaded');
      sessionId = res.body.id;
    });

    it('rejects non-.pptx files', async () => {
      const res = await request
        .post('/api/upload')
        .attach('file', Buffer.from('not a pptx'), 'test.txt')
        .expect(400);

      expect(res.body.error).toContain('.pptx');
    });
  });

  describe('POST /api/generate/:id', () => {
    it('enqueues conversion and returns status', async () => {
      const res = await request
        .post(`/api/generate/${sessionId}`)
        .expect(200);

      expect(['queued', 'processing']).toContain(res.body.status);
    });

    it('rejects generate on non-existent session', async () => {
      await request
        .post('/api/generate/nonexistent-id')
        .expect(404);
    });

    it('rejects double-generate', async () => {
      const res = await request
        .post(`/api/generate/${sessionId}`)
        .expect(409);

      expect(res.body.error).toContain('Cannot generate');
    });
  });

  describe('GET /api/status/:id', () => {
    it('returns current status', async () => {
      const res = await request
        .get(`/api/status/${sessionId}`)
        .expect(200);

      expect(res.body.id).toBe(sessionId);
      expect(['queued', 'processing', 'ready']).toContain(res.body.status);
    });

    it('returns 404 for unknown session', async () => {
      await request
        .get('/api/status/nonexistent-id')
        .expect(404);
    });

    it('eventually reaches ready status', async () => {
      let status = 'queued';
      let attempts = 0;
      while (status !== 'ready' && attempts < 30) {
        await new Promise(r => setTimeout(r, 200));
        const res = await request.get(`/api/status/${sessionId}`);
        status = res.body.status;
        attempts++;
      }
      expect(status).toBe('ready');
    }, 10000);
  });

  describe('GET /api/download/:id', () => {
    it('downloads the converted docx', async () => {
      const res = await request
        .get(`/api/download/${sessionId}`)
        .responseType('blob')
        .expect(200);

      expect(res.headers['content-type']).toContain('wordprocessingml');
      expect(res.headers['content-disposition']).toContain('handout.docx');
      expect(res.body.byteLength || res.body.length).toBeGreaterThan(0);
    });
  });
});
