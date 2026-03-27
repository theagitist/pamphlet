import express from 'express';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { initStorage, scrubAll } from './services/storage.js';
import { purgeAll } from './services/sessionManager.js';
import apiRoutes from './routes/api.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

export function createApp() {
  const app = express();

  // API routes
  app.use('/api', apiRoutes);

  // Serve Vite-built static frontend
  const clientDir = path.resolve(__dirname, '../../dist/client');
  app.use(express.static(clientDir));

  return app;
}

export function startServer(app, port) {
  initStorage();

  const server = app.listen(port, () => {
    console.log(`Pamphlet listening on port ${port}`);
  });

  const shutdown = () => {
    console.log('Shutting down — purging sessions and scrubbing storage...');
    purgeAll();
    scrubAll();
    server.close(() => process.exit(0));
  };

  process.on('SIGTERM', shutdown);
  process.on('SIGINT', shutdown);

  return server;
}
