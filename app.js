import 'dotenv/config';
import { createApp, startServer } from './src/server/index.js';

process.on('unhandledRejection', (reason, promise) => {
  console.error('Unhandled Rejection:', reason);
});

process.on('uncaughtException', (err) => {
  console.error('Uncaught Exception:', err);
  process.exit(1);
});

const app = createApp();
const port = process.env.PORT || 3000;

startServer(app, port);
