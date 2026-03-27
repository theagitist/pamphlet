import 'dotenv/config';
import { createApp, startServer } from './src/server/index.js';

const app = createApp();
const port = process.env.PORT || 3000;

startServer(app, port);
