# Pamphlet

A privacy-first web utility that converts PowerPoint (.pptx) lecture slides into clean, editable Word (.docx) handouts. Designed for instructors and academics.

**Live at** [pamphlet.polivoxia.ca](https://pamphlet.polivoxia.ca)

## Features

- **Verbatim Text Fidelity** — all text extracted exactly as-is, never added, modified, or deleted
- **Visual Formatting** — fonts, colors, bold, italic, underline mapped directly with run-level formatting (no Word Heading styles)
- **Typography Scaling** — slide-sized fonts proportionally scaled to document-appropriate sizes (10-22pt range)
- **Spatial Ordering** — shapes sorted by their Y/X position on the slide so the document reads in visual order
- **Image Support** — images extracted, optimized via sharp (82% size reduction typical), alt-text preserved
- **Smart Layout** — slides flow naturally across pages with horizontal rule separators; `keepNext` prevents orphaned headers; short slides share pages
- **Unsupported Object Detection** — SmartArt, charts, and embedded objects flagged with descriptive placeholders
- **Modern UI** — clean, "saturated-pastel" aesthetic using Inter typography and responsive layout
- **Progress Tracking** — real-time upload progress (MB/MB) and phase-based animated "candy bar" conversion progress
- **Page Numbers** — optional footer page numbers, toggled in the UI

### Privacy & Security

- **Zero Persistence** — all files purged 10 minutes after conversion via `setTimeout`, with a cron sweep failsafe every 5 minutes
- **Graceful Shutdown** — full storage scrub + session purge on `SIGTERM`/`SIGINT`
- **Queue Management** — 3 concurrent conversions, 10 pending max, 5-minute task timeout, 100 session cap
- **Rate Limiting** — 2 uploads/IP/min, 1 download/5s, 60 status polls/min, keyed on Cloudflare `CF-Connecting-IP`
- **Input Validation** — ZIP magic byte check, .pptx extension filter, 50MB size limit, UUID session ID format validation
- **ZIP Bomb Guard** — rejects PPTX files with >200MB uncompressed content
- **Path Traversal Protection** — image paths validated to stay within `ppt/media/`
- **Security Headers** — helmet CSP, HSTS, Referrer-Policy, Permissions-Policy
- **Error Sanitization** — no internal paths or stack traces exposed to clients

## Quick Start

```bash
npm install
npm run build
npm start
```

## Development

```bash
npm install
npm run dev       # Vite watch + Express with --watch
npm test          # Run 55 tests (unit + integration)
npm run test:watch
```

## Production Deployment

```bash
# Build frontend
npm run build

# Start with PM2
pm2 start ecosystem.config.cjs

# Install cron failsafe (purges orphaned files every 5 min)
crontab -e
# Add: */5 * * * * /path/to/deploy/cron-sweep.sh
```

### Nginx

Copy `deploy/pamphlet.polivoxia.ca.conf` to `/etc/nginx/sites-available/` and enable it. Includes Cloudflare real IP ranges, 50M upload limit, gzip, security headers, and SSL via certbot.

### Environment

| Variable | Description | Default |
|---|---|---|
| `PAMPHLET_STORAGE_ROOT` | Ephemeral file storage path | `/dev/shm/pamphlet` |
| `PORT` | Server port | `3000` |
| `MAX_FILE_SIZE_MB` | Upload size limit | `50` |
| `CONCURRENCY_LIMIT` | Max concurrent conversions | `3` |
| `MAX_QUEUE_DEPTH` | Max pending conversions | `10` |
| `CLEANUP_MINUTES` | File expiry time | `10` |

## API

| Method | Endpoint | Description |
|---|---|---|
| `POST` | `/api/upload` | Upload a .pptx file, returns `{ id, status }` |
| `POST` | `/api/generate/:id` | Start conversion, returns `{ id, status, queuePosition }` |
| `GET` | `/api/status/:id` | Poll progress: phase, phaseIndex, queue info |
| `GET` | `/api/download/:id` | Download the converted .docx |

All `:id` endpoints validate UUID format. Rate limits return `429` with descriptive messages. Queue at capacity returns `503`.

## Tech Stack

- **Runtime:** Node.js >= 22
- **Backend:** Express.js (ESM)
- **Frontend:** Vanilla JS + Modern CSS (Custom Properties, Gradients, Animations), Inter typography, built with Vite
- **PPTX Parsing:** jszip + @xmldom/xmldom with namespace-aware XML traversal
- **DOCX Generation:** docx (npm)
- **Image Optimization:** sharp
- **Queueing:** p-queue (with timeout)
- **Security:** helmet, express-rate-limit
- **Process Manager:** PM2
- **Reverse Proxy:** Nginx + Cloudflare
- **SSL:** Let's Encrypt via certbot

## License

[Creative Commons Attribution 4.0 International (CC BY 4.0)](https://creativecommons.org/licenses/by/4.0/). See [LICENSE](LICENSE).
