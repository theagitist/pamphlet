# Pamphlet

A privacy-first web utility that converts PowerPoint (.pptx) lecture slides into clean, editable Word (.docx) handouts. Designed for instructors and academics.

**Live at** [pamphlet.polivoxia.ca](https://pamphlet.polivoxia.ca)

## Features

- **Verbatim Text Fidelity** — all text is extracted exactly as-is, never added, modified, or deleted
- **Visual Formatting** — fonts, colors, bold, italic, underline mapped directly from PPTX to DOCX with run-level formatting (no Word Heading styles)
- **Typography Scaling** — slide-sized fonts (40pt+) are proportionally scaled to document-appropriate sizes (10-22pt)
- **Spatial Ordering** — shapes are sorted by their position on the slide (top-to-bottom, left-to-right) so the document reads in visual order
- **Image Support** — images extracted, optimized via sharp, and embedded with existing alt-text preserved
- **Smart Layout** — slides flow naturally across pages with horizontal rule separators; short slides share pages, headers never orphan from their content
- **Unsupported Object Detection** — SmartArt, charts, and embedded objects are flagged with descriptive placeholders
- **Page Numbers** — optional, toggled in the UI
- **Privacy First** — all files purged 10 minutes after conversion via setTimeout + cron failsafe; full storage scrub on shutdown
- **Queue Management** — 3 concurrent conversions, 10 pending max, with real-time queue feedback in the UI
- **Rate Limiting** — 2 uploads/IP/min, 1 download/5s, keyed on Cloudflare's `CF-Connecting-IP`
- **Security** — helmet CSP, multer file validation (.pptx only, 50MB max)

## Quick Start

```bash
npm install
npm run build
npm run start
```

## Development

```bash
npm install
npm run dev       # Vite watch + Express with --watch
npm test          # Run all tests
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

## Tech Stack

- **Runtime:** Node.js >= 22
- **Backend:** Express.js
- **Frontend:** Vanilla JS + CSS, built with Vite
- **PPTX Parsing:** jszip + @xmldom/xmldom with namespace-aware XML traversal
- **DOCX Generation:** docx (npm)
- **Image Optimization:** sharp
- **Queueing:** p-queue
- **Security:** helmet, express-rate-limit
- **Process Manager:** PM2
- **Reverse Proxy:** Nginx + Cloudflare
- **SSL:** Let's Encrypt via certbot

## License

[Creative Commons Attribution 4.0 International (CC BY 4.0)](https://creativecommons.org/licenses/by/4.0/). See [LICENSE](LICENSE).
