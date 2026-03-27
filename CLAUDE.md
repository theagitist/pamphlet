# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

**Pamphlet** is a privacy-first web utility that converts PowerPoint (.pptx) lecture slides into clean, editable Word (.docx) handouts with verbatim text fidelity. It preserves original typography, colors, and layout without imposing semantic structures (no Heading styles — direct run-level formatting only).

## Tech Stack

- **Runtime:** Node.js >= 22
- **Backend:** Express.js (middleware-based, ESM)
- **Frontend:** Vanilla JS/CSS built with Vite, served by Express
- **PPTX Parsing:** `jszip` + `@xmldom/xmldom` with namespace-aware XML traversal
- **Word Generation:** `docx` npm package
- **Image Optimization:** `sharp`
- **Queueing:** `p-queue` (3 concurrent, 10 pending max)
- **Security:** `helmet` (CSP), `express-rate-limit` (keyed on `CF-Connecting-IP`), `multer` (50MB, .pptx only)
- **Process Manager:** PM2 (`ecosystem.config.cjs`)
- **Reverse Proxy:** Nginx + Cloudflare
- **SSL:** Let's Encrypt via certbot

## Key Design Rules

- **Strict Verbatim:** No text may be added, modified, or deleted during conversion.
- **Visual Fidelity over Semantics:** Direct run-level formatting (font, color, size) on every text node. Never use Word `Heading` styles.
- **Aesthetic:** Clean "saturated-pastel" theme (Coral, Teal, Yellow) with Inter typography.
- **UI Progress:** Animated "candy bar" progress for conversion; real-time MB tracking for uploads.
- **Typography Scaling:** Slide-sized fonts are proportionally scaled down for document readability (max 22pt heading, 14pt body cap, 10pt floor).
- **Spatial Ordering:** Elements are sorted by Y then X position from `a:xfrm/a:off`, not XML tree order.
- **Unsupported Objects:** Insert placeholder `[Object not supported: {Description}]` for SmartArt/charts/OLE objects.
- **Images:** Optimized via sharp before embedding. Background and decorative images (full-slide-sized, <4KB) are filtered out. Full-height side-panel images are placed after text content.
- **Slide Layout:** Slides flow in a single section with horizontal rules between them. All paragraphs within a slide have `keepNext` to prevent orphaned headers. An empty terminator paragraph after each slide breaks the chain.

## Architecture

```
src/server/
  converter/
    xmlHelper.js        # Namespace-aware DOM traversal helpers
    pptxParser.js       # PPTX XML parsing, theme/color/style inheritance
    styleMapper.js      # PPTX props → docx TextRun options, typography scaler
    imageOptimizer.js   # sharp-based image compression
    docxWriter.js       # Builds Document with spatial ordering, keepNext, page numbers
  services/
    storage.js          # PAMPHLET_STORAGE_ROOT init/scrub/cleanup
    sessionManager.js   # In-memory UUID-keyed session Map with expiry timers
    queueManager.js     # p-queue wrapper (3 concurrent, 10 pending)
    converter.js        # 4-phase orchestrator: parse → extract → format → write
  middleware/
    upload.js           # multer config (.pptx, 50MB)
    rateLimiter.js      # express-rate-limit (CF-Connecting-IP aware)
    security.js         # helmet CSP
  routes/
    api.js              # POST upload, POST generate, GET status, GET download
  index.js              # Express app factory + server with graceful shutdown
```

## Build & Run

```bash
npm install
npm run build          # Vite builds src/client/ → dist/client/
npm start              # node app.js
npm test               # vitest run
pm2 start ecosystem.config.cjs   # production
```

## Environment Variables

- `PAMPHLET_STORAGE_ROOT` — ephemeral file storage (default `/dev/shm/pamphlet`)
- `PORT` — server port (default `3000`)
- `MAX_FILE_SIZE_MB` — upload limit (default `50`)
- `CONCURRENCY_LIMIT` — max concurrent conversions (default `3`)
- `MAX_QUEUE_DEPTH` — max pending conversions (default `10`)
- `CLEANUP_MINUTES` — file expiry time (default `10`)

## Deployment

- **Live:** https://pamphlet.polivoxia.ca
- **Nginx config:** `deploy/pamphlet.polivoxia.ca.conf`
- **PM2 config:** `ecosystem.config.cjs`
- **Cron failsafe:** `deploy/cron-sweep.sh` (every 5 min, purges files >12 min old)
