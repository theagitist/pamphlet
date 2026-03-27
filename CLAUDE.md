# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

**Pamphlet** is a privacy-first web utility that converts PowerPoint (.pptx) lecture slides into clean, editable Word (.docx) handouts with 100% verbatim text fidelity. It preserves original typography, colors, and layout without imposing semantic structures (no Heading styles — direct run-level formatting only).

## Tech Stack

- **Runtime:** Node.js LTS (>= 22)
- **Backend:** Express.js (middleware-based)
- **Frontend:** Static HTML/JS built with Vite, served by Express
- **PPTX Parsing:** `jszip` + direct XML traversal
- **Word Generation:** `docx` npm package
- **Queueing:** `p-queue` (3 concurrent, 10 pending max)
- **Security:** `helmet.js`, `express-rate-limit` (keyed on `CF-Connecting-IP`), `multer` (50MB limit)
- **Process Manager:** PM2
- **Reverse Proxy:** Nginx (`client_max_body_size 50M`)

## Key Design Rules

- **Strict Verbatim:** No text may be added, modified, or deleted during conversion. 1:1 text match between PPTX and DOCX.
- **Visual Fidelity over Semantics:** Apply direct run-level formatting (RGB/HEX color, font family, font size) to every text node. Never use Word `Heading` styles.
- **Unsupported Objects:** Insert descriptive placeholder `[Object not supported: {Description}]` for 3D/SmartArt/complex objects.
- **Images:** Render images in output; preserve existing alt-text verbatim; never add alt-text.

## Storage & Privacy

- All file storage is under `PAMPHLET_STORAGE_ROOT` (default: `/dev/shm/pamphlet`) with `/uploads` and `/downloads` subdirs.
- Files auto-purge 10 minutes after session creation via `setTimeout`, with a cron sweep as fail-safe.
- On `SIGTERM`/`SIGINT`, scrub entire `PAMPHLET_STORAGE_ROOT` before exit.
- Download links use UUIDv4 session paths (`/api/download/:uuid`).

## Rate Limits

- Uploads: 2 per IP per minute
- Downloads: 1 per 5 seconds per session ID

## Build & Run Commands

```bash
# Install dependencies
npm install

# Development (Vite dev server + Express)
npm run dev

# Production build
npm run build

# Start production server
pm2 start app.js --name "pamphlet"

# Stop (triggers graceful shutdown + storage scrub)
pm2 stop pamphlet
```

## Environment

Required `.env` variables:
- `PAMPHLET_STORAGE_ROOT` — path for ephemeral file storage (e.g., `/dev/shm/pamphlet`)
