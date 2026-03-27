# Pamphlet

Verbatim PowerPoint to Word handout converter.

A privacy-first web utility designed for instructors to transform lecture slides into clean, editable, and paginated Word documents. No data retention, no text alterations — just pure structural reformatting.

## Features

- **100% Verbatim Fidelity** — text is never added, modified, or deleted
- **Visual Inheritance** — fonts, colors, and sizes are mapped directly from PPTX to DOCX
- **Image Support** — images and existing alt-text are preserved as-is
- **Privacy First** — all uploaded and generated files are automatically purged 10 minutes after conversion
- **Queue Management** — concurrent processing with clear feedback when the server is busy

## Quick Start

```bash
npm install
npm run dev
```

## Production

```bash
npm run build
pm2 start app.js --name "pamphlet"
```

### Environment

| Variable | Description | Example |
|---|---|---|
| `PAMPHLET_STORAGE_ROOT` | Ephemeral file storage path | `/dev/shm/pamphlet` |

## Tech Stack

- **Runtime:** Node.js (>= 22 LTS)
- **Backend:** Express.js
- **Frontend:** Vite (static HTML/JS)
- **PPTX Parsing:** jszip + XML traversal
- **DOCX Generation:** docx
- **Queueing:** p-queue

## License

This project is licensed under the [Creative Commons Attribution 4.0 International License (CC BY 4.0)](https://creativecommons.org/licenses/by/4.0/). See [LICENSE](LICENSE) for details.
