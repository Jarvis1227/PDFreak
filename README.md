# PDFreak PDF to Word

A browser based converter built with PDF.js, docx, and Vite. The PDF and opening password remain in the browser; no document upload endpoint exists.

## Run

Requires Node.js 20.19+ or 22.12+.

```powershell
npm install
npm run dev
```

Open the local URL printed by Vite. For deployment, run `npm run build` and host `dist/` as a static website. The Sites identity is retained in `.openai/hosting.json`.

## Features

- File picker and drag and drop; local PDF preview with page navigation.
- Editable text with font sizing, page dimensions, line spacing, and page boundaries.
- Exact appearance conversion using rendered page images; works with scans but is not editable.
- All pages or selected page ranges, password prompt, progress, cancellation, and repeat download.
- Copy restrictions respected, invalid files rejected, 25 MB file size and 60 page conversion limits.

## Limits

This is a lightweight local text reconstruction tool. Editable mode does not reproduce images, tables, or complex multi-column reading order. It does not implement OCR. Pages without selectable text require exact appearance mode. Rendering and conversion consume browser memory, especially on mobile. Exact appearance images are rasterized at up to 1.75x resolution, capped to 2400 pixels per edge, then saved as JPEGs.

## Check

```powershell
npm test
npm run build
```

Libraries: [Mozilla PDF.js](https://mozilla.github.io/pdf.js/) and [docx](https://docx.js.org/). Their license notices are supplied in `THIRD_PARTY_NOTICES.md`.
