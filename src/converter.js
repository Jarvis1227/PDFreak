import { Document, Packer, Paragraph, TextRun, ImageRun, SectionType, LineRuleType } from 'docx';

export const MAX_PAGES = 60;
export const MAX_BYTES = 25 * 1024 * 1024;

export function parsePages(value, total) {
  if (!Number.isInteger(total) || total < 1) throw new Error('This PDF has no pages.');
  if (!value.trim()) {
    if (total > MAX_PAGES) throw new Error(`Choose up to ${MAX_PAGES} pages, for example 1-${MAX_PAGES}.`);
    return Array.from({ length: total }, (_, i) => i + 1);
  }
  const pages = new Set();
  for (const part of value.split(',')) {
    const match = part.trim().match(/^(\d+)(?:\s*-\s*(\d+))?$/);
    if (!match) throw new Error('Use page numbers like 1-3, 5, 8.');
    const start = Number(match[1]), end = Number(match[2] || match[1]);
    if (start < 1 || end > total || end < start) throw new Error(`Page numbers must be between 1 and ${total}, with ranges in ascending order.`);
    if (end - start + 1 > MAX_PAGES) throw new Error(`Convert up to ${MAX_PAGES} pages at a time.`);
    for (let n = start; n <= end; n++) {
      pages.add(n);
      if (pages.size > MAX_PAGES) throw new Error(`Convert up to ${MAX_PAGES} pages at a time.`);
    }
  }
  return [...pages].sort((a, b) => a - b);
}

export function outputName(name) {
  const stem = name.replace(/\.pdf$/i, '').replace(/[<>:"/\\|?*\x00-\x1F]/g, '_').trim() || 'document';
  return `${stem}.docx`;
}

// PDF text positions are transformed to the displayed page coordinates first.
// Then text fragments on the same baseline become one editable Word paragraph.
export function textRows(content, viewport) {
  const t = viewport.transform;
  const pieces = content.items.filter(item => item.str && item.str.trim()).map(item => {
    const m = item.transform;
    const x = t[0] * m[4] + t[2] * m[5] + t[4];
    const y = t[1] * m[4] + t[3] * m[5] + t[5];
    const size = Math.max(4, Math.min(100, Math.hypot(m[2], m[3])));
    const style = content.styles?.[item.fontName] || {};
    const name = style.fontFamily || '';
    return { text: item.str, x, y, size, width: item.width || 0, bold: /bold|black|heavy/i.test(item.fontName + name), italics: /italic|oblique/i.test(item.fontName + name), font: /mono|courier/i.test(name) ? 'Courier New' : /sans|arial|helvetica/i.test(name) ? 'Arial' : 'Times New Roman' };
  }).sort((a, b) => a.y - b.y || a.x - b.x);
  const rows = [];
  for (const item of pieces) {
    const last = rows.at(-1);
    if (last && Math.abs(last.y - item.y) <= Math.max(2, Math.min(last.size, item.size) * .28)) {
      last.items.push(item); last.size = Math.max(last.size, item.size);
    } else rows.push({ y: item.y, size: item.size, items: [item] });
  }
  for (const row of rows) row.items.sort((a, b) => a.x - b.x);
  return rows;
}

function checkAbort(signal) {
  if (signal?.aborted) throw new DOMException('Conversion cancelled.', 'AbortError');
}

function editableSection(page, content) {
  const viewport = page.getViewport({ scale: 1 });
  const rows = textRows(content, viewport);
  if (!rows.length) return null;
  const first = rows[0];
  const left = Math.max(8, Math.min(54, ...rows.flatMap(row => row.items.map(item => item.x))));
  const top = Math.max(5, first.y - first.size - 1);
  const children = rows.map((row, i) => {
    const runs = []; let previous;
    for (const item of row.items) {
      if (previous) {
        const gap = item.x - (previous.x + previous.width);
        if (gap > item.size * .15 && !previous.text.endsWith(' ') && !item.text.startsWith(' ')) {
          runs.push(new TextRun({ text: ' '.repeat(Math.min(20, Math.max(1, Math.round(gap / (item.size * .5))))), size: item.size * 2 }));
        }
      }
      runs.push(new TextRun({ text: item.text, size: Math.round(item.size * 2), font: item.font, bold: item.bold, italics: item.italics }));
      previous = item;
    }
    const previousRow = rows[i - 1];
    const gap = previousRow ? Math.max(0, row.y - previousRow.y - row.size * 1.08) : 0;
    return new Paragraph({ children: runs, indent: { left: Math.max(0, Math.round((row.items[0].x - left) * 20)) }, spacing: { before: Math.round(gap * 20), after: 0, line: Math.round(row.size * 1.08 * 20), lineRule: LineRuleType.EXACT }, widowControl: false });
  });
  return { properties: { type: SectionType.NEXT_PAGE, page: { size: { width: Math.round(viewport.width * 20), height: Math.round(viewport.height * 20) }, margin: { top: Math.round(top * 20), bottom: 80, left: Math.round(left * 20), right: 100, header: 0, footer: 0 } } }, children };
}

async function appearanceSection(page) {
  const natural = page.getViewport({ scale: 1 });
  // Cap image dimensions so long pages cannot allocate unbounded canvases.
  const scale = Math.min(1.75, 2400 / Math.max(natural.width, natural.height));
  const viewport = page.getViewport({ scale });
  const canvas = document.createElement('canvas');
  canvas.width = Math.ceil(viewport.width); canvas.height = Math.ceil(viewport.height);
  try {
    await page.render({ canvasContext: canvas.getContext('2d'), viewport, background: '#ffffff' }).promise;
    const blob = await new Promise(resolve => canvas.toBlob(resolve, 'image/jpeg', .9));
    if (!blob) throw new Error('Your browser could not render this page. Try fewer pages.');
    const width = natural.width - 36, height = natural.height - 40;
    const factor = Math.min(width / natural.width, height / natural.height);
    const image = new ImageRun({ data: new Uint8Array(await blob.arrayBuffer()), type: 'jpg', transformation: { width: natural.width * factor * 96 / 72, height: natural.height * factor * 96 / 72 }, altText: { title: 'Original PDF page', description: 'This PDF page is preserved as an image. Its text is not editable.', name: 'PDF page' } });
    return { properties: { type: SectionType.NEXT_PAGE, page: { size: { width: Math.round(natural.width * 20), height: Math.round(natural.height * 20) }, margin: { top: 360, bottom: 360, left: 360, right: 360, header: 0, footer: 0 } } }, children: [new Paragraph({ children: [image], spacing: { before: 0, after: 0, line: 1, lineRule: LineRuleType.EXACT } })] };
  } finally { canvas.width = 0; canvas.height = 0; }
}

export async function convertPdf(pdf, pages, mode, onProgress = () => {}, signal) {
  const sections = [], imageOnly = [];
  for (let i = 0; i < pages.length; i++) {
    checkAbort(signal);
    onProgress(Math.round(i / pages.length * 88), `Converting page ${pages[i]} · ${i + 1} of ${pages.length}`);
    const page = await pdf.getPage(pages[i]);
    if (mode === 'appearance') sections.push(await appearanceSection(page));
    else {
      const section = editableSection(page, await page.getTextContent());
      if (section) sections.push(section);
      else imageOnly.push(pages[i]);
    }
    checkAbort(signal);
    // Let the browser paint progress and handle the Cancel button.
    await new Promise(resolve => setTimeout(resolve, 0));
  }
  if (imageOnly.length) throw new Error(`Page${imageOnly.length === 1 ? '' : 's'} ${imageOnly.join(', ')} ${imageOnly.length === 1 ? 'has' : 'have'} no selectable text. Choose “Exact appearance” to include scanned or blank pages. OCR is not available in this version.`);
  checkAbort(signal);
  onProgress(94, 'Building your Word document…');
  const document = new Document({ creator: 'PDFreak', title: 'Converted document', description: mode === 'appearance' ? 'PDF pages preserved as images' : 'Editable text converted from PDF', sections, styles: { default: { document: { run: { font: 'Arial', size: 22 }, paragraph: { spacing: { after: 0 } } } } } });
  const blob = await Packer.toBlob(document);
  checkAbort(signal); onProgress(100, 'Your Word document is ready.');
  return blob;
}
