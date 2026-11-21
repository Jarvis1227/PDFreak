import './styles.css';
import * as pdfjs from 'pdfjs-dist';
import workerUrl from 'pdfjs-dist/build/pdf.worker.min.mjs?url';
import { MAX_BYTES, parsePages, outputName, convertPdf } from './converter.js';

pdfjs.GlobalWorkerOptions.workerSrc = workerUrl;
const icons = {
  shield: '<path d="M12 3 4 6v5c0 5 8 10 8 10s8-5 8-10V6z"/><path d="m8 12 3 3 5-6"/>',
  lock: '<rect x="5" y="10" width="14" height="11" rx="2"/><path d="M8 10V7a4 4 0 0 1 8 0v3"/>',
  upload: '<path d="M12 16V3m-5 5 5-5 5 5M4 16v4h16v-4"/>',
  download: '<path d="M12 3v13m-5-5 5 5 5-5M4 17v4h16v-4"/>',
  plus: '<path d="M12 5v14M5 12h14"/>', close: '<path d="m6 6 12 12M6 18 18 6"/>',
  files: '<path d="M8 5V2h8l5 5v12h-3M3 6h9l5 5v11H3z"/><path d="M12 6v5h5M7 15h6M7 18h6"/>',
  file: '<path d="M6 3h8l5 5v13H6zM14 3v6h5M9 13h7M9 16h5"/>',
  type: '<path d="M4 5h16M12 5v15M8 20h8M4 5v3m16-3v3"/>',
  image: '<rect x="3" y="3" width="18" height="18" rx="2"/><circle cx="8" cy="8" r="1.5"/><path d="m3 17 5-5 4 4 4-6 5 7"/>',
  word: '<path d="M8 3h9l4 4v14H8M17 3v5h4"/><rect x="2" y="8" width="13" height="11" rx="1"/><path d="m5 11 1.5 5L8.5 12l2 4 1.5-5"/>',
  swap: '<path d="M4 8h16m-4-4 4 4-4 4M20 16H4m4-4-4 4 4 4"/>',
  tune: '<path d="M4 6h7m4 0h5M4 12h2m4 0h10M4 18h10m4 0h2"/><circle cx="13" cy="6" r="2"/><circle cx="8" cy="12" r="2"/><circle cx="16" cy="18" r="2"/>',
  'chevron-left': '<path d="m14 6-6 6 6 6"/>', 'chevron-right': '<path d="m10 6 6 6-6 6"/>'
};
for (const node of document.querySelectorAll('[data-icon]')) node.innerHTML = `<svg viewBox="0 0 24 24" aria-hidden="true">${icons[node.dataset.icon] || ''}</svg>`;
const $ = id => document.getElementById(id);
let pdf = null, file = null, loadingTask = null, loadRevision = 0, previewPage = 1, renderTask = null, renderRevision = 0, downloadingUrl = null, abortController = null, passwordCallback = null, resultName = '';
const say = message => { $('status').textContent = message; };
function error(message) { $('file-error').textContent = message; $('file-error').hidden = !message; if (message) say(message); }
function clearResult() { if (downloadingUrl) URL.revokeObjectURL(downloadingUrl); downloadingUrl = null; $('conversion-result').hidden = true; }

function reset() {
  loadRevision++; renderRevision++; renderTask?.cancel(); renderTask = null;
  loadingTask?.destroy(); loadingTask = null; pdf = null; file = null;
  passwordCallback = null; if ($('password-dialog').open) $('password-dialog').close();
  $('file-input').value = ''; $('page-range').value = '';
  $('file-card').hidden = true; $('dropzone').hidden = false; $('sample-row').hidden = false;
  $('empty-preview').hidden = false; $('preview-canvas').hidden = true; $('preview-loading').hidden = true;
  $('preview-count').textContent = 'No file selected'; $('preview-navigation').hidden = true;
  $('preview-description').textContent = 'Your document, your device.'; $('convert-button').disabled = true;
  $('conversion-progress').hidden = true; clearResult(); error('');
}

async function showPage(number) {
  if (!pdf) return;
  const currentPdf = pdf, revision = ++renderRevision;
  renderTask?.cancel(); renderTask = null; $('preview-loading').hidden = false;
  try {
    const page = await currentPdf.getPage(number);
    if (revision !== renderRevision) return;
    const native = page.getViewport({ scale: 1 });
    const scale = Math.min(1.5, 720 / native.width, 1000 / native.height);
    const viewport = page.getViewport({ scale });
    // A detached canvas avoids overlapping renders when users page quickly.
    const scratch = document.createElement('canvas'); scratch.width = Math.ceil(viewport.width); scratch.height = Math.ceil(viewport.height);
    renderTask = page.render({ canvasContext: scratch.getContext('2d'), viewport, background: '#ffffff' });
    await renderTask.promise;
    if (revision !== renderRevision) return;
    const canvas = $('preview-canvas'); canvas.width = scratch.width; canvas.height = scratch.height; canvas.getContext('2d').drawImage(scratch, 0, 0);
    scratch.width = 0; scratch.height = 0;
    $('empty-preview').hidden = true; canvas.hidden = false; previewPage = number;
    $('preview-position').textContent = `${number} / ${currentPdf.numPages}`;
    $('preview-description').textContent = `Page ${number}`;
    $('previous-page').disabled = number <= 1; $('next-page').disabled = number >= currentPdf.numPages;
  } catch (e) { if (e.name !== 'RenderingCancelledException' && revision === renderRevision) error('Preview unavailable. You can still try converting this PDF.'); }
  finally { if (revision === renderRevision) { $('preview-loading').hidden = true; renderTask = null; } }
}

async function openFile(chosen) {
  if (abortController) return;
  if (!chosen) return;
  if (!/\.pdf$/i.test(chosen.name)) { error('Choose a PDF file ending in .pdf.'); return; }
  if (chosen.size > MAX_BYTES) { error('This PDF is larger than 25 MB. Choose a smaller file.'); return; }
  if (chosen.size === 0) { error('This file is empty. Choose a PDF with content.'); return; }
  reset(); const revision = loadRevision; file = chosen;
  $('file-name').textContent = chosen.name; $('file-meta').textContent = 'Opening your PDF…';
  $('dropzone').hidden = true; $('sample-row').hidden = true; $('file-card').hidden = false; $('preview-loading').hidden = false;
  say('Opening PDF.');
  try {
    const buffer = await chosen.arrayBuffer();
    if (revision !== loadRevision) return;
    const signature = new TextDecoder().decode(buffer.slice(0, 1024));
    if (!signature.includes('%PDF-')) throw new Error('This file is not a valid PDF. Try exporting it as PDF again.');
    loadingTask = pdfjs.getDocument({ data: buffer, isEvalSupported: false, cMapUrl: '/pdf-assets/cmaps/', cMapPacked: true, standardFontDataUrl: '/pdf-assets/standard_fonts/', wasmUrl: '/pdf-assets/wasm/' });
    loadingTask.onPassword = (callback, reason) => {
      if (revision !== loadRevision) return;
      passwordCallback = callback;
      $('password-prompt').textContent = reason === pdfjs.PasswordResponses.INCORRECT_PASSWORD ? 'That password did not work. Please try again.' : 'Enter its opening password to continue.';
      $('pdf-password').value = ''; if (!$('password-dialog').open) $('password-dialog').showModal(); $('pdf-password').focus();
    };
    const opened = await loadingTask.promise;
    if (revision !== loadRevision) { opened.destroy(); return; }
    pdf = opened;
    const permissions = await opened.getPermissions();
    if (revision !== loadRevision) return;
    if (permissions && !permissions.has(pdfjs.PermissionFlag.COPY)) {
      reset(); error('This PDF does not permit copying its content. Ask the owner for an unrestricted copy.'); return;
    }
    $('file-meta').textContent = `${opened.numPages} ${opened.numPages === 1 ? 'page' : 'pages'} · ${(chosen.size / 1024 / 1024).toFixed(2)} MB · Ready to convert`;
    $('preview-count').textContent = `${opened.numPages} ${opened.numPages === 1 ? 'page' : 'pages'}`;
    $('preview-navigation').hidden = false; $('convert-button').disabled = false;
    say(`${chosen.name} loaded, ${opened.numPages} pages.`); await showPage(1);
  } catch (e) {
    if (revision !== loadRevision) return;
    reset(); error(e.name === 'InvalidPDFException' ? 'We could not read this PDF. It may be damaged or incomplete.' : e.message || 'Unable to open this PDF. Try another file.');
  }
}

function setBusy(busy) {
  $('remove-file').disabled = busy; $('page-range').disabled = busy; $('convert-button').disabled = busy || !pdf;
  document.querySelectorAll('input[name="mode"]').forEach(input => { input.disabled = busy; });
  $('conversion-progress').hidden = !busy;
  $('convert-button').querySelector('span').textContent = busy ? 'Creating your Word document…' : 'Convert to Word';
}
function download() {
  if (!downloadingUrl) return;
  const link = document.createElement('a'); link.href = downloadingUrl; link.download = resultName; document.body.append(link); link.click(); link.remove();
}
async function convert() {
  if (!pdf || abortController) return;
  error(''); clearResult();
  let pages;
  try { pages = parsePages($('page-range').value, pdf.numPages); } catch (e) { error(e.message); $('page-range').focus(); return; }
  abortController = new AbortController(); setBusy(true);
  try {
    const mode = document.querySelector('input[name="mode"]:checked').value;
    const blob = await convertPdf(pdf, pages, mode, (percent, message) => { $('progress-bar').value = percent; $('progress-message').textContent = message; say(message); }, abortController.signal);
    resultName = outputName(file.name); downloadingUrl = URL.createObjectURL(blob);
    const result = $('conversion-result'); result.replaceChildren();
    const line = document.createElement('div'); line.textContent = `Ready! ${pages.length} ${pages.length === 1 ? 'page' : 'pages'} converted${mode === 'appearance' ? ' as images' : ' to editable text'}.`;
    const again = document.createElement('button'); again.className = 'text-button'; again.textContent = 'Download Word document'; again.addEventListener('click', download);
    result.append(line, again); result.hidden = false; say('Conversion complete. Your Word document is ready.'); download();
  } catch (e) { if (e.name === 'AbortError') { say('Conversion cancelled.'); error('Conversion cancelled. Your PDF is still ready when you are.'); } else error(e.message || 'Conversion failed. Try fewer pages or exact appearance mode.'); }
  finally { abortController = null; setBusy(false); }
}
$('file-input').addEventListener('change', e => openFile(e.target.files[0]));
$('file-input').addEventListener('focus', () => $('dropzone').classList.add('dragging'));
$('file-input').addEventListener('blur', () => $('dropzone').classList.remove('dragging'));
$('remove-file').addEventListener('click', () => { reset(); $('file-input').focus(); });
$('convert-button').addEventListener('click', convert);
$('cancel-button').addEventListener('click', () => { abortController?.abort(); $('progress-message').textContent = 'Cancelling after the current page…'; });
$('previous-page').addEventListener('click', () => showPage(Math.max(1, previewPage - 1)));
$('next-page').addEventListener('click', () => showPage(Math.min(pdf.numPages, previewPage + 1)));
document.querySelectorAll('input[name="mode"]').forEach(input => input.addEventListener('change', () => {
  clearResult(); $('mode-note').textContent = input.value === 'appearance' ? 'Keeps the page design, images, and scanned content. Words will not be editable.' : 'Keeps text, font sizes, and page breaks. Complex columns, tables, and images may need adjustment.';
}));
$('page-range').addEventListener('input', clearResult);
for (const event of ['dragenter', 'dragover']) $('dropzone').addEventListener(event, e => { e.preventDefault(); $('dropzone').classList.add('dragging'); });
for (const event of ['dragleave', 'drop']) $('dropzone').addEventListener(event, e => { e.preventDefault(); $('dropzone').classList.remove('dragging'); });
$('dropzone').addEventListener('drop', e => { if (e.dataTransfer.files.length !== 1) { error('Choose one PDF at a time.'); return; } openFile(e.dataTransfer.files[0]); });
// Prevent the browser from navigating away when a file lands outside the drop area.
window.addEventListener('dragover', e => e.preventDefault()); window.addEventListener('drop', e => e.preventDefault());
$('sample-button').addEventListener('click', async () => {
  $('sample-button').disabled = true;
  try { const response = await fetch('/sample.pdf'); if (!response.ok) throw new Error('Sample unavailable. Please choose your own PDF.'); await openFile(new File([await response.blob()], 'A-fresh-start.pdf', { type: 'application/pdf' })); }
  catch (e) { error(e.message); } finally { $('sample-button').disabled = false; }
});
$('password-form').addEventListener('submit', e => { e.preventDefault(); const callback = passwordCallback; passwordCallback = null; const password = $('pdf-password').value; $('pdf-password').value = ''; $('password-dialog').close(); callback?.(password); });
function cancelPassword() { reset(); say('Opening PDF cancelled.'); }
$('password-cancel').addEventListener('click', cancelPassword);
$('password-dialog').addEventListener('cancel', e => { e.preventDefault(); cancelPassword(); });
window.addEventListener('beforeunload', () => { if (downloadingUrl) URL.revokeObjectURL(downloadingUrl); });
