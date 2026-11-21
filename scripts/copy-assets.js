import { cpSync, mkdirSync } from 'node:fs';
for (const directory of ['cmaps', 'standard_fonts', 'wasm']) {
  mkdirSync(`public/pdf-assets/${directory}`, { recursive: true });
  cpSync(`node_modules/pdfjs-dist/${directory}`, `public/pdf-assets/${directory}`, { recursive: true });
}
for (const [source, target] of [
  ['node_modules/pdfjs-dist/LICENSE', 'public/pdf-assets/PDFJS-LICENSE.txt'],
  ['node_modules/docx/LICENSE', 'public/pdf-assets/DOCX-LICENSE.txt']
]) cpSync(source, target);
