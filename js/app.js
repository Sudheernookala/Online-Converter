// Word (.docx) to high-resolution PDF, fully in the browser.
// Libraries: docx-preview (window.docx) lays out the document, modern-screenshot
// captures each page as a high-resolution image, jsPDF (window.jspdf) builds the PDF.
import { domToCanvas } from '../vendor/modern-screenshot.mjs?v=4';

// Pages are captured at 300 DPI (print quality).
const DPI = 300;
const JPEG_QUALITY = 0.95;
// Browsers lay out at 96 px per inch; PDF uses 72 points per inch.
const CSS_DPI = 96;
const PX_TO_PT = 72 / 96;
// Stay under the canvas size limit of iOS Safari (~16.7M pixels).
const MAX_CANVAS_PIXELS = 16_000_000;

const $ = (id) => document.getElementById(id);
const drop = $('drop');
const input = $('file');
const stage = $('stage');

let currentUrl = null;
let busy = false;

// ---------- Page ----------

function show(el, visible) { el.hidden = !visible; }

function setProgress(msg, ratio) {
  $('status-text').textContent = msg;
  $('bar-fill').style.width = `${Math.round(ratio * 100)}%`;
}

function formatSize(bytes) {
  return bytes > 1024 * 1024 ? `${(bytes / 1024 / 1024).toFixed(1)} MB` : `${Math.round(bytes / 1024)} KB`;
}

async function handleFile(file) {
  if (!file || busy) return;
  show($('result'), false);
  show($('error'), false);

  const name = file.name.toLowerCase();
  if (!name.endsWith('.docx')) {
    showError(
      name.endsWith('.doc')
        ? 'Old .doc files are not supported. Open the file in Word, use "Save As" → "Word Document (.docx)", then try again.'
        : 'Please choose a Word file (.docx).'
    );
    return;
  }

  busy = true;
  drop.classList.add('busy');
  show($('status'), true);
  setProgress('Reading file…', 0.02);
  if (currentUrl) URL.revokeObjectURL(currentUrl);

  try {
    const { blob, pageCount, brokenImages } = await convertDocxToPdf(file);
    const pdfName = file.name.replace(/\.docx$/i, '') + '.pdf';
    currentUrl = URL.createObjectURL(blob);
    const link = $('download');
    link.href = currentUrl;
    link.download = pdfName;
    $('result-text').textContent =
      `Done: ${pdfName} (${pageCount} page${pageCount === 1 ? '' : 's'}, ${formatSize(blob.size)}). The download has started.`;
    const warning = $('warning');
    warning.textContent = brokenImages
      ? `${brokenImages} image${brokenImages === 1 ? '' : 's'} could not be shown. They are in a format browsers cannot display ` +
        '(usually EMF/WMF). In Word, right-click the image → "Change Picture" or save it as PNG/JPEG, then convert again.'
      : '';
    show(warning, brokenImages > 0);
    show($('status'), false);
    show($('result'), true);
    link.click(); // start the download right away
  } catch (err) {
    console.error(err);
    show($('status'), false);
    showError(`Could not convert this file: ${err.message || err}. The file may be damaged or password-protected.`);
  } finally {
    busy = false;
    drop.classList.remove('busy');
    input.value = '';
    stage.innerHTML = '';
  }
}

function showError(msg) {
  $('error').textContent = msg;
  show($('error'), true);
}

input.addEventListener('change', () => handleFile(input.files[0]));
drop.addEventListener('keydown', (e) => {
  if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); input.click(); }
});
['dragenter', 'dragover'].forEach((t) =>
  drop.addEventListener(t, (e) => { e.preventDefault(); drop.classList.add('over'); })
);
['dragleave', 'drop'].forEach((t) =>
  drop.addEventListener(t, (e) => { e.preventDefault(); drop.classList.remove('over'); })
);
drop.addEventListener('drop', (e) => handleFile(e.dataTransfer.files[0]));
$('again').addEventListener('click', () => input.click());

// ---------- Conversion ----------

async function convertDocxToPdf(file) {
  const buffer = await file.arrayBuffer();
  stage.innerHTML = '';
  const styles = document.createElement('div');
  const body = document.createElement('div');
  stage.append(styles, body);

  setProgress('Laying out pages…', 0.05);
  await window.docx.renderAsync(buffer, body, styles, {
    className: 'docx',
    inWrapper: true,
    breakPages: true,
    // Use the page breaks Word recorded when it last saved the file,
    // so pages match the original as closely as possible.
    ignoreLastRenderedPageBreak: false,
    ignoreWidth: false,
    ignoreHeight: false,
    renderHeaders: true,
    renderFooters: true,
    renderFootnotes: true,
    renderEndnotes: true,
    // Images are embedded as data URLs, so they are captured at full quality.
    useBase64URL: true,
    experimental: true,
  });

  const info = await readDocxInfo(buffer);
  fixLibraryStyles(styles);
  fixTables(body);
  applyContextualSpacing(body, info.contextualStyles);
  setProgress('Loading fonts…', 0.07);
  await loadUsedFonts(body);
  applyWordLineSpacing(body);
  const brokenImages = await waitForImages(stage);

  // Word stores its own page count in the file. When the rendered pages line
  // up with Word's page breaks, keep exactly those pages (a page that runs a
  // little long is shrunk to fit). Otherwise split long pages at line breaks.
  const sections = [...body.querySelectorAll('section.docx')];
  const pages = info.wordPages === sections.length
    ? sections.flatMap(fitSection)
    : sections.flatMap(splitTallSection);
  if (!pages.length) throw new Error('the document has no pages');

  const pdf = await buildPdf(pages);
  pdf.setProperties({ title: file.name.replace(/\.docx$/i, '') });
  setProgress('Finishing…', 1);
  return { blob: pdf.output('blob'), pageCount: pdf.getNumberOfPages(), brokenImages };
}

// ---------- Fixes for the Word layout library (docx-preview 0.4.1) ----------

// Word bullets often use private-use characters from the Symbol / Wingdings
// fonts, which only display on Windows. Map them to normal Unicode characters.
const SYMBOL_BULLETS = {
  '': '•', // • Symbol bullet
  '': '▪', // ▪ Wingdings square
  '': '➢', // ➢ Wingdings arrowhead
  '': '✔', // ✔ Wingdings check
  '': '❖', // ❖ Wingdings diamond
  '': '■', // ■ Wingdings black square
  '': '❑', // ❑ Wingdings box
  '': '➔', // ➔ Wingdings arrow
  '': '□', // □
  '': '□', // □
};

function fixLibraryStyles(container) {
  for (const style of container.querySelectorAll('style')) {
    let css = style.textContent;
    // The library writes the default paragraph style as ".docx p, p.docx_normal span",
    // so its font and size never reach the text. Apply it to the text (spans)
    // with low priority, so headings and other styles still win.
    css = css.replace(/\.docx (\w+), (\1\.[\w-]+) (\w+)(\s*\{)/g, '.docx :where($1) $3, $2 $3$4');
    // Word's "multiple" line spacing is a multiple of the font's natural line
    // height, not of the font size. --docx-lh is set per paragraph later.
    css = css.replace(/line-height:\s*([\d.]+)\s*;/g, 'line-height: calc($1 * var(--docx-lh, 1));');
    // Symbol / Wingdings bullets.
    css = css.replace(/[-]/g, (ch) => SYMBOL_BULLETS[ch] || '•');
    css = css.replace(/font-family:\s*['"]?(Symbol|Wingdings[^;'"]*)['"]?\s*;/gi, 'font-family: inherit;');
    style.textContent = css;
  }
}

/** Makes sure every font face used in the document is loaded before measuring. */
async function loadUsedFonts(root) {
  const wanted = new Map();
  for (const el of root.querySelectorAll('span, p')) {
    const text = el.textContent.trim();
    if (!text) continue;
    const cs = getComputedStyle(el);
    const key = `${cs.fontStyle} ${cs.fontWeight} 16px ${cs.fontFamily}`;
    if (!wanted.has(key)) wanted.set(key, text.slice(0, 64));
    if (wanted.size > 200) break;
  }
  await Promise.all([...wanted].map(([font, text]) => document.fonts.load(font, text).catch(() => {})));
  await document.fonts.ready;
}

/** Natural line height of a font, as a multiple of its size (Calibri ≈ 1.22). */
const lineRatioCache = new Map();
function naturalLineRatio(fontFamily) {
  if (!lineRatioCache.has(fontFamily)) {
    const probe = document.createElement('span');
    probe.textContent = 'Hg';
    Object.assign(probe.style, {
      fontFamily, fontSize: '100px', lineHeight: 'normal', display: 'inline-block', position: 'absolute',
    });
    stage.append(probe);
    const ratio = probe.getBoundingClientRect().height / 100;
    probe.remove();
    lineRatioCache.set(fontFamily, ratio > 0.8 && ratio < 2 ? ratio : 1.15);
  }
  return lineRatioCache.get(fontFamily);
}

function applyWordLineSpacing(root) {
  // Direct paragraph formatting is written as inline styles.
  for (const el of root.querySelectorAll('[style*="line-height"]')) {
    const v = el.style.lineHeight;
    if (/^[\d.]+$/.test(v)) el.style.lineHeight = `calc(${v} * var(--docx-lh, 1))`;
  }
  for (const p of root.querySelectorAll('p')) {
    const textEl = p.querySelector('span') || p;
    p.style.setProperty('--docx-lh', naturalLineRatio(getComputedStyle(textEl).fontFamily).toFixed(3));
  }
}

/**
 * Reads details the layout library does not use:
 * - wordPages: the page count Word saved (only trusted when Word also recorded
 *   where its pages break), or null
 * - contextualStyles: CSS classes of paragraph styles with "Don't add space
 *   between paragraphs of the same style" (common for lists)
 */
async function readDocxInfo(buffer) {
  const info = { wordPages: null, contextualStyles: new Set() };
  try {
    const zip = await window.JSZip.loadAsync(buffer);
    const read = (name) => zip.file(name)?.async('string') ?? Promise.resolve('');
    const [app, docXml, stylesXml] = await Promise.all([
      read('docProps/app.xml'), read('word/document.xml'), read('word/styles.xml'),
    ]);
    const pages = app.match(/<(?:\w+:)?Pages>(\d+)</);
    if (pages && docXml.includes('lastRenderedPageBreak')) info.wordPages = Number(pages[1]);

    const doc = new DOMParser().parseFromString(stylesXml, 'application/xml');
    for (const style of doc.getElementsByTagNameNS('*', 'style')) {
      const cs = style.getElementsByTagNameNS('*', 'contextualSpacing')[0];
      if (!cs) continue;
      const val = cs.getAttributeNS(cs.namespaceURI, 'val') ?? cs.getAttribute('w:val');
      if (val === '0' || val === 'false') continue;
      const id = style.getAttributeNS(style.namespaceURI, 'styleId') || style.getAttribute('w:styleId');
      // Same naming rule as the library: docx_<styleId, lowercase>.
      if (id) info.contextualStyles.add('docx_' + id.replace(/[ .]+/g, '-').replace(/[&]+/g, 'and').toLowerCase());
    }
  } catch (err) {
    console.warn('Could not read document details', err);
  }
  return info;
}

/** Word removes the space between paragraphs of the same style when asked to. */
function applyContextualSpacing(root, classes) {
  if (!classes.size) return;
  for (const p of root.querySelectorAll('p')) {
    const cls = [...p.classList].find((c) => classes.has(c));
    if (!cls) continue;
    const next = p.nextElementSibling;
    if (next?.tagName === 'P' && next.classList.contains(cls)) {
      p.style.marginBottom = '0';
      next.style.marginTop = '0';
    }
  }
}

/**
 * Table styles (header row, banded rows, first column…) only apply when the
 * rows and cells carry marker classes. Word usually writes these markers, but
 * not always; add them from the table's settings when they are missing.
 */
function fixTables(root) {
  for (const table of root.querySelectorAll('table')) {
    if (table.querySelector('tr.first-row, tr.odd-row, tr.even-row, td.first-col, td.odd-col')) continue;
    const rows = [...table.rows].filter((r) => r.closest('table') === table);
    if (!rows.length) continue;
    const has = (c) => table.classList.contains(c);
    if (has('first-row')) rows[0].classList.add('first-row');
    if (has('last-row')) rows[rows.length - 1].classList.add('last-row');
    let band = 0;
    rows.forEach((row, ri) => {
      const isHeader = (ri === 0 && has('first-row')) || (ri === rows.length - 1 && has('last-row'));
      if (!isHeader && !has('no-hband')) row.classList.add(band++ % 2 === 0 ? 'odd-row' : 'even-row');
      const cells = [...row.cells];
      if (has('first-col') && cells[0]) cells[0].classList.add('first-col');
      if (has('last-col') && cells.length) cells[cells.length - 1].classList.add('last-col');
      let colBand = 0;
      cells.forEach((cell, ci) => {
        const isEdge = (ci === 0 && has('first-col')) || (ci === cells.length - 1 && has('last-col'));
        if (!isEdge && !has('no-vband')) cell.classList.add(colBand++ % 2 === 0 ? 'odd-col' : 'even-col');
      });
    });
  }
}

// ---------- Pages ----------

/**
 * One rendered section = one Word page. If the content runs a little past the
 * page (small layout differences from Word), shrink it to fit instead of
 * adding a page. Content much longer than a page is still split.
 */
function fitSection(el) {
  const width = el.offsetWidth;
  const totalHeight = el.offsetHeight;
  const pageHeight = parsePx(getComputedStyle(el).minHeight) || totalHeight;
  if (totalHeight <= pageHeight + 2) return [{ el, width, height: pageHeight, slice: null }];
  const shrink = pageHeight / totalHeight;
  if (shrink < 0.8) return splitTallSection(el);
  return [{ el, width, height: pageHeight, slice: null, shrink, fullHeight: totalHeight }];
}

/** Waits for every image to load. Returns how many could not be displayed. */
async function waitForImages(root) {
  const imgs = [...root.querySelectorAll('img')];
  await Promise.all(
    imgs.map((img) =>
      img.complete
        ? img.decode?.().catch(() => {})
        : new Promise((res) => {
            img.addEventListener('load', res, { once: true });
            img.addEventListener('error', res, { once: true });
          })
    )
  );
  const broken = imgs.filter((img) => !img.naturalWidth);
  // Hide broken images so the PDF does not show a broken-image icon.
  broken.forEach((img) => { img.style.visibility = 'hidden'; });
  return broken.length;
}

function parsePx(value) {
  const n = parseFloat(value);
  return Number.isFinite(n) ? n : 0;
}

/** Returns one or more page descriptors for a rendered Word section. */
function splitTallSection(el) {
  const width = el.offsetWidth;
  const totalHeight = el.offsetHeight;
  const cs = getComputedStyle(el);
  // docx-preview sets min-height to the page height (computed style gives px).
  const pageHeight = parsePx(cs.minHeight) || totalHeight;

  if (totalHeight <= pageHeight + 2) {
    return [{ el, width, height: pageHeight, slice: null }];
  }

  const padTop = parsePx(cs.paddingTop);
  const padBottom = parsePx(cs.paddingBottom);
  const sectionTop = el.getBoundingClientRect().top;

  // Safe places to cut: the bottom of every line-level block. Pictures sit
  // inside paragraphs, so a cut never goes through a picture unless the
  // picture alone is taller than the page.
  const cuts = new Set();
  el.querySelectorAll('p, tr, li, img, h1, h2, h3, h4, h5, h6').forEach((node) => {
    const r = node.getBoundingClientRect();
    if (r.height > 0) cuts.add(Math.round(r.bottom - sectionTop));
  });
  const candidates = [...cuts].sort((a, b) => a - b);

  const slices = [];
  let start = 0;
  const contentEnd = totalHeight - padBottom;
  while (start < contentEnd - 1) {
    const first = slices.length === 0;
    // The first page keeps the real top margin; later pages get it added back.
    const room = first ? pageHeight - padBottom : pageHeight - padTop - padBottom;
    const limit = start + room;
    let end = limit >= contentEnd ? contentEnd : 0;
    if (!end) {
      for (const c of candidates) {
        if (c > start + 1 && c <= limit) end = c;
        else if (c > limit) break;
      }
      if (!end) end = limit; // a single block taller than a page: hard cut
    }
    slices.push({ srcY: start, srcH: end - start, destY: first ? 0 : padTop });
    start = end;
  }
  return slices.map((slice) => ({ el, width, height: pageHeight, slice }));
}

function pickScale(width, height) {
  let scale = DPI / CSS_DPI;
  const area = width * scale * height * scale;
  if (area > MAX_CANVAS_PIXELS) scale *= Math.sqrt(MAX_CANVAS_PIXELS / area);
  return scale;
}

async function buildPdf(pages) {
  const { jsPDF } = window.jspdf;
  let pdf = null;

  for (let i = 0; i < pages.length; i++) {
    const page = pages[i];
    setProgress(`Converting page ${i + 1} of ${pages.length}…`, 0.1 + 0.85 * (i / pages.length));

    const scale = pickScale(page.width, page.height);
    let canvas;
    if (page.slice) canvas = await captureSlice(page, scale);
    else if (page.shrink) canvas = await captureShrunk(page, scale);
    else canvas = await domToCanvas(page.el, { scale, backgroundColor: '#ffffff' });

    const wPt = page.width * PX_TO_PT;
    const hPt = page.height * PX_TO_PT;
    const orientation = wPt > hPt ? 'landscape' : 'portrait';
    if (!pdf) pdf = new jsPDF({ unit: 'pt', format: [wPt, hPt], orientation, compress: true });
    else pdf.addPage([wPt, hPt], orientation);
    pdf.addImage(canvas.toDataURL('image/jpeg', JPEG_QUALITY), 'JPEG', 0, 0, wPt, hPt);

    // Free memory before the next page, and let the progress bar repaint.
    canvas.width = canvas.height = 0;
    await new Promise((r) => setTimeout(r, 0));
  }
  return pdf;
}

/**
 * Captures one page-sized part of a tall section. A copy of the section is
 * placed in a page-sized box (same CSS classes, so the same styles apply)
 * and shifted up so only this page's content shows.
 */
async function captureSlice(page, scale) {
  const { srcY, srcH, destY } = page.slice;
  const box = document.createElement('div');
  box.className = page.el.parentElement.className; // keeps docx-wrapper styles
  Object.assign(box.style, {
    position: 'relative', overflow: 'hidden', display: 'block',
    width: `${page.width}px`, height: `${page.height}px`,
    padding: '0', margin: '0', background: '#ffffff',
  });
  const view = document.createElement('div');
  Object.assign(view.style, {
    position: 'absolute', left: '0', top: `${destY}px`,
    width: `${page.width}px`, height: `${srcH}px`, overflow: 'hidden',
  });
  const copy = page.el.cloneNode(true);
  Object.assign(copy.style, { position: 'absolute', left: '0', top: `${-srcY}px`, margin: '0' });
  view.append(copy);
  box.append(view);
  page.el.parentElement.after(box);
  try {
    return await domToCanvas(box, { scale, backgroundColor: '#ffffff' });
  } finally {
    box.remove();
  }
}

/** Captures a slightly-too-long page and scales it down onto one page. */
async function captureShrunk(page, scale) {
  // Captured at the normal scale; drawing it smaller keeps it at least as sharp.
  const full = await domToCanvas(page.el, {
    scale: Math.min(scale, pickScale(page.width, page.fullHeight)),
    backgroundColor: '#ffffff',
  });
  const canvas = document.createElement('canvas');
  canvas.width = Math.round(page.width * scale);
  canvas.height = Math.round(page.height * scale);
  const ctx = canvas.getContext('2d');
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  ctx.imageSmoothingQuality = 'high';
  const w = canvas.width * page.shrink;
  ctx.drawImage(full, (canvas.width - w) / 2, 0, w, canvas.height);
  full.width = full.height = 0;
  return canvas;
}

window.__converterReady = true;
