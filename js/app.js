// Word (.docx) to high-resolution PDF, fully in the browser.
// Libraries: docx-preview (window.docx) lays out the document, modern-screenshot
// captures each page as a high-resolution image, jsPDF (window.jspdf) builds the PDF.
import { domToCanvas } from '../vendor/modern-screenshot.mjs';

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

  await document.fonts.ready;
  const brokenImages = await waitForImages(stage);

  // Each rendered section is one page, unless its content ran past the page
  // height (no page breaks in the file). Those are split into several pages.
  const pages = [...body.querySelectorAll('section.docx')].flatMap(splitTallSection);
  if (!pages.length) throw new Error('the document has no pages');

  const pdf = await buildPdf(pages);
  pdf.setProperties({ title: file.name.replace(/\.docx$/i, '') });
  setProgress('Finishing…', 1);
  return { blob: pdf.output('blob'), pageCount: pdf.getNumberOfPages(), brokenImages };
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
    const canvas = page.slice
      ? await captureSlice(page, scale)
      : await domToCanvas(page.el, { scale, backgroundColor: '#ffffff' });

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

window.__converterReady = true;
