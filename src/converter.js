import { domToCanvas } from 'modern-screenshot';
import { jsPDF } from 'jspdf';

// Browsers lay out at 96 px per inch; PDF uses 72 points per inch.
const CSS_DPI = 96;
const PX_TO_PT = 72 / 96;
// Stay under the canvas area limit of iOS Safari (~16.7M pixels).
const MAX_CANVAS_PIXELS = 16_000_000;

export class UnsupportedFileError extends Error {}

export function detectType(file) {
  const name = file.name.toLowerCase();
  if (name.endsWith('.docx')) return 'docx';
  if (name.endsWith('.pptx')) return 'pptx';
  if (name.endsWith('.doc') || name.endsWith('.ppt')) {
    throw new UnsupportedFileError(
      'Old .doc / .ppt files are not supported. Open the file in Word or PowerPoint, ' +
        'use "Save As" to save it as .docx or .pptx, then upload it again.'
    );
  }
  throw new UnsupportedFileError('Please upload a Word (.docx) or PowerPoint (.pptx) file.');
}

/**
 * Convert a .docx or .pptx File to a PDF Blob.
 * @param {File} file
 * @param {{ dpi: number, imageFormat: 'jpeg' | 'png', stage: HTMLElement, onProgress?: (msg: string, ratio: number) => void }} opts
 */
export async function convertToPdf(file, { dpi, imageFormat, stage, onProgress = () => {} }) {
  const type = detectType(file);
  const buffer = await file.arrayBuffer();
  stage.innerHTML = '';

  try {
    onProgress('Reading file…', 0.05);
    const pages = type === 'docx' ? await renderDocx(buffer, stage) : await renderPptx(buffer, stage);
    if (!pages.length) throw new Error('The file has no pages to convert.');

    await document.fonts.ready;
    await waitForImages(stage);

    const pdf = await buildPdf(pages, { dpi, imageFormat, type, onProgress });
    pdf.setProperties({ title: file.name.replace(/\.(docx|pptx)$/i, ''), creator: 'Online Converter' });
    onProgress('Finishing…', 1);
    return { blob: pdf.output('blob'), pageCount: pdf.getNumberOfPages() };
  } finally {
    stage.innerHTML = '';
  }
}

// ---------- Word ----------

async function renderDocx(buffer, stage) {
  // Loaded on demand so the page itself opens fast.
  const { renderAsync } = await import('docx-preview');
  const body = document.createElement('div');
  const styles = document.createElement('div');
  stage.append(styles, body);

  await renderAsync(buffer, body, styles, {
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
    useBase64URL: true,
    experimental: true,
  });

  const sections = [...body.querySelectorAll('section.docx')];
  // Each rendered section is one page, unless its content ran past the page
  // height (no page breaks in the file). Those are split into several pages.
  return sections.flatMap((el) => splitTallSection(el));
}

function parsePx(value) {
  const n = parseFloat(value);
  return Number.isFinite(n) ? n : 0;
}

/** Returns a list of page descriptors for a docx section. */
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

  // Safe places to cut: the bottom of every line-level block.
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
    // First page keeps the real top margin; later pages get it added back.
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

// ---------- PowerPoint ----------

async function renderPptx(buffer, stage) {
  const { init: initPptx } = await import('pptx-preview');
  const host = document.createElement('div');
  stage.append(host);
  const RENDER_WIDTH = 960;
  const previewer = initPptx(host, { width: RENDER_WIDTH, mode: 'list' });
  await previewer.preview(buffer);

  // Slides are laid out at 960px wide. The real slide size comes from the file
  // (pptx-preview reports it in points) and sets the PDF page size and scale.
  const realWidth = previewer.pptx?.width ? previewer.pptx.width / PX_TO_PT : RENDER_WIDTH;
  const zoom = realWidth / RENDER_WIDTH;
  const slides = [...host.querySelectorAll('.pptx-preview-slide-wrapper')];
  return slides.map((el) => ({ el, width: el.offsetWidth, height: el.offsetHeight, zoom, slice: null }));
}

// ---------- PDF ----------

function waitForImages(root) {
  const imgs = [...root.querySelectorAll('img')];
  return Promise.all(
    imgs.map((img) =>
      img.complete && img.naturalWidth
        ? img.decode?.().catch(() => {})
        : new Promise((res) => {
            img.addEventListener('load', res, { once: true });
            img.addEventListener('error', res, { once: true });
          })
    )
  );
}

function pickScale(width, height, dpi) {
  let scale = dpi / CSS_DPI;
  const area = width * scale * height * scale;
  if (area > MAX_CANVAS_PIXELS) scale *= Math.sqrt(MAX_CANVAS_PIXELS / area);
  return scale;
}

async function buildPdf(pages, { dpi, imageFormat, type, onProgress }) {
  let pdf = null;

  for (let i = 0; i < pages.length; i++) {
    const page = pages[i];
    onProgress(`Rendering ${type === 'pptx' ? 'slide' : 'page'} ${i + 1} of ${pages.length}…`, 0.1 + 0.85 * (i / pages.length));

    // `zoom` maps layout px to real-size px (only differs for slides).
    const zoom = page.zoom || 1;
    const scale = pickScale(page.width * zoom, page.height * zoom, dpi) * zoom;
    let canvas;

    if (!page.slice) {
      canvas = await domToCanvas(page.el, { scale, backgroundColor: '#ffffff' });
    } else {
      canvas = await captureSlice(page, scale);
    }

    const wPt = page.width * zoom * PX_TO_PT;
    const hPt = page.height * zoom * PX_TO_PT;
    const orientation = wPt > hPt ? 'landscape' : 'portrait';
    if (!pdf) {
      pdf = new jsPDF({ unit: 'pt', format: [wPt, hPt], orientation, compress: true });
    } else {
      pdf.addPage([wPt, hPt], orientation);
    }

    if (imageFormat === 'png') {
      pdf.addImage(canvas, 'PNG', 0, 0, wPt, hPt, undefined, 'FAST');
    } else {
      pdf.addImage(canvas.toDataURL('image/jpeg', 0.95), 'JPEG', 0, 0, wPt, hPt);
    }

    // Free memory before the next page.
    canvas.width = canvas.height = 0;
    // Yield so the progress bar can repaint.
    await new Promise((r) => setTimeout(r, 0));
  }

  return pdf;
}

/**
 * Capture one page-sized part of a tall Word section. A copy of the section is
 * placed inside a page-sized box (same CSS classes, so the same styles apply)
 * and shifted up so only this page's content is visible.
 */
async function captureSlice(page, scale) {
  const { srcY, srcH, destY } = page.slice;
  const box = document.createElement('div');
  box.className = page.el.parentElement.className; // keeps docx-wrapper styles
  Object.assign(box.style, {
    position: 'relative',
    overflow: 'hidden',
    width: `${page.width}px`,
    height: `${page.height}px`,
    padding: '0',
    margin: '0',
    background: '#ffffff',
    display: 'block',
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
