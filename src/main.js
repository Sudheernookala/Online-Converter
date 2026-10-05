import { convertToPdf, UnsupportedFileError } from './converter.js';

const $ = (id) => document.getElementById(id);
const drop = $('drop');
const input = $('file');
const status = $('status');
const statusText = $('status-text');
const barFill = $('bar-fill');
const result = $('result');
const resultText = $('result-text');
const download = $('download');
const openLink = $('open');
const errorBox = $('error');

let currentUrl = null;
let busy = false;

function show(el, visible) { el.hidden = !visible; }

function formatSize(bytes) {
  return bytes > 1024 * 1024 ? `${(bytes / 1024 / 1024).toFixed(1)} MB` : `${Math.round(bytes / 1024)} KB`;
}

async function handleFile(file) {
  if (!file || busy) return;
  busy = true;
  drop.classList.add('busy');
  show(result, false);
  show(errorBox, false);
  show(status, true);
  if (currentUrl) URL.revokeObjectURL(currentUrl);

  const started = performance.now();
  try {
    const { blob, pageCount } = await convertToPdf(file, {
      dpi: Number($('dpi').value),
      imageFormat: $('format').value,
      stage: $('stage'),
      onProgress: (msg, ratio) => {
        statusText.textContent = msg;
        barFill.style.width = `${Math.round(ratio * 100)}%`;
      },
    });

    const name = file.name.replace(/\.(docx|pptx)$/i, '') + '.pdf';
    currentUrl = URL.createObjectURL(blob);
    download.href = currentUrl;
    download.download = name;
    openLink.href = currentUrl;
    const secs = ((performance.now() - started) / 1000).toFixed(1);
    resultText.textContent = `Done: ${name} – ${pageCount} page${pageCount === 1 ? '' : 's'}, ${formatSize(blob.size)} (${secs}s).`;
    show(status, false);
    show(result, true);
    download.click(); // start the download right away
  } catch (err) {
    console.error(err);
    show(status, false);
    errorBox.textContent =
      err instanceof UnsupportedFileError
        ? err.message
        : `Could not convert this file: ${err.message || err}. The file may be damaged or use features that are not supported.`;
    show(errorBox, true);
  } finally {
    busy = false;
    drop.classList.remove('busy');
    input.value = '';
  }
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
$('again').addEventListener('click', () => { show(result, false); input.click(); });

window.__converterReady = true;
