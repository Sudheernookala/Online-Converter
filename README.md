# Word to PDF

A web page that converts a Word file (**.docx**) to a high-resolution PDF.
Choose a file, and the PDF downloads automatically. Everything runs in the browser; the file never leaves your device.

Live site: https://sudheernookala.github.io/Online-Converter/

## How it works

1. [docx-preview](https://github.com/VolodymyrBaydalka/docxjs) lays out the Word document in the browser.
2. Each page is captured at **300 DPI** (print quality) with [modern-screenshot](https://github.com/qq15725/modern-screenshot). A Letter page becomes 2550 × 3300 pixels, so pictures stay sharp.
3. [jsPDF](https://github.com/parallax/jsPDF) puts the pages into a PDF at the original page size.

Pictures are never cut in half at a page break; a picture that does not fit moves to the next page.

## No build step

The site is plain HTML, CSS and JavaScript. The libraries are copied into `vendor/`.
GitHub Pages can publish it with either setting ("Deploy from a branch" or "GitHub Actions").

Run locally:

```bash
python3 -m http.server 8000
# open http://localhost:8000
```

## Limits

- **Text in the PDF is part of the page image**, so it cannot be selected or searched.
- **Layout is close to Word, but not identical.** This is a browser renderer, not Microsoft Word. Text boxes, SmartArt, charts and complex columns may look different.
- **Fonts:** if a font in the document is not installed on the device, a similar one is used.
- **EMF/WMF pictures** (common for pasted charts and clip art) cannot be shown by browsers. The page warns when this happens.
- Old **.doc** files are not supported. Save them as .docx first.
