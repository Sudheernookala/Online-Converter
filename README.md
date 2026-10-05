# Word to PDF

A web page that converts a Word file (**.docx**) to a high-resolution PDF.
Choose a file, and the PDF downloads automatically. Everything runs in the browser; the file never leaves your device.

Live site: https://sudheernookala.github.io/Online-Converter/

## How it works

1. [docx-preview](https://github.com/VolodymyrBaydalka/docxjs) lays out the Word document in the browser.
2. Each page is captured at **300 DPI** (print quality) with [modern-screenshot](https://github.com/qq15725/modern-screenshot).
3. **Pictures are placed into the PDF as their original files**, on top of the page at their exact position. JPEG photos are copied byte for byte (no re-compression), PNG pictures stay lossless with transparency, and a picture used twice is stored once. So pictures keep their full resolution, the same as Word's own "Save as PDF". Cropped pictures keep the original file and are clipped in the PDF.
   A picture stays part of the page image only when it is rotated, faded, overlapped by text or other pictures, or in a format other than JPEG/PNG.
4. [jsPDF](https://github.com/parallax/jsPDF) builds the PDF at the original page size.

Pictures are never cut in half at a page break; a picture that does not fit moves to the next page.

### Keeping Word's formatting and page count

- **Same page count as Word:** Word saves its page count and page breaks in the file. When they are present, the PDF keeps exactly those pages. A page that comes out slightly too long is shrunk a little to fit, instead of spilling onto an extra page.
- **Fonts:** `fonts/` holds free fonts with exactly the same letter widths as Calibri, Cambria, Times New Roman, Arial and Courier New (Carlito, Caladea, Tinos, Arimo, Cousine; SIL Open Font License). So text wraps like in Word even on devices without Microsoft fonts. An installed copy of the real font is always used first.
- **Fixes for the layout library** (`js/app.js`): the document's main font and size (the "Normal" style) are applied to the text, Word's line spacing is calculated the way Word does it, Symbol/Wingdings bullets are shown as •, table header rows and banded rows get their style, and list items have no extra space between them when Word says so.

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
- **Fonts:** fonts other than the five above (for example Aptos, the newer Office default) are only exact if installed on the device; otherwise a similar font is used.
- **EMF/WMF pictures** (common for pasted charts and clip art) cannot be shown by browsers. The page warns when this happens.
- Old **.doc** files are not supported. Save them as .docx first.
