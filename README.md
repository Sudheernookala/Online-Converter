# Word & PowerPoint to PDF

A static web page that converts **.docx** and **.pptx** files to high-resolution PDF.
All conversion runs in the browser, so it works on GitHub Pages and files never leave the user's device.

## How it works

1. The file is laid out in the browser ([docx-preview](https://github.com/VolodymyrBaydalka/docxjs) for Word, [pptx-preview](https://www.npmjs.com/package/pptx-preview) for PowerPoint).
2. Each page or slide is captured at the chosen resolution (200 / 300 / 450 DPI) with [modern-screenshot](https://github.com/qq15725/modern-screenshot).
3. The pages are put into a PDF with [jsPDF](https://github.com/parallax/jsPDF), at the original page or slide size. The PDF downloads straight away.

Images keep their full detail: at 300 DPI an A4/Letter page is about 2550 × 3300 pixels and a 16:9 slide is 4000 × 2250 pixels.

## Limits

- **Text in the PDF is an image**, so it cannot be selected or searched.
- **Layout is close to the original, but not exact.** This is a browser renderer, not Microsoft Office. Complex layouts (text boxes, SmartArt, charts, some animations or effects) may look different.
- **Fonts:** if a font used in the file is not installed on the device, a similar font is used instead.
- **Word page breaks:** files saved by Word keep Word's page breaks. Other files are split into pages at line boundaries, and in that case headers and footers appear only on the first page of each section.
- Old binary **.doc / .ppt** files are not supported. Save them as .docx / .pptx first.
- Large files use a lot of memory, mostly on phones. Choose a lower resolution if a conversion fails.

## Run locally

```bash
npm install
npm run dev      # development server
npm run build    # static site in dist/
```

## Deploy to GitHub Pages

1. In the repository, go to **Settings → Pages** and set **Source** to **GitHub Actions**.
2. Push to `main` (or run the "Deploy to GitHub Pages" workflow by hand). The site is published at `https://<user>.github.io/<repo>/`.
