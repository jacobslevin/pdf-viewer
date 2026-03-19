# Assisted Spec Capture POC

Static proof of concept for Phase 1 of assisted spec capture.

## What it demonstrates

- Split-screen workflow with source content on the left and editable attributes on the right
- Page targeting based on product, brand, category, and live attribute labels
- Text selection capture in the left-side source viewer
- Contextual `Send to field` action
- Field picker with replace/append handling for pre-filled fields
- Visual confirmation via focus, highlight, and toast

## Run

Serve the folder with any static server and open the app in the browser.

Example:

```bash
python3 -m http.server 4173
```

Then open [http://localhost:4173](http://localhost:4173).

For `Analyze dimensions`, paste an OpenAI API key into the browser field. The key is kept only in browser memory for the current session.

## Important prototype shortcut

This prototype parses page text from deliberately simple ASCII PDF source strings embedded in `app.js` and renders a PDF-style text viewer, rather than using a full PDF.js-based renderer.

That means:

- page-by-page text extraction is real for the included sample PDFs
- selection-to-field flow is real
- auto-targeting is real
- visual PDF rendering is approximated with page cards for speed and zero dependencies

For production, the viewer should be replaced with PDF.js or the current app viewer if it can expose text-layer selection and page navigation hooks.
