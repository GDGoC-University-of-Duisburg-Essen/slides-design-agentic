# Agent guide: building GDG UDE slide decks

This repository is the slide design system of **GDG on Campus University of Duisburg-Essen**. If a user asks you to build a presentation "in the GDG UDE design" or "based on `https://github.com/GDGoC-University-of-Duisburg-Essen/slides-design-agentic`", follow this guide.

Read these before you write slides:

1. [`docs/content-guidelines.md`](docs/content-guidelines.md): how much text goes on a slide, and when to warn the user. **Mandatory.**
2. [`docs/slide-types.md`](docs/slide-types.md): available slide types with HTML markup, Python calls and JSON spec.
3. [`docs/design-rules.md`](docs/design-rules.md): colours, fonts, sizes and layout (only needed for custom content such as diagrams).

## 0. Get the repository

If you are not inside a clone already:

```bash
git clone --depth 1 https://github.com/GDGoC-University-of-Duisburg-Essen/slides-design-agentic.git gdg-slides
cd gdg-slides
npm install && npx playwright install chromium        # for HTML → PDF/PNG
python3 -m pip install -r requirements.txt            # for PowerPoint (python-pptx)
```

LibreOffice (`soffice`) is optional. It turns PPTX into PDF and is needed to rebuild the `.otp` templates.

## 1. Collect the inputs

You need the **title**, **speaker name(s)**, **date**, **language** (`en` or `de`), **theme** (`light` by default, `dark` on request), **output format** (HTML, PDF and/or PPTX) and the **content**. Ask only for what is missing and cannot be inferred. Use today's date if none is given.

Decide which content case applies (see the content guidelines):

- **Exact text given:** use it verbatim, build it anyway, and warn about slides that break the rules.
- **Only topics or an outline:** write assertion titles, 3–5 short bullets or a visual, and speaker notes. Do not invent facts. Mark gaps as `TODO`.

## 2. Build

| Requested output | Route |
|---|---|
| HTML (to present in a browser) | HTML deck (`html/gdg-slides.css` + `gdg-slides.js`) |
| PDF | HTML deck → `tools/render.mjs` (best quality). Alternatively PPTX → LibreOffice |
| PowerPoint / editable deck | `tools/gdg_pptx.py` (uses `templates/gdg-ude-*.potx`) |
| Several formats | Build each route from the same outline; keep slide order and text identical |

### HTML / PDF

1. Copy [`html/starter.html`](html/starter.html) to the deck location (e.g. `decks/<slug>/index.html`), or start from [`examples/example-deck.html`](examples/example-deck.html).
2. Point the `<link>`/`<script>` to `html/gdg-slides.css` and `html/gdg-slides.js`, either with a relative path into this repo or via the CDN: `https://cdn.jsdelivr.net/gh/GDGoC-University-of-Duisburg-Essen/slides-design-agentic@main/html/…`. Put images next to the deck.
3. Fill `<main class="deck" data-title data-author data-date data-lang>`. Write one `<section class="slide slide--TYPE">` per slide. **Do not** write footers, slide numbers, logos or decorations: the script adds them.
4. Render, lint and look at the result:
   ```bash
   node tools/render.mjs decks/<slug>/index.html --both --png decks/<slug>/preview --scale 0.5
   ```
   This writes `index-light.pdf` + `index-dark.pdf` (drop `--both` for light only, or use `--theme dark`) and prints the linter warnings.

### PowerPoint (PPTX)

Write a JSON spec (format in [`docs/slide-types.md`](docs/slide-types.md#json-deck-spec-for-toolsgdg_pptxpy), example [`examples/example-deck.json`](examples/example-deck.json)) and run:

```bash
python3 tools/gdg_pptx.py decks/<slug>/deck.json decks/<slug>/deck.pptx --pdf
```

Or use the Python API (`from gdg_pptx import Deck`, see the docstring in [`tools/gdg_pptx.py`](tools/gdg_pptx.py)) when you need custom shapes on a `title_only()` canvas. Both routes print linter warnings. `--pdf` also exports a PDF via LibreOffice, which you can rasterise to check the layout visually.

## 3. Verify (always)

1. **Linter:** fix every overflow or cut-off warning. Text running out of the slide is never acceptable. Fix text-amount warnings in case B. In case A, report them to the user.
2. **Look at the slides:** open the PNG previews (or the PDF) and check overlaps, awkward line breaks, empty slides, low-contrast images and wrong slide types.
3. **Facts:** every number, quote and claim must come from the user or from a source on the references slide.

## 4. Hard rules

- Use only the slide types in `docs/slide-types.md`. Never add your own backgrounds, gradients, logos, fonts or colours outside the palette.
- Use Google Sans and Google Sans Code only. Body text is 24 pt (48 px), slide titles 36 pt (72 px). Content text is never smaller than 16 pt (32 px), and you never shrink text to make it fit: split the slide instead.
- Every base slide has a title. Use one message per slide, with an assertion-style title of max. 2 lines.
- Use ≤ 6 bullets of ≤ 1–2 lines and ≤ 2 list levels. Details go into speaker notes.
- Use the slide order: title → (agenda) → chapters with content → key takeaways → references (if sources) → questions → next event (if the user mentions an event).
- Keep the user's language. German decks: `data-lang="de"` / `"lang": "de"`.
- Images need `alt` text (HTML). Diagrams use the brand colours and ≥ 16 pt text.

## 5. Final message to the user

Report briefly: the files created (paths), the theme, the number of slides, **all remaining linter warnings** (as short ⚠ lines), any `TODO` placeholders, and how to present it (HTML: open in a browser and use the arrow keys; F = fullscreen, T = light/dark).

## Maintaining this repository

- Design values: `design/tokens.json` (the HTML CSS and the content rules in `html/gdg-slides.js` mirror it; keep them in sync).
- Rebuild templates after design changes: `python3 tools/build_templates.py` (writes `templates/*.potx`, plus `*.otp` if LibreOffice is installed).
- Rebuild examples: `npm run examples`, and `python3 tools/gdg_pptx.py examples/example-deck.json examples/example-deck.pptx`.
