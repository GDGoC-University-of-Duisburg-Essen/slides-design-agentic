# GDG UDE Slide Design

The slide design system of **Google Developer Group on Campus – University of Duisburg-Essen**. It contains design rules, assets, templates for PowerPoint and LibreOffice Impress, an HTML slide framework, and tools that let AI agents (Claude, Claude Code, …) build complete decks as **PDF, PowerPoint or HTML**.

| Title | Default | Next event |
|---|---|---|
| ![Title slide](examples/screenshots/example-deck-light-01-title.png) | ![Default slide](examples/screenshots/example-deck-light-04-default.png) | ![Next event slide](examples/screenshots/example-deck-light-17-next-event.png) |
| ![Title slide, dark](examples/screenshots/example-deck-dark-01-title.png) | ![Default slide, dark](examples/screenshots/example-deck-dark-04-default.png) | ![Next event slide, dark](examples/screenshots/example-deck-dark-17-next-event.png) |

Full example deck: [light PDF](examples/example-deck-light.pdf) · [dark PDF](examples/example-deck-dark.pdf) · [PowerPoint](examples/example-deck.pptx) · [HTML source](examples/example-deck.html)

## How to use it

### Option 1: Ask an agent

Give Claude, Claude Code or another agent your content and point it at this repository:

> Build a slide deck for my GDG UDE talk "Intro to the Gemini API" (speaker: Jane Doe, 4 December 2026) as **PDF in light mode**, based on the design from `https://github.com/GDGoC-University-of-Duisburg-Essen/slides-design-agentic`. Content: …

The agent follows [`AGENTS.md`](AGENTS.md). It clones the repo, picks the right slide types, keeps the text amount within the [content guidelines](docs/content-guidelines.md), renders the deck, checks it with the built-in linter and reports problems.

- **Exact text given** → the agent uses it verbatim and only *warns* when a slide has too much (or too little) text.
- **Only an outline or keywords** → the agent writes assertion-style titles, short bullets and speaker notes itself.
- Ask for **PowerPoint** if you want to keep editing the slides, for **PDF** to share them, or for **HTML** to present from the browser. Add "dark mode" for the dark theme.

In Claude Code inside a clone of this repo, the skill `/gdg-ude-slides` does the same.

### Option 2: Use a template manually

| Application | Light | Dark |
|---|---|---|
| PowerPoint | [`templates/gdg-ude-light.potx`](templates/gdg-ude-light.potx) | [`templates/gdg-ude-dark.potx`](templates/gdg-ude-dark.potx) |
| LibreOffice Impress | [`templates/gdg-ude-light.otp`](templates/gdg-ude-light.otp) | [`templates/gdg-ude-dark.otp`](templates/gdg-ude-dark.otp) |

1. **Install the fonts** in [`assets/fonts/`](assets/fonts/) (Google Sans, Google Sans Medium and Google Sans Code; double-click → *Install*). Without them a fallback font is shown.
2. Open the template. It contains one sample slide per slide type with instructions. Delete the samples you do not need, and duplicate the ones you do.
3. PowerPoint: *Home → New Slide* offers all 18 layouts. Fill the footer via *Insert → Header & Footer → Footer* (`Title · Author · Date`) with *Slide number* on, then *Apply to All*. Impress: the layouts are master pages (*Slide → Change Slide Master*).
4. Follow the [content guidelines](docs/content-guidelines.md): one message per slide, at most 6 short bullets, details in the speaker notes.

### Option 3: Write HTML yourself

Copy [`html/starter.html`](html/starter.html), keep the links to `html/gdg-slides.css` and `html/gdg-slides.js` (or use the CDN links in the file), and write one `<section class="slide slide--TYPE">` per slide. Open it in a browser to present (← → keys, **O** overview, **T** light/dark, **F** fullscreen). For a PDF:

```bash
npm install && npx playwright install chromium
node tools/render.mjs path/to/deck.html --both      # deck-light.pdf + deck-dark.pdf, plus lint report
```

## The design at a glance

- **Font:** [Google Sans](https://fonts.google.com/specimen/Google+Sans), bundled. Titles 36 pt, body 24 pt, minimum 16 pt.
- **Colours:** white (or `#202124` in dark mode) with Google blue `#4285F4`, red `#EA4335`, yellow `#F9AB00` and green `#34A853` as accents, cycling in that order.
- **Base slides:** a four-colour bar above the title, logo + `title · author · date` footer, the slide number in a blue corner bubble, and small corner bubbles.
- **Special slides:** title, agenda, chapter, questions and next event. The bubbles and dot grids come from the *du + tech = ?* poster.
- **Slide types:** title, agenda, chapter, default (text + picture), full width / table, two columns, picture, code, key facts, statement / quote, canvas, references, questions, next event.

Details: [`docs/design-rules.md`](docs/design-rules.md) · [`docs/slide-types.md`](docs/slide-types.md) · [`docs/content-guidelines.md`](docs/content-guidelines.md) · tokens in [`design/tokens.json`](design/tokens.json)

## Repository structure

```
AGENTS.md                  instructions for AI agents (CLAUDE.md points here)
.claude/skills/            Claude Code skill "gdg-ude-slides"
docs/                      design rules, slide types, content guidelines
design/tokens.json         colours, type scale, geometry, content rules
assets/logo/               GDG UDE logos (original, trimmed, mark only)
assets/fonts/              Google Sans, Google Sans Code (OFL)
assets/icons/              Material icons for the next-event slide (Apache 2.0)
html/                      gdg-slides.css + gdg-slides.js + starter.html
templates/                 .potx (PowerPoint) and .otp (Impress), light and dark
examples/                  example deck (HTML, JSON, PDF light/dark, PPTX), screenshots, images
tools/render.mjs           HTML → PDF/PNG + content linter (Playwright)
tools/gdg_pptx.py          build PowerPoint decks from Python or a JSON spec + linter
tools/build_templates.py   generates the .potx/.otp templates from the tokens
tools/svg2png.mjs          SVG diagrams → PNG with the bundled fonts
```

## Development

```bash
npm install && npx playwright install chromium   # HTML rendering
python3 -m pip install -r requirements.txt       # python-pptx
npm run templates     # rebuild templates/*.potx (+ *.otp if LibreOffice is installed)
npm run examples      # re-render examples/*.pdf and screenshots
python3 tools/gdg_pptx.py examples/example-deck.json examples/example-deck.pptx
```

If you change a design value, change it in `design/tokens.json` **and** in `html/gdg-slides.css`. Then rebuild the templates and examples.

## Licenses and trademarks

- Google Sans and Google Sans Code: SIL Open Font License 1.1 (`assets/fonts/OFL*.txt`, `TRADEMARKS.md`).
- Material icons: Apache License 2.0 (`assets/icons/LICENSE-ICONS.txt`).
- The GDG logo and the Google brand colours are trademarks of Google. Use them only for official GDG on Campus UDE purposes, following the GDG brand guidelines.
