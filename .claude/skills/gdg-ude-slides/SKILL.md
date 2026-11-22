---
name: gdg-ude-slides
description: Build presentation slides (HTML, PDF or PowerPoint) in the slide design of GDG on Campus University of Duisburg-Essen (GDG UDE). Use when the user asks for slides, a deck, a talk or a presentation for GDG UDE / Google Developer Group Duisburg-Essen, or references the slides-design-agentic repository.
---

# GDG UDE slides

The design system lives in the repository `https://github.com/GDGoC-University-of-Duisburg-Essen/slides-design-agentic`.

1. Locate the repository root: it contains `AGENTS.md`, `html/gdg-slides.css` and `tools/gdg_pptx.py`. If you are not inside it, clone it:
   `git clone --depth 1 https://github.com/GDGoC-University-of-Duisburg-Essen/slides-design-agentic.git`
2. Read `AGENTS.md` and follow it step by step. Also read `docs/content-guidelines.md` and `docs/slide-types.md`.
3. Build the deck (HTML → `tools/render.mjs` for PDF; PowerPoint → `tools/gdg_pptx.py`). Lint it, look at the rendered previews, and fix overflows.
4. Report the files, the remaining linter warnings (⚠) and any TODO placeholders to the user.

Key rules: use the user's exact text verbatim when given (warn instead of shortening). Otherwise write assertion titles, 3–5 short bullets and speaker notes. Body 24 pt, titles 36 pt, minimum 16 pt. Use only the palette and slide types of the design system, and never invent facts.
