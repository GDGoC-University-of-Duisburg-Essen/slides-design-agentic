# Design rules

The slide design of **GDG on Campus University of Duisburg-Essen** is clean and white (or dark grey in dark mode), with a few Google-coloured accents. It borrows the bubbles and dot grids from our *du + tech = ?* poster.

All values are defined in [`design/tokens.json`](../design/tokens.json). The HTML implementation is [`html/gdg-slides.css`](../html/gdg-slides.css), and the PowerPoint/Impress templates are generated from the same values by [`tools/build_templates.py`](../tools/build_templates.py).

## 1. Canvas

| | |
|---|---|
| Aspect ratio | **16:9** |
| HTML canvas | 1920 × 1080 px |
| PowerPoint / Impress | 13.333 × 7.5 in (33.867 × 19.05 cm) |
| Conversion | 1 pt = 2 px · 1 px = 6350 EMU |
| Side margins | 112 px (0.78 in) left and right |
| Content area (base slides) | x 112–1808 px, y ≈ 300–940 px |
| Footer zone | bottom 112 px, kept free for logo, meta line and slide number |

## 2. Colours

### Brand accents (taken from the GDG logo)

| Name | Hex | Typical use |
|---|---|---|
| Blue | `#4285F4` | primary accent, slide-number bubble, 1st bullet / chapter / column |
| Red | `#EA4335` | 2nd accent, "important" callouts |
| Yellow | `#F9AB00` | 3rd accent, "warning" callouts, marker highlight. **Never use it for text on white.** |
| Green | `#34A853` | 4th accent, "tip" callouts |

When several items get colours (bullets, agenda numbers, chapters, columns, fact cards), cycle through **blue → red → yellow → green**.

### Light mode (default)

| Role | Hex |
|---|---|
| Background | `#FFFFFF` |
| Surface (cards, code, table bands) | `#F8F9FA` / `#F1F3F4` |
| Text | `#1E1E1E` |
| Muted text (subtitles, captions, footer) | `#5F6368` |
| Lines | `#DADCE0` |
| Accent text, links | `#1A73E8` (passes WCAG AA, unlike `#4285F4` for small text) |
| Logo | `assets/logo/gdg_ude_logo_darktext_trimmed.png` |

### Dark mode

| Role | Hex |
|---|---|
| Background | `#202124` |
| Surface | `#2A2B2E` / `#35363A` |
| Text | `#E8EAED` |
| Muted text | `#9AA0A6` |
| Lines | `#3C4043` |
| Accent text, links | `#8AB4F8` |
| Logo | `assets/logo/gdg_ude_logo_lighttext_trimmed.png` |

The brand accents stay the same in both modes. Use dark mode in dark lecture halls or when the talk itself is about dark-themed UIs. **Never mix modes within one deck.**

In PowerPoint/Impress the colours are theme colours: *Text 1* = text, *Background 1* = background, *Text 2* = muted, *Background 2* = surface, *Accent 1–4* = blue/red/yellow/green, *Accent 5* = accent text, *Accent 6* = line colour.

## 3. Typography

**Font:** [Google Sans](https://fonts.google.com/specimen/Google+Sans) for all text and [Google Sans Code](https://fonts.google.com/specimen/Google+Sans+Code) for code. Both are under the SIL Open Font License. The static TTF files are bundled in [`assets/fonts/`](../assets/fonts/). The HTML version loads them from there, so PDFs render correctly offline. **Install them on your machine** before you use the PowerPoint/Impress templates; otherwise a fallback font is shown.

| Style | Size (pt) | Size (px @1920) | Weight | Use |
|---|---|---|---|---|
| Display | 54 | 108 | Bold | title slide, chapter title |
| Outro title | 80 | 160 | Bold | "Any questions?" |
| Statement | 44 | 88 | Medium | statement / quote slide |
| Slide title | **36** | 72 | Medium | every content slide, max. 2 lines |
| Subtitle / lead | 28 | 56 | Regular, muted | subtitles, agenda items |
| **Body** | **24** | 48 | Regular | bullets, text (level 1) |
| Body level 2 | 20 | 40 | Regular, muted | sub-bullets, tables, dense columns |
| Code | 18 | 36 | Google Sans Code | code blocks |
| Small | 16 | 32 | Regular | captions, references, sources: **absolute minimum for content** |
| Footer | 12 | 24 | Regular, muted | footer meta line only |

Why "medium": university accessibility guides recommend 18–28 pt as the minimum body size, and 24 pt is the most common recommendation for lecture halls ([UCL](https://www.ucl.ac.uk/isd/services/websites-apps/creating-accessible-content/slides), [Texas A&M](https://service.tamu.edu/TDClient/36/Portal/KB/Article/1374/Accessibility-Creating-accessible-presentations), [Hope College](https://hope.edu/offices/disability-services/faculty-staff-resources/accessible-presentations.html)). We use 24 pt for body text and 36 pt for titles. That reads well from the back of a lecture hall without making the slides look like posters.

Further rules:

- Line height: 1.35 for body text, 1.15 for titles, 1.08 for display text.
- Left-align text. Centre only the statement slide and short labels.
- Use **bold** to emphasise a few words per slide. Do not use underline (it is reserved for links in PDFs), CAPS (except small eyebrow labels) or more than one emphasis style at once.
- Use sentence case for titles ("Firebase lets you ship without a server"), not Title Case.

## 4. Layout and grid

- Every base slide has the same frame: a **four-colour bar** (128 × 8 px) at the top left, the **title** below it, the content area, and the **footer**.
- Content aligns to the left margin (112 px). The default two-column split is text (≈ 916 px) plus picture (700 px) with an 80 px gap.
- Leave whitespace. Do not fill every slide edge to edge, and do not shrink text to make it fit. Split the slide instead.
- Pictures get rounded corners (24 px radius) in HTML. Screenshots can get a thin frame (`figure.frame`).

## 5. Base slide chrome

All slides **except** title, chapter, outro (questions) and next-event carry:

| Element | Position |
|---|---|
| GDG UDE logo | bottom left, 40 px high |
| Meta line `Presentation title · Author · Date` | footer, right of the logo, 12 pt muted |
| Slide number | white, inside a blue quarter-circle in the bottom-right corner |
| Corner bubbles | small yellow quarter-circle and green dot in the top-right corner |
| Colour bar | four Google colours above the slide title |

Do not add further logos, page frames, background images or gradients to base slides.

## 6. Decorations ("bubbles")

The decorations come from the *du + tech = ?* poster: quarter/half circles bleeding off the slide edge, a yellow block with one rounded corner, and 4 × 4 dot grids (dots 14 px, 40 px pitch).

| Slide | Decorations |
|---|---|
| Base slides | top-right yellow quarter-circle + green dot, blue number bubble |
| Title, outro | poster cluster on the right: yellow block, blue circle, green and blue dot grids, red and yellow circles bottom right, small red dot |
| Chapter | large half-circle on the left in the chapter colour (with number or image), dot grid top right, faint bubble bottom right |
| Next event | yellow block + blue circle top right, green dot grid, red bubble bottom right |

Decorations stay inside these positions. Never place content under them, never animate them, and never add extra ones on content slides.

## 7. Logo

- Use the files in `assets/logo/`. The `*_trimmed.png` versions have no padding and are used by the templates.
- Use `darktext` on light backgrounds and `lighttext` on dark backgrounds. `gdg_ude_mark.png` (brackets only) is for very small spaces such as social previews.
- Never recolour, stretch, rotate or re-type the logo. Keep clear space of at least the height of the brackets' red/green part around it.

## 8. Images, diagrams, QR codes

- Prefer one meaningful picture, diagram or screenshot over decorative stock photos.
- Diagrams use the brand colours, Google Sans and at least 16 pt text (at the size the diagram is shown).
- QR codes are always dark on a white card, at least 200 px (HTML) / 1.4 in wide, with the quiet zone kept.
- Posters on the next-event slide keep the DIN A3 portrait ratio 1 : 1.414.
- Every image in HTML gets an `alt` text.

## 9. Accessibility checklist

- Contrast: text colours above pass WCAG AA on their backgrounds. Never put text on yellow except dark text.
- Never use colour as the only carrier of meaning (e.g. in charts add labels).
- Minimum 16 pt for any content. Body text is 24 pt.
- One message per slide, logical reading order (title → content).
