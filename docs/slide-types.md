# Slide types

Each slide type exists in three forms with the same look:

- **HTML:** `<section class="slide slide--TYPE">` inside `<main class="deck">` (see [`html/starter.html`](../html/starter.html))
- **PowerPoint / Impress:** a layout (master page) of the same name in `templates/*.potx` / `*.otp`
- **Python / JSON:** a method of `tools/gdg_pptx.py` → `Deck` / a `{"type": …}` entry in a JSON deck spec

Screenshots of every type: [`examples/example-deck-light.pdf`](../examples/example-deck-light.pdf).

| Type | Footer | Layout (PPTX / Impress) | `Deck` method | JSON `type` |
|---|---|---|---|---|
| [Title](#title) | – | Title | `title_slide()` | `title` |
| [Agenda](#agenda--table-of-contents) | ✓ | Agenda | `agenda()` | `agenda` |
| [Chapter](#chapter) | – | Chapter Blue / Red / Yellow / Green | `chapter()` | `chapter` |
| [Default](#default-text--optional-picture) | ✓ | Content with Picture / Content | `content()` | `content` |
| [Full width](#full-width-text-table) | ✓ | Content | `full()` | `full` |
| [Two columns](#two-columns) | ✓ | Two Columns | `two_columns()` | `two-columns` |
| [Picture](#picture) | ✓ | Picture | `image()` | `image` |
| [Code](#code) | ✓ | Code | `code()` | `code` |
| [Key facts](#key-facts) | ✓ | Key Facts | `facts()` | `facts` |
| [Statement](#statement--quote) | ✓ | Statement | `statement()` | `statement` |
| [Canvas](#canvas-title-only) | ✓ | Title Only | `title_only()` | `canvas` |
| [References](#references) | ✓ | References | `references()` | `references` |
| [Questions (outro)](#questions-outro) | – | Questions | `questions()` | `questions` |
| [Next event](#next-event-special-slide) | – | Next Event | `next_event()` | `next-event` |
| Blank | ✓ | Blank with Footer | – | – |

---

## HTML deck skeleton

```html
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <title>Talk title</title>
  <link rel="stylesheet" href="PATH/html/gdg-slides.css">
  <script src="PATH/html/gdg-slides.js" defer></script>
</head>
<body>
<main class="deck" data-title="Talk title" data-author="Speaker" data-date="2026-11-22" data-lang="en">
  <!-- slides -->
</main>
</body>
</html>
```

- `data-title`, `data-author` and `data-date` (ISO date, formatted automatically; `data-lang="de"` → German months) fill the footer of every base slide and, if omitted on the slide, the title slide.
- `?theme=dark` in the URL or `data-theme="dark"` on `.deck` switches to dark mode. Press **T** to toggle while presenting.
- The footer, slide numbers, decorations, chapter numbers and colours are added by `gdg-slides.js`. **Do not write them by hand.**
- `<aside class="notes">…</aside>` inside a slide holds speaker notes (hidden).
- Inline helpers: `<strong>`, `<mark>` (yellow marker), `<code>`, `.muted`, `.accent`, `.text-blue|red|yellow|green`, `.small`, `.lead`, `.eyebrow`, `.chip`.
- Callouts: `<div class="callout">`, `.callout.tip` (green), `.callout.warning` (yellow), `.callout.important` (red). The first `<strong>` becomes the label.

---

## Title

The first slide: title, optional subtitle, author, date and logo, with the poster bubbles on the right.

```html
<section class="slide slide--title">
  <h1>Your First Web App with Firebase</h1>
  <p class="subtitle">From zero to a deployed app in one evening</p>   <!-- optional -->
  <p class="author">Max Mustermann</p>
  <p class="date">2026-11-22</p>
</section>
```
```python
deck.title_slide(subtitle="From zero to a deployed app in one evening")
```

Rules: title max. 2 lines (≈ 50 characters), subtitle max. 1 line.

## Agenda / table of contents

Numbered chapter overview. The numbers cycle through the Google colours, each with an optional one-line description. More than 5 items flow into two columns automatically. Use `data-current="2"` (HTML) / `current=2` (Python) to highlight the current chapter when you repeat the agenda between chapters.

```html
<section class="slide slide--agenda">
  <h2>Agenda</h2>
  <ol class="agenda">
    <li><strong>Why Firebase?</strong><span>What it is and when it fits</span></li>
    <li><strong>Hands-on</strong><span>Set up, store data, deploy</span></li>
    <li><strong>Q&amp;A</strong></li>
  </ol>
</section>
```
An **empty** `<ol class="agenda"></ol>` (HTML) or `deck.agenda()` without items (Python) is filled automatically from the chapter slides.

In PowerPoint/Impress, type `01`, press Tab, then the chapter name. Use the second list level (Tab at line start) for the description.

## Chapter

Section divider before each chapter: a large half circle in the chapter colour with the chapter number, the title, and an optional subtitle. Colours cycle blue → red → yellow → green automatically (override with `data-color` / `color=`). An optional image fills the half circle, and the number then moves into a badge.

```html
<section class="slide slide--chapter">
  <img src="images/chapter.jpg" alt="">          <!-- optional -->
  <h2>Hands-on: your first app</h2>
  <p class="subtitle">Set up, store data, deploy</p>   <!-- optional -->
</section>
```
```python
deck.chapter("Hands-on: your first app", subtitle="Set up, store data, deploy", image="chapter.jpg")
```
PowerPoint: pick the layout *Chapter Blue/Red/Yellow/Green*. If you insert a picture, move or delete the number.

## Default: text + optional picture

The workhorse: title, bullets on the left, and an optional picture on the right. Without a picture, the text uses the full width.

```html
<section class="slide slide--default">
  <h2>Firebase lets you ship a full app without running a server</h2>
  <div class="body">
    <div class="text">
      <ul>
        <li><strong>Hosting</strong> serves your web app worldwide</li>
        <li><strong>Firestore</strong> stores and syncs data</li>
      </ul>
    </div>
    <figure><img src="images/illustration.png" alt="…"></figure>   <!-- optional -->
  </div>
  <aside class="notes">What the speaker says …</aside>
</section>
```
```python
deck.content("Firebase lets you ship …", ["**Hosting** serves your web app", ("Firestore", ["sub-bullet"])],
             image="illustration.png", notes="…")
```

## Full width (text, table)

Full-width body for longer lists, tables and charts.

```html
<section class="slide slide--full">
  <h2>Pick the service that matches what your app needs</h2>
  <div class="body">
    <table>
      <thead><tr><th>Need</th><th>Service</th></tr></thead>
      <tbody><tr><td>Sign-in</td><td>Authentication</td></tr></tbody>
    </table>
  </div>
</section>
```
```python
deck.full("Pick the service …", table=[["Need", "Service"], ["Sign-in", "Authentication"]])
deck.full("Three pitfalls …", bullets=["**Security rules:** …", "…"])
```
Tables: the header row is bold with a blue rule, then banded rows at 20 pt. Use `class="num"` on numeric cells (HTML). Callouts also go here.

## Two columns

Comparison or two parallel lists. The coloured top rules cycle blue/red. Add `class="three"` to `.columns` for three columns (HTML only).

```html
<section class="slide slide--two-columns">
  <h2>Firestore fits most new projects better than Realtime Database</h2>
  <div class="columns">
    <div class="col"><h3>Cloud Firestore</h3><ul><li>…</li></ul></div>
    <div class="col"><h3>Realtime Database</h3><ul><li>…</li></ul></div>
  </div>
</section>
```
```python
deck.two_columns("Firestore fits …", ("Cloud Firestore", ["…"]), ("Realtime Database", ["…"]))
```

## Picture

One large picture, diagram or screenshot with an optional caption. Add `class="no-title"` and omit the `<h2>` for a picture-only slide.

```html
<section class="slide slide--image">
  <h2>A typical student project combines three services</h2>
  <figure>
    <img src="images/architecture.png" alt="…">
    <figcaption>What to look at in this picture.</figcaption>
  </figure>
</section>
```
```python
deck.image("A typical student project …", "architecture.png", caption="…")
```

## Code

Code block in Google Sans Code 18 pt, with an optional explanation column. HTML: highlight.js class names are supported. Include highlight.js, or wrap tokens in `<span class="hljs-keyword">` etc.

```html
<section class="slide slide--code">
  <h2>Saving an event to Firestore takes a few lines</h2>
  <div class="body">
    <pre><code class="language-js">await addDoc(collection(db, "events"), {…});</code></pre>
    <div class="explain"><ul><li>…</li></ul></div>   <!-- optional -->
  </div>
</section>
```
```python
deck.code("Saving an event …", 'await addDoc(collection(db, "events"), {...});')
```

## Key facts

2–4 big numbers with short labels, in cards that cycle the Google colours.

```html
<section class="slide slide--facts">
  <h2>Today's app needs surprisingly little</h2>
  <div class="facts">
    <div class="fact"><strong>3</strong><span>Firebase services</span></div>
    <div class="fact"><strong>1</strong><span>command to deploy</span></div>
  </div>
</section>
```
```python
deck.facts("Today's app needs surprisingly little", [("3", "Firebase services"), ("1", "command to deploy")])
```

## Statement / quote

One sentence (≤ 20 words) as a big statement, or a quote with attribution.

```html
<section class="slide slide--statement">
  <p class="statement">One strong statement per slide is easier to remember.</p>
  <p class="attribution">Optional attribution</p>
</section>
<!-- or as a quote (with a blue quotation mark) -->
<section class="slide slide--statement">
  <blockquote><p>Ship a small thing early - then iterate.</p><footer>Our rule of thumb</footer></blockquote>
</section>
```
```python
deck.statement("Ship a small thing early - then iterate.", attribution="Our rule of thumb")
```

## Canvas (title only)

Title plus a free area (x 112–1808 px, y ≈ 300–940 px) for custom diagrams.

```html
<section class="slide slide--canvas"><h2>…</h2><div class="body"><!-- free content --></div></section>
```
```python
s = deck.title_only("…")   # then add python-pptx shapes to s
```

## References

Numbered list `[1] …` at 16 pt. More than 5 entries flow into two columns (HTML); more than 8 are split across slides (Python).

```html
<section class="slide slide--references">
  <h2>References</h2>
  <ol class="references">
    <li>Google. <cite>Firebase Hosting documentation</cite>. <a href="https://firebase.google.com/docs/hosting">firebase.google.com/docs/hosting</a></li>
  </ol>
</section>
```
```python
deck.references(["Google. Firebase Hosting documentation. https://firebase.google.com/docs/hosting"])
```

## Questions (outro)

Large "Any questions?" with a blue question mark, a thank-you line, contact details and an optional QR code (slides, repo, feedback form).

```html
<section class="slide slide--outro">
  <h2>Any questions?</h2>
  <p class="subtitle">Thank you for joining!</p>
  <ul class="contact">
    <li class="name">Max Mustermann</li>
    <li>max.mustermann@example.com</li>
    <li>Instagram: @gdg_ude</li>
  </ul>
  <figure class="qr"><img src="qr.svg" alt="QR code to the slides"><figcaption>Slides &amp; code</figcaption></figure>
</section>
```
```python
deck.questions(contact=["Max Mustermann", "max@example.com", "Instagram: @gdg_ude"], qr="qr.png", qr_caption="Slides & code")
```

## Next event (special slide)

Advertises the next GDG UDE event at the end of a talk or before a lecture. It shows the DIN A3 poster (portrait 1 : 1.414) on the left, then the logo, the *Next event* eyebrow, the event title, date/time and location with icons, a short description (≤ 3 lines), the **DU + TECH = ?** tagline and a QR code card. It has no footer.

```html
<section class="slide slide--next-event">
  <img class="poster" src="poster.png" alt="Event poster">
  <p class="eyebrow">Next event</p>
  <h2>UDE meets Google: Semester Kick-off</h2>
  <ul class="event-meta">
    <li class="date">Wednesday, 3 December 2026 · 18:00</li>   <!-- classes: date, time, place, link -->
    <li class="place">SH-403, Schützenbahn 70, Essen</li>
  </ul>
  <p class="description">Talks, workshops and community. Open to all students.</p>
  <figure class="qr"><img src="qr.svg" alt="QR code to register"><figcaption>Register now</figcaption></figure>
</section>
```
```python
deck.next_event("UDE meets Google: Semester Kick-off", "Wednesday, 3 December 2026 · 18:00",
                "SH-403, Schützenbahn 70, Essen", "Talks, workshops and community.",
                poster="poster.png", qr="qr.png")
```
`data-tagline="off"` hides the tagline. Without a poster image, a dashed placeholder is shown. To generate a QR code: `python -c "import segno; segno.make('https://…').save('qr.svg', scale=10, border=0)"` (or any QR tool).

---

## JSON deck spec (for `tools/gdg_pptx.py`)

```json
{
  "title": "Talk title", "author": "Speaker", "date": "2026-11-22", "theme": "light", "lang": "en",
  "slides": [
    {"type": "title", "subtitle": "…"},
    {"type": "agenda"},
    {"type": "chapter", "title": "…", "subtitle": "…", "image": "optional.png"},
    {"type": "content", "title": "…", "bullets": ["…", ["with sub-bullets", ["…"]]], "image": "optional.png", "notes": "…"},
    {"type": "full", "title": "…", "table": [["Header", "…"], ["Cell", "…"]]},
    {"type": "two-columns", "title": "…", "left": ["Heading", ["…"]], "right": ["Heading", ["…"]]},
    {"type": "image", "title": "…", "image": "diagram.png", "caption": "…"},
    {"type": "code", "title": "…", "code": "…"},
    {"type": "facts", "title": "…", "facts": [["3", "label"], ["1", "label"]]},
    {"type": "statement", "text": "…", "attribution": "…"},
    {"type": "references", "refs": ["…"]},
    {"type": "questions", "contact": ["Name", "e-mail"], "qr": "qr.png", "qr_caption": "Slides"},
    {"type": "next-event", "title": "…", "date": "…", "place": "…", "description": "…", "poster": "poster.png", "qr": "qr.png"}
  ]
}
```
Paths are relative to the JSON file. Every slide accepts `"notes"`. A complete example is [`examples/example-deck.json`](../examples/example-deck.json).
