# Content guidelines: how much goes on a slide?

GDG UDE decks serve two purposes:

1. **During the talk**, the audience should listen to the speaker and not read paragraphs.
2. **After the talk**, the slides are published as a PDF and should still make sense without the speaker.

These goals pull in opposite directions. The rules below are a compromise between them, based on the research summarised at the end.

## The rules

### 1. One message per slide, stated in the title

Write slide titles as a **short sentence that states the takeaway** (an *assertion*), not as a topic label.

| Topic title (avoid) | Assertion title (prefer) |
|---|---|
| Firestore | Firestore fits most new projects better than Realtime Database |
| Security | Never ship the default security rules |

Assertion titles make the downloaded PDF readable: someone who only reads the titles gets the storyline. Keep them to **≤ 2 lines (≈ 14 words)**. Agenda, references, "Key takeaways" and chapter slides may use plain labels.

### 2. Support the message with evidence, not paragraphs

Below the title, show a **picture, diagram, code, table, numbers, or 3–6 short bullets** that support the claim.

| | Recommended | Linter warns at |
|---|---|---|
| Bullets per list | 3–5 | > 6 |
| Words per bullet | ≤ 12 (one line) | > 22 |
| List depth | 1 level, sometimes 2 | > 2 levels |
| Words on a content slide (without title) | 15–50 | > 60 (default), > 70 (full / two columns) |
| Statement slide | ≤ 20 words | > 30 |
| Code | ≤ 12 lines | > 15 lines or > 80 chars per line |
| Table | ≤ 6 rows × 4 columns | > 8 rows |
| Very little text | only with a visual | < 8 words and no visual → info |

Bullets are **phrases, not sentences**. Drop articles and filler words ("Hosting serves your web app worldwide", not "Firebase Hosting is a service that can be used to serve your web app all over the world.").

### 3. Put the rest in the speaker notes

Anything the speaker *says* but the audience does not need to *see* goes into the speaker notes: PowerPoint notes, or `<aside class="notes">` in HTML. Examples are explanations, anecdotes and the exact numbers behind a chart. Notes are where the detail lives. They are not exported to the PDF, so the slide itself must still carry the claim (title) and its evidence.

### 4. Make the downloaded deck self-explanatory

- Use assertion titles (rule 1).
- Give diagrams and screenshots a caption that says what to look at.
- End every chapter, or the talk, with a **"Key takeaways"** slide of 3–5 bullets.
- Collect all sources on a **references** slide. Cite them on the slide as `[1]` where you use them.
- Add links (docs, repo, slides) as text or a QR code on the outro slide.

### 5. Visual hierarchy

- Use **bold** for the 1–3 most important words of a slide, never whole sentences.
- Use the marker highlight (`<mark>`) at most once per slide.
- Use colour only from the palette, and with meaning (callout types, chapter colour).

### 6. Deck structure

A typical 20–30 min GDG talk has **12–20 slides**:

```
Title → Agenda → (Chapter → 2–5 content slides) × 2–4 → Key takeaways → References → Questions → Next event
```

Budget about 1–2 minutes per content slide. Use a chapter slide for every new section, plus an agenda slide when the talk has ≥ 3 chapters.

## How agents should apply these rules

**Case A: the user dictates the slide text exactly.**
Use the user's text **verbatim**. Do not shorten, reword or move it to the notes. Run the linter. If a slide breaks a rule, **build it anyway** and report a short warning, e.g.:

> ⚠ Slide 5 has 94 words (recommended ≤ 60) and 8 bullets (max. 6). Consider splitting it into two slides or moving the details to the speaker notes. Shall I do that?

Also warn when a slide has very little text and no visual: it works in the talk, but readers of the PDF may not understand it.

**Case B: the user gives only topics, an outline or a few keywords.**
Choose the amount of text yourself, following the rules above:

- Turn each key point into an assertion title.
- Add 3–5 bullets of ≤ 12 words, or a visual, as evidence.
- Write what the speaker should say into the speaker notes (2–5 sentences per slide).
- Add agenda, chapter, key-takeaways, references (if you used sources), questions and, if the user mentions an upcoming event, next-event slides.
- Do not invent facts, numbers or quotes. Mark placeholders clearly (e.g. `TODO: add benchmark numbers`) and list them in your final message.

**In both cases:** after building, run the linter (`tools/render.mjs` for HTML, `Deck.save()` for PPTX). Fix overflow problems, because text running out of the slide is never acceptable. Then look at the rendered PNGs/PDF before you hand the deck over.

## Background: what the research says

- **Assertion–evidence structure.** Slides with a sentence headline plus visual evidence led to better comprehension and recall of complex concepts than topic–subtopic bullet slides (Garner & Alley; [ASEE paper](https://peer.asee.org/assertion-evidence-slides-appear-to-lead-to-better-comprehension-and-recall-of-more-complex-concepts.pdf), [Penn State overview](https://writing.engr.psu.edu/research.html)). Building assertion–evidence slides also improved the presenters' own understanding. Critics note that brief, well-phrased bullets do not harm technical talks ([Farkas](https://faculty.washington.edu/farkas/dfpubs/Farkas-Assessment%20Of%20Alleys%20Slide%20Design.pdf)). We therefore allow short bullets as evidence.
- **Redundancy principle (Mayer).** Learners do worse when on-screen text repeats the narration word for word, because reading and listening compete for attention ([Mayer & Moreno](https://www.sfu.ca/~jcnesbit/EDUC220/ThinkPaper/MayerMoreno2002.pdf)). Short on-screen text that highlights key words is fine, and slightly *different* wording from the narration can even help ([Yue, Bjork & Bjork 2013](https://sites.lifesci.ucla.edu/psych-bjorklab/wp-content/uploads/sites/13/2016/07/YueBjorkBjork2013_redundancy.pdf)). Hence: keywords on the slide, full sentences in the speaker's mouth (and notes).
- **Coherence and signalling.** Leave out decorative extras and highlight the essential ([Mayer's principles](https://services.dartmouth.edu/TDClient/1806/Portal/KB/PrintArticle?ID=171655)). Hence the restrained decorations and the "bold 1–3 words" rule.
- **Slides vs. documents.** Duarte distinguishes projected slides (≈ 40 words or less) from *slidedocs* meant for reading (≈ 150–200 words per page) ([Duarte](https://www.duarte.com/resources/guides-tools/how-to-improve-dense-slides-with-slidedocs/)). Our decks are projected first, so we stay at the projected end (≤ 60 words). Assertion titles, captions, takeaways and references cover the reading use case.
- **Font size.** University accessibility guides recommend 18–28 pt minimum body text, with 24 pt the most common figure for lecture halls ([UCL](https://www.ucl.ac.uk/isd/services/websites-apps/creating-accessible-content/slides), [Texas A&M](https://service.tamu.edu/TDClient/36/Portal/KB/Article/1374/Accessibility-Creating-accessible-presentations), [York](https://subjectguides.york.ac.uk/accessibility/presentations)). We use 24 pt body text, 36 pt titles and a 16 pt minimum.
