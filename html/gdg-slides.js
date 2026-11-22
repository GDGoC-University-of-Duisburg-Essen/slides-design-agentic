/* ==========================================================================
   GDG on Campus University of Duisburg-Essen - slide runtime
   --------------------------------------------------------------------------
   - adds footer (logo, "title · author · date", slide number) to base slides
   - adds the Google-style decorations, chapter numbers/colours, agenda
   - presenter mode (one slide, scaled), overview mode, print/PDF mode
   - content linter: window.GDG.lint() -> [{slide, type, level, message}]

   Keys: → ↓ Space PgDn next · ← ↑ PgUp previous · Home/End · O overview
         T toggle light/dark · F fullscreen · P print view
   URL:  deck.html?theme=dark  ·  deck.html#7 (open slide 7)  ·  ?mode=print
   ========================================================================== */
(function () {
  "use strict";

  const BASE_TYPES = ["default", "full", "two-columns", "image", "code", "facts",
    "statement", "canvas", "references", "agenda"];
  const SPECIAL_TYPES = ["title", "chapter", "outro", "next-event"];
  const COLORS = ["blue", "red", "yellow", "green"];

  // Text-amount rules - keep in sync with docs/content-guidelines.md
  const RULES = {
    maxWords: { default: 60, full: 70, "two-columns": 70, code: 35, canvas: 40, image: 25,
      facts: 50, statement: 30, agenda: 50, references: 9999, title: 30, chapter: 25,
      outro: 45, "next-event": 45 },
    minWordsWithoutVisual: 8,
    maxBullets: 6,
    maxWordsPerBullet: 22,
    maxListDepth: 2,
    maxTitleWords: 14,
    maxTitleLines: 2,
    maxAgendaItems: 8,
    maxReferences: 8,
    minFontPx: 32,         // 16 pt
  };

  const I18N = {
    en: { months: ["January", "February", "March", "April", "May", "June", "July",
      "August", "September", "October", "November", "December"] },
    de: { months: ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli",
      "August", "September", "Oktober", "November", "Dezember"] },
  };

  const params = new URLSearchParams(location.search);
  let deck, slides = [], current = 0, mode = "present";

  // ---------- helpers ------------------------------------------------------
  const el = (tag, cls, html) => {
    const e = document.createElement(tag);
    if (cls) e.className = cls;
    if (html != null) e.innerHTML = html;
    return e;
  };
  const typeOf = (s) => {
    for (const t of [...SPECIAL_TYPES, ...BASE_TYPES]) if (s.classList.contains("slide--" + t)) return t;
    return "default";
  };
  const words = (txt) => (txt || "").trim().split(/\s+/).filter((w) => /[\p{L}\p{N}]/u.test(w)).length;
  const lang = () => (deck.dataset.lang || document.documentElement.lang || "en").slice(0, 2).toLowerCase();

  function formatDate(raw) {
    if (!raw) return "";
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(raw.trim());
    if (!m) return raw;
    const l = I18N[lang()] ? lang() : "en";
    const month = I18N[l].months[parseInt(m[2], 10) - 1];
    const day = parseInt(m[3], 10);
    return l === "de" ? `${day}. ${month} ${m[1]}` : `${day} ${month} ${m[1]}`;
  }

  function meta() {
    const d = deck.dataset;
    return {
      title: d.title || document.title || "",
      author: d.author || "",
      date: formatDate(d.date || ""),
      footer: d.footer, // optional custom footer text
    };
  }

  // ---------- build --------------------------------------------------------
  function build() {
    deck = document.querySelector(".deck") || document.body;
    slides = Array.from(deck.querySelectorAll(":scope > .slide, :scope > .slide-wrap > .slide"));
    const m = meta();
    const theme = params.get("theme") || deck.dataset.theme;
    if (theme) document.documentElement.dataset.theme = theme;

    let chapterIdx = 0;
    const chapters = [];

    slides.forEach((s, i) => {
      const type = typeOf(s);
      s.dataset.type = type;
      s.dataset.index = i + 1;

      // wrapper (used by overview mode only)
      if (!s.parentElement.classList.contains("slide-wrap")) {
        const w = el("div", "slide-wrap");
        s.before(w);
        w.appendChild(s);
      }

      // decorations
      if (s.dataset.deco !== "off" && !s.querySelector(":scope > .deco")) {
        const d = el("div", "deco");
        d.setAttribute("aria-hidden", "true");
        for (let k = 1; k <= 7; k++) d.appendChild(el("i", "b" + k + ([3, 7].includes(k) && ["title", "outro"].includes(type) ? " dots" : "")));
        if (type === "chapter") d.querySelector(".b1").classList.add("dots");
        if (type === "next-event") d.querySelector(".b3").classList.add("dots");
        s.prepend(d);
      }

      // brand logo on special slides
      if (["title", "outro", "next-event"].includes(type) && !s.querySelector(".brand-logo")) {
        const logo = el("div", "brand-logo");
        logo.setAttribute("role", "img");
        logo.setAttribute("aria-label", "Google Developer Group University of Duisburg-Essen");
        s.appendChild(logo);
      }

      if (type === "title") {
        if (!s.querySelector("h1")) s.appendChild(el("h1", "", m.title));
        if (!s.querySelector(".author") && m.author) s.appendChild(el("p", "author", m.author));
        if (!s.querySelector(".date") && m.date) s.appendChild(el("p", "date", m.date));
        const dt = s.querySelector(".date");
        if (dt && /^\d{4}-\d{2}-\d{2}$/.test(dt.textContent.trim())) dt.textContent = formatDate(dt.textContent);
      }

      if (type === "chapter") {
        chapterIdx++;
        const color = s.dataset.color || COLORS[(chapterIdx - 1) % 4];
        s.dataset.color = color;
        const num = s.dataset.number || String(chapterIdx).padStart(2, "0");
        if (!s.querySelector(".chapter-circle")) {
          const circle = el("div", "chapter-circle");
          const img = s.querySelector(":scope > img, :scope > figure img");
          if (img) {
            const fig = img.closest("figure");
            circle.appendChild(img);
            if (fig && fig.parentElement === s) fig.remove();
            s.classList.add("has-image");
          }
          s.prepend(circle);
          circle.after(el("div", "chapter-number", num));
        }
        const h = s.querySelector("h2");
        const sub = s.querySelector(".subtitle");
        chapters.push({ title: h ? h.innerHTML : "", subtitle: sub ? sub.innerHTML : "", slide: i + 1 });
      }

      if (type === "outro") {
        const h = s.querySelector("h2");
        if (h && !h.querySelector(".q")) h.innerHTML = h.innerHTML.replace(/([?？])\s*$/, '<span class="q">$1</span>');
      }

      if (type === "next-event" && s.dataset.tagline !== "off" && !s.querySelector(".tagline")) {
        const t = el("div", "tagline", 'DU <span class="plus">+</span> TECH = <span class="q">?</span>');
        s.appendChild(t);
      }

      if (type === "next-event" && !s.querySelector(".poster")) {
        s.appendChild(el("div", "poster poster-placeholder", "DIN A3 poster<br>(1 : 1.414)"));
      }

      // footer on base slides
      if (BASE_TYPES.includes(type) && s.dataset.footer !== "off" && !s.querySelector(".slide-footer")) {
        const f = el("footer", "slide-footer");
        f.setAttribute("aria-hidden", "true");
        f.appendChild(el("span", "logo"));
        f.appendChild(el("span", "sep"));
        const metaText = m.footer || [m.title, m.author, m.date].filter(Boolean).join("  ·  ");
        f.appendChild(el("span", "meta", metaText));
        f.appendChild(el("span", "num", String(i + 1)));
        s.appendChild(f);
      }
    });

    // agenda: auto-fill from chapter slides when the list is empty
    deck.querySelectorAll(".slide--agenda").forEach((s) => {
      let ol = s.querySelector("ol.agenda");
      if (!ol) { ol = el("ol", "agenda"); s.appendChild(ol); }
      if (!ol.children.length) {
        chapters.forEach((c) => {
          const li = el("li", "", `<strong>${c.title}</strong>` + (c.subtitle ? `<span>${c.subtitle}</span>` : ""));
          ol.appendChild(li);
        });
      }
      const items = Array.from(ol.children);
      const rows = items.length > 5 ? Math.ceil(items.length / 2) : items.length;
      ol.style.setProperty("--rows", rows);
      ol.style.setProperty("--cols", items.length > 5 ? 2 : 1);
      if (s.dataset.current) {
        items.forEach((li, k) => {
          if (k + 1 === parseInt(s.dataset.current, 10)) li.classList.add("current");
        });
      }
    });

    deck.querySelectorAll(".slide--references ol.references").forEach((ol) => {
      if (ol.children.length > 5 && !ol.style.getPropertyValue("--cols")) ol.style.setProperty("--cols", 2);
    });

    if (window.hljs) deck.querySelectorAll("pre code").forEach((c) => window.hljs.highlightElement(c));
  }

  // ---------- modes --------------------------------------------------------
  function fit() {
    const s = Math.min(window.innerWidth / 1920, window.innerHeight / 1080);
    document.documentElement.style.setProperty("--scale", s);
    const cols = Math.max(1, Math.floor((window.innerWidth - 48) / 520));
    const os = Math.min(0.5, (window.innerWidth - 48 - (cols - 1) * 24) / cols / 1920);
    document.documentElement.style.setProperty("--oscale", os);
  }

  function setMode(m) {
    mode = m;
    document.body.classList.toggle("gdg-present", m === "present");
    document.body.classList.toggle("gdg-overview", m === "overview");
    document.body.classList.toggle("gdg-print", m === "print");
    if (m === "present") show(current);
    fit();
  }

  function show(i) {
    current = Math.max(0, Math.min(slides.length - 1, i));
    slides.forEach((s, k) => s.classList.toggle("active", k === current));
    if (mode === "present" && location.hash !== "#" + (current + 1)) history.replaceState(null, "", "#" + (current + 1));
  }

  function toggleTheme() {
    const r = document.documentElement;
    r.dataset.theme = r.dataset.theme === "dark" ? "light" : "dark";
  }

  function onKey(e) {
    if (e.target.closest && e.target.closest("input, textarea, [contenteditable]")) return;
    const k = e.key;
    if (["ArrowRight", "ArrowDown", "PageDown", " ", "Enter"].includes(k)) { show(current + 1); e.preventDefault(); }
    else if (["ArrowLeft", "ArrowUp", "PageUp", "Backspace"].includes(k)) { show(current - 1); e.preventDefault(); }
    else if (k === "Home") show(0);
    else if (k === "End") show(slides.length - 1);
    else if (k === "o" || k === "O" || k === "Escape") setMode(mode === "overview" ? "present" : "overview");
    else if (k === "t" || k === "T") toggleTheme();
    else if (k === "p" || k === "P") setMode(mode === "print" ? "present" : "print");
    else if (k === "f" || k === "F") {
      if (document.fullscreenElement) document.exitFullscreen(); else document.documentElement.requestFullscreen();
    }
  }

  // ---------- lint ---------------------------------------------------------
  function lint() {
    const out = [];
    const prevMode = mode;
    if (mode !== "print") setMode("print");
    const add = (s, level, message) => out.push({ slide: +s.dataset.index, type: s.dataset.type, level, message });

    slides.forEach((s) => {
      const type = s.dataset.type;
      const sr = s.getBoundingClientRect();
      const scale = sr.width / 1920 || 1;
      const isBase = BASE_TYPES.includes(type) && s.dataset.footer !== "off";
      const bottomLimit = isBase ? 952 : 1060;
      const skip = (n) => n.closest(".deco, .slide-footer, .notes, .chapter-circle, .chapter-number, .brand-logo, .tagline");

      // 1. overflow
      let worst = 0, worstRight = 0;
      s.querySelectorAll("*").forEach((n) => {
        if (skip(n)) return;
        const r = n.getBoundingClientRect();
        if (!r.width || !r.height) return;
        const b = (r.bottom - sr.top) / scale, rt = (r.right - sr.left) / scale;
        if (b > bottomLimit + 2) worst = Math.max(worst, b - bottomLimit);
        if (rt > 1920 - 40 && !n.matches("img, svg, .poster")) worstRight = Math.max(worstRight, rt - 1880);
        const cs = getComputedStyle(n);
        if (/(hidden|clip|auto|scroll)/.test(cs.overflow) && n.scrollHeight > n.clientHeight + 4 && !n.matches("img, svg"))
          add(s, "warn", `content is cut off inside <${n.tagName.toLowerCase()}${n.className ? "." + String(n.className).split(" ")[0] : ""}> (${n.scrollHeight - n.clientHeight}px hidden)`);
      });
      // next-event: free-flowing text must not run into the tagline / QR code
      const tag = s.querySelector(":scope > .tagline");
      if (type === "next-event" && tag) {
        const tTop = (tag.getBoundingClientRect().top - sr.top) / scale;
        s.querySelectorAll(":scope > p, :scope > ul, :scope > h2").forEach((n) => {
          const b = (n.getBoundingClientRect().bottom - sr.top) / scale;
          if (b > tTop - 16) add(s, "warn", `text runs into the "DU + TECH = ?" tagline - shorten the description (max. ~3 lines) or the title`);
        });
      }
      if (worst > 0) add(s, "warn", `content overflows the slide area by ${Math.round(worst)}px - shorten the text or split the slide`);
      if (worstRight > 0) add(s, "warn", `content runs ${Math.round(worstRight)}px past the right margin`);

      // 2. text amount
      const clone = s.cloneNode(true);
      clone.querySelectorAll(".deco, .slide-footer, .notes, h1, h2, .chapter-number, .brand-logo, .tagline, pre, code, figcaption, .eyebrow").forEach((n) => n.remove());
      const bodyWords = words(clone.textContent);
      const max = RULES.maxWords[type] ?? 60;
      if (bodyWords > max) add(s, "warn", `too much text: ${bodyWords} words on the slide (recommended max. ${max} for "${type}") - move details to speaker notes or split the slide`);
      const hasVisual = s.querySelector("img, svg:not(.deco svg), table, pre, figure, .facts, .columns, video, canvas, blockquote, .statement");
      if (["default", "full"].includes(type) && !hasVisual && bodyWords < RULES.minWordsWithoutVisual)
        add(s, "info", `very little text (${bodyWords} words) and no visual - fine for the talk, but readers of the PDF may miss context. Consider an assertion-style title or one more line.`);

      // 3. title
      const h = s.querySelector(":scope > h2, :scope > h1");
      if (h && !["outro"].includes(type)) {
        const tw = words(h.textContent);
        if (tw > RULES.maxTitleWords) add(s, "warn", `title has ${tw} words (max. ${RULES.maxTitleWords})`);
        const lh = parseFloat(getComputedStyle(h).lineHeight) || 80;
        const bar = type === "title" ? 0 : 32;
        const lines = Math.round((h.getBoundingClientRect().height / scale - bar) / lh);
        if (lines > RULES.maxTitleLines) add(s, "warn", `title wraps to ${lines} lines (max. ${RULES.maxTitleLines})`);
      }
      if (BASE_TYPES.includes(type) && !h && type !== "statement" && type !== "image") add(s, "info", "slide has no title (<h2>)");

      // 4. bullets
      s.querySelectorAll("ul, ol").forEach((list) => {
        if (list.closest(".notes, .slide-footer") || list.matches(".agenda, .references, .contact, .event-meta")) return;
        if (list.parentElement.closest("ul, ol")) return; // top-level lists only
        const items = list.querySelectorAll(":scope > li");
        if (items.length > RULES.maxBullets) add(s, "warn", `${items.length} bullet points in one list (max. ${RULES.maxBullets})`);
        items.forEach((li, k) => {
          const c = li.cloneNode(true);
          c.querySelectorAll("ul, ol").forEach((n) => n.remove());
          const w = words(c.textContent);
          if (w > RULES.maxWordsPerBullet) add(s, "warn", `bullet ${k + 1} has ${w} words (max. ${RULES.maxWordsPerBullet}) - make it a short phrase`);
        });
        if (list.querySelector("li li li")) add(s, "warn", `list nesting deeper than ${RULES.maxListDepth} levels`);
      });
      const agenda = s.querySelector("ol.agenda");
      if (agenda && agenda.children.length > RULES.maxAgendaItems) add(s, "warn", `agenda has ${agenda.children.length} items (max. ${RULES.maxAgendaItems})`);
      const refs = s.querySelector("ol.references");
      if (refs && refs.children.length > RULES.maxReferences * 2) add(s, "warn", `${refs.children.length} references on one slide - split into several reference slides`);

      // 5. minimum font size
      let small = 0;
      const walker = document.createTreeWalker(s, NodeFilter.SHOW_TEXT);
      while (walker.nextNode()) {
        const p = walker.currentNode.parentElement;
        if (!walker.currentNode.textContent.trim() || skip(p)) continue;
        if (parseFloat(getComputedStyle(p).fontSize) < RULES.minFontPx - 0.5) small++;
      }
      if (small) add(s, "warn", `${small} text element(s) smaller than 16 pt - not readable in a lecture hall`);

      // 6. images without alt text
      s.querySelectorAll("img").forEach((img) => {
        if (!img.hasAttribute("alt")) add(s, "info", `image without alt text: ${img.getAttribute("src")}`);
        if (img.complete && img.naturalWidth === 0) add(s, "warn", `image failed to load: ${img.getAttribute("src")}`);
      });
    });

    if (prevMode !== "print") setMode(prevMode);
    return out;
  }

  // ---------- init ---------------------------------------------------------
  function init() {
    build();
    const m = params.get("mode") || (location.hash === "#overview" ? "overview" : "present");
    const n = parseInt((location.hash || "").slice(1), 10);
    if (n) current = n - 1;
    setMode(m);
    window.addEventListener("resize", fit);
    document.addEventListener("keydown", onKey);
    deck.addEventListener("click", (e) => {
      if (mode !== "overview") return;
      const s = e.target.closest(".slide-wrap");
      if (!s) return;
      current = slides.indexOf(s.querySelector(".slide"));
      setMode("present");
    });
    window.addEventListener("beforeprint", () => { window.__gdgPrev = mode; setMode("print"); });
    window.addEventListener("afterprint", () => setMode(window.__gdgPrev || "present"));
    const ready = document.fonts ? document.fonts.ready : Promise.resolve();
    ready.then(() => {
      window.GDG.ready = true;
      if (params.has("lint")) lint().forEach((w) => console.warn(`[gdg-lint] slide ${w.slide} (${w.type}) ${w.level}: ${w.message}`));
    });
  }

  window.GDG = { lint, setMode, show, toggleTheme, ready: false, rules: RULES, get slides() { return slides; } };
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", init); else init();
})();
