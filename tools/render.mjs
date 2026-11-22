#!/usr/bin/env node
/**
 * Render an HTML slide deck (built with html/gdg-slides.css + gdg-slides.js)
 * to PDF and/or PNG, and print the content-linter warnings.
 *
 *   node tools/render.mjs deck.html                      -> deck.pdf (light) + lint
 *   node tools/render.mjs deck.html --theme dark         -> deck.pdf in dark mode
 *   node tools/render.mjs deck.html --both --out dist    -> dist/deck-light.pdf + dist/deck-dark.pdf
 *   node tools/render.mjs deck.html --png shots          -> shots/deck-light-01.png ...
 *   node tools/render.mjs deck.html --png shots --slides 1,4 --scale 0.5
 *   node tools/render.mjs deck.html --lint               -> only check the slides
 *
 * Options:
 *   --pdf <file>       explicit PDF output path (single theme)
 *   --theme <t>        light (default) | dark
 *   --both             render light and dark
 *   --out <dir>        output directory for PDFs (default: next to the deck)
 *   --png <dir>        also write one PNG per slide into <dir>
 *   --slides <list>    PNG only for these slide numbers (comma separated)
 *   --scale <n>        PNG device scale factor (default 1 -> 1920x1080)
 *   --lint             lint only, write nothing
 *   --strict           exit with code 1 when the linter reports warnings
 *
 * Requires: npm install && npx playwright install chromium
 */
import { chromium } from "playwright";
import { pathToFileURL } from "node:url";
import path from "node:path";
import fs from "node:fs";

const args = process.argv.slice(2);
const opt = { theme: "light", scale: 1 };
const files = [];
for (let i = 0; i < args.length; i++) {
  const a = args[i];
  const next = () => args[++i];
  if (a === "--pdf") opt.pdf = next();
  else if (a === "--theme") opt.theme = next();
  else if (a === "--both") opt.both = true;
  else if (a === "--out") opt.out = next();
  else if (a === "--png" || a === "--screenshots") opt.png = next();
  else if (a === "--slides") opt.slides = next().split(",").map((n) => parseInt(n, 10));
  else if (a === "--scale") opt.scale = parseFloat(next());
  else if (a === "--lint") opt.lintOnly = true;
  else if (a === "--strict") opt.strict = true;
  else if (a === "-h" || a === "--help") { console.log(fs.readFileSync(new URL(import.meta.url)).toString().split("*/")[0]); process.exit(0); }
  else files.push(a);
}
if (!files.length) { console.error("usage: node tools/render.mjs <deck.html> [options]  (--help)"); process.exit(2); }

const deckPath = path.resolve(files[0]);
const base = path.basename(deckPath, path.extname(deckPath));
const outDir = path.resolve(opt.out || path.dirname(deckPath));
const themes = opt.both ? ["light", "dark"] : [opt.theme];

const browser = await chromium.launch();
let warnCount = 0;
try {
  for (const theme of themes) {
    const page = await browser.newPage({ viewport: { width: 1920, height: 1080 }, deviceScaleFactor: opt.scale });
    const consoleErrors = [];
    page.on("pageerror", (e) => consoleErrors.push(String(e)));
    page.on("requestfailed", (r) => consoleErrors.push(`failed to load ${r.url()}`));
    const url = pathToFileURL(deckPath).href + `?mode=print&theme=${theme}`;
    await page.goto(url, { waitUntil: "networkidle" });
    await page.waitForFunction(() => window.GDG && window.GDG.ready, null, { timeout: 30000 });
    await page.evaluate(async () => {
      await Promise.all(Array.from(document.images).map((img) => img.complete ? null : new Promise((r) => { img.onload = img.onerror = r; })));
    });

    // lint once (theme independent)
    if (theme === themes[0]) {
      const warnings = await page.evaluate(() => window.GDG.lint());
      const n = await page.evaluate(() => window.GDG.slides.length);
      console.log(`\n${path.relative(process.cwd(), deckPath)}: ${n} slides`);
      for (const e of consoleErrors) console.log(`  ERROR  ${e}`);
      if (!warnings.length) console.log("  lint: no issues found");
      for (const w of warnings) {
        if (w.level === "warn") warnCount++;
        console.log(`  ${w.level === "warn" ? "WARN " : "info "} slide ${String(w.slide).padStart(2)} (${w.type}): ${w.message}`);
      }
    }
    if (opt.lintOnly) { await page.close(); continue; }

    fs.mkdirSync(outDir, { recursive: true });
    const pdfPath = opt.pdf && !opt.both ? path.resolve(opt.pdf)
      : path.join(outDir, themes.length > 1 || theme !== "light" ? `${base}-${theme}.pdf` : `${base}.pdf`);
    await page.emulateMedia({ media: "print" });
    await page.pdf({ path: pdfPath, preferCSSPageSize: true, printBackground: true });
    console.log(`  wrote ${path.relative(process.cwd(), pdfPath)}`);

    if (opt.png) {
      await page.emulateMedia({ media: "screen" });
      const pngDir = path.resolve(opt.png);
      fs.mkdirSync(pngDir, { recursive: true });
      const handles = await page.$$(".deck .slide");
      const types = await page.evaluate(() => window.GDG.slides.map((s) => s.dataset.type));
      for (let i = 0; i < handles.length; i++) {
        if (opt.slides && !opt.slides.includes(i + 1)) continue;
        const p = path.join(pngDir, `${base}-${theme}-${String(i + 1).padStart(2, "0")}-${types[i]}.png`);
        await handles[i].screenshot({ path: p });
      }
      console.log(`  wrote PNGs to ${path.relative(process.cwd(), pngDir) || "."}`);
    }
    await page.close();
  }
} finally {
  await browser.close();
}
if (opt.strict && warnCount) process.exit(1);
