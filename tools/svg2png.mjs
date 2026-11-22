#!/usr/bin/env node
/**
 * Convert SVG files to PNG (e.g. diagrams for PowerPoint / Impress decks).
 * The bundled Google Sans fonts are made available to the SVG.
 *
 *   node tools/svg2png.mjs diagram.svg [more.svg ...] [--scale 2]
 *   -> diagram.png next to each SVG
 */
import { chromium } from "playwright";
import { pathToFileURL, fileURLToPath } from "node:url";
import path from "node:path";
import fs from "node:fs";

const args = process.argv.slice(2);
let scale = 2;
const files = [];
for (let i = 0; i < args.length; i++) {
  if (args[i] === "--scale") scale = parseFloat(args[++i]);
  else files.push(args[i]);
}
if (!files.length) { console.error("usage: node tools/svg2png.mjs <file.svg> [...] [--scale 2]"); process.exit(2); }

const fontDir = pathToFileURL(path.join(path.dirname(fileURLToPath(import.meta.url)), "..", "assets", "fonts")).href;
const css = [
  ["Google Sans", 400, "GoogleSans-Regular.ttf"], ["Google Sans", 500, "GoogleSans-Medium.ttf"],
  ["Google Sans", 700, "GoogleSans-Bold.ttf"], ["Google Sans Code", 400, "GoogleSansCode-Regular.ttf"],
  ["Google Sans Code", 700, "GoogleSansCode-Bold.ttf"],
].map(([f, w, file]) => `@font-face{font-family:"${f}";font-weight:${w};src:url("${fontDir}/${file}")}`).join("\n");

const browser = await chromium.launch();
try {
  for (const f of files) {
    const svg = fs.readFileSync(f, "utf8");
    const page = await browser.newPage({ deviceScaleFactor: scale });
    await page.setContent(`<!doctype html><html><head><style>${css} html,body{margin:0;background:transparent} svg{display:block}</style></head><body>${svg}</body></html>`);
    await page.evaluate(() => document.fonts.ready);
    const el = await page.$("svg");
    const out = f.replace(/\.svg$/i, ".png");
    await el.screenshot({ path: out, omitBackground: true });
    console.log(`wrote ${out}`);
    await page.close();
  }
} finally {
  await browser.close();
}
