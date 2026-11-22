#!/usr/bin/env python3
"""
gdg_pptx - build editable PowerPoint decks in the GDG UDE design.

The deck is created from templates/gdg-ude-<theme>.potx, so every slide uses a
real layout and stays editable. Text is checked against the content rules
(design/tokens.json -> content_rules) and warnings are printed on save.

Python API
----------
    import sys; sys.path.insert(0, "tools")
    from gdg_pptx import Deck

    deck = Deck(title="Your First Web App with Firebase", author="Max Mustermann",
                date="2026-11-22", theme="light", lang="en")
    deck.title_slide(subtitle="From zero to a deployed app in one evening")
    deck.agenda()                                   # filled from the chapters automatically
    deck.chapter("Why Firebase?", subtitle="What it is and when it fits")
    deck.content("Firebase lets you ship without a server",
                 ["**Hosting** serves your web app", "**Firestore** stores data",
                  ("Authentication", ["Google, GitHub, e-mail"])],
                 image="images/illustration.png", notes="Speaker notes ...")
    deck.full("Pick the right service", table=[["Need", "Service"], ["Sign-in", "Authentication"]])
    deck.two_columns("A vs. B", ("Option A", ["..."]), ("Option B", ["..."]))
    deck.image("Architecture", "diagram.png", caption="Source: ...")
    deck.code("Saving data", 'await addDoc(collection(db, "events"), {...});')
    deck.facts("Numbers", [("3", "services"), ("1", "command"), ("0", "servers")])
    deck.statement("Ship a small thing early - then iterate.", attribution="Rule of thumb")
    deck.references(["Google. Firebase docs. https://firebase.google.com/docs"])
    deck.questions(contact=["Max Mustermann", "max@example.com"], qr="qr.png")
    deck.next_event("Semester Kick-off", "Wed, 3 Dec 2026 · 18:00", "SH-403, Essen",
                    "Talks, workshops and community.", poster="poster.png", qr="qr.png")
    deck.save("my-talk.pptx", pdf=True)            # pdf=True also exports a PDF via LibreOffice

Text formatting inside strings: **bold** and `code`.
A bullet can be a string or (text, [sub-bullets]).

CLI (JSON spec, see docs/slide-types.md)
---------------------------------------
    python tools/gdg_pptx.py deck.json out.pptx [--pdf]
"""
from __future__ import annotations

import copy
import datetime as _dt
import io
import json
import math
import pathlib
import re
import shutil
import subprocess
import sys
import zipfile

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.enum.shapes import MSO_SHAPE, PP_PLACEHOLDER
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

ROOT = pathlib.Path(__file__).resolve().parent.parent
TOKENS = json.loads((ROOT / "design" / "tokens.json").read_text())
RULES = TOKENS["content_rules"]
EMU = 6350
TABLE_STYLE_ID = "{6E2F6D3A-4C1B-4B7E-9C55-1D2A3B4C5D6E}"
COLORS = ["blue", "red", "yellow", "green"]
THEME_SLOT = {"blue": MSO_THEME_COLOR.ACCENT_1, "red": MSO_THEME_COLOR.ACCENT_2,
              "yellow": MSO_THEME_COLOR.ACCENT_3, "green": MSO_THEME_COLOR.ACCENT_4,
              "accent_text": MSO_THEME_COLOR.ACCENT_5, "muted": MSO_THEME_COLOR.TEXT_2,
              "text": MSO_THEME_COLOR.TEXT_1}
SCHEME_VAL = {"blue": "accent1", "red": "accent2", "yellow": "accent3", "green": "accent4"}

I18N = {
    "en": {"agenda": "Agenda", "references": "References", "questions": "Any questions?",
           "thanks": "Thank you for your attention!", "next_event": "Next event", "register": "Register now",
           "months": ["January", "February", "March", "April", "May", "June", "July", "August",
                      "September", "October", "November", "December"]},
    "de": {"agenda": "Agenda", "references": "Quellen", "questions": "Noch Fragen?",
           "thanks": "Vielen Dank für eure Aufmerksamkeit!", "next_event": "Nächstes Event",
           "register": "Jetzt anmelden", "months": ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli",
                                                    "August", "September", "Oktober", "November", "Dezember"]},
}


def px(v: float) -> Emu:
    return Emu(int(round(v * EMU)))


def words(text: str) -> int:
    return len([w for w in re.split(r"\s+", text or "") if re.search(r"\w", w)])


def strip_md(text: str) -> str:
    return re.sub(r"(\*\*|`)", "", text)


class Deck:
    def __init__(self, title: str, author: str = "", date: str = "", theme: str = "light", lang: str = "en",
                 footer: str | None = None, template: str | pathlib.Path | None = None):
        self.title, self.author, self.theme, self.lang = title, author, theme, lang if lang in I18N else "en"
        self.date = self.format_date(date)
        self.footer_text = footer or "  ·  ".join(x for x in (title, author, self.date) if x)
        tpl = pathlib.Path(template) if template else ROOT / "templates" / f"gdg-ude-{theme}.potx"
        self.prs = Presentation(self._as_pptx(tpl))
        self._drop_sample_slides()
        self.layouts = {l.name: l for l in self.prs.slide_layouts}
        self.chapters: list[tuple[str, str | None]] = []
        self._agenda_slides: list[tuple] = []
        self.warnings: list[str] = []
        self._chapter_count = 0

    # ------------------------------------------------------------------ setup
    @staticmethod
    def _as_pptx(path: pathlib.Path) -> io.BytesIO:
        """python-pptx cannot open .potx directly - patch the main content type in memory."""
        src = zipfile.ZipFile(path)
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as dst:
            for item in src.infolist():
                data = src.read(item.filename)
                if item.filename == "[Content_Types].xml":
                    data = data.replace(b"presentationml.template.main+xml", b"presentationml.presentation.main+xml")
                dst.writestr(item, data)
        buf.seek(0)
        return buf

    def _drop_sample_slides(self):
        lst = self.prs.slides._sldIdLst
        for sld in list(lst):
            self.prs.part.drop_rel(sld.rId)
            lst.remove(sld)

    def format_date(self, raw: str) -> str:
        m = re.fullmatch(r"(\d{4})-(\d{2})-(\d{2})", (raw or "").strip())
        if not m:
            return raw or ""
        y, mo, d = int(m[1]), int(m[2]), int(m[3])
        month = I18N[self.lang]["months"][mo - 1]
        return f"{d}. {month} {y}" if self.lang == "de" else f"{d} {month} {y}"

    # ------------------------------------------------------------- utilities
    def _new(self, layout: str, kind: str, footer: bool = True):
        slide = self.prs.slides.add_slide(self.layouts[layout])
        slide._gdg_kind = kind
        if footer:
            self._add_footer(slide)
        return slide

    def _add_footer(self, slide):
        """python-pptx does not copy footer placeholders - clone them from the layout."""
        for ph in slide.slide_layout.placeholders:
            t = ph.placeholder_format.type
            if t not in (PP_PLACEHOLDER.FOOTER, PP_PLACEHOLDER.SLIDE_NUMBER):
                continue
            el = copy.deepcopy(ph._element)
            slide.shapes._spTree.append(el)
            if t == PP_PLACEHOLDER.FOOTER:
                shape = slide.shapes[-1]
                tf = shape.text_frame
                tf.paragraphs[0].runs[0].text = self.footer_text
                for r in tf.paragraphs[0].runs[1:]:
                    r._r.getparent().remove(r._r)
            # strip layout-only formatting so the slide inherits it from the layout
            sp_pr = el.find(qn("p:spPr"))
            for child in list(sp_pr):
                sp_pr.remove(child)

    @staticmethod
    def _ph(slide, idx):
        for ph in slide.placeholders:
            if ph.placeholder_format.idx == idx:
                return ph
        raise KeyError(f"placeholder idx {idx} not found on layout {slide.slide_layout.name}")

    @staticmethod
    def _remove(shape):
        shape._element.getparent().remove(shape._element)

    def _cleanup(self, slide, keep: set[int]):
        """Remove unused content placeholders (footer/number are always kept)."""
        for ph in list(slide.placeholders):
            t = ph.placeholder_format.type
            if t in (PP_PLACEHOLDER.FOOTER, PP_PLACEHOLDER.SLIDE_NUMBER):
                continue
            if ph.placeholder_format.idx not in keep:
                self._remove(ph)

    @staticmethod
    def _runs(paragraph, text: str, color: str | None = None, bold=None, size=None, font=None):
        """Add text with **bold** and `code` markup to a paragraph."""
        for part in re.split(r"(\*\*[^*]+\*\*|`[^`]+`)", text):
            if not part:
                continue
            r = paragraph.add_run()
            if part.startswith("**") and part.endswith("**"):
                r.text = part[2:-2]
                r.font.bold = True
            elif part.startswith("`") and part.endswith("`"):
                r.text = part[1:-1]
                r.font.name = "Google Sans Code"
            else:
                r.text = part
            if bold is not None:
                r.font.bold = bold
            if color:
                if color.startswith("#"):
                    r.font.color.rgb = RGBColor.from_string(color[1:])
                else:
                    r.font.color.theme_color = THEME_SLOT[color]
            if size:
                r.font.size = Pt(size)
            if font:
                r.font.name = font
        return paragraph

    @staticmethod
    def _bullet_color(paragraph, color: str):
        pPr = paragraph._p.get_or_add_pPr()
        for old in pPr.findall(qn("a:buClr")):
            pPr.remove(old)
        bu = etree.SubElement(pPr, qn("a:buClr"))
        etree.SubElement(bu, qn("a:schemeClr"), val=SCHEME_VAL[color])
        # buClr must precede buSzPct/buFont/buChar
        for tag in ("a:buSzPct", "a:buFont", "a:buChar", "a:buNone", "a:tabLst", "a:defRPr"):
            for el in pPr.findall(qn(tag)):
                pPr.remove(el)
                pPr.append(el)

    def _fill_bullets(self, tf, items, cycle=True):
        tf.clear()
        first = True
        n = 0
        for item in items or []:
            text, subs = (item, []) if isinstance(item, str) else (item[0], item[1])
            p = tf.paragraphs[0] if first else tf.add_paragraph()
            first = False
            self._runs(p, text)
            if cycle:
                self._bullet_color(p, COLORS[n % 4])
            n += 1
            for sub in subs or []:
                sp = tf.add_paragraph()
                sp.level = 1
                self._runs(sp, sub)

    def _title(self, slide, text):
        slide.shapes.title.text_frame.text = ""
        self._runs(slide.shapes.title.text_frame.paragraphs[0], text)

    def _notes(self, slide, notes):
        if notes:
            slide.notes_slide.notes_text_frame.text = notes

    @staticmethod
    def _shadow(pic, white_bg=False):
        """Soft drop shadow (and optional white background) for posters and QR codes."""
        spPr = pic._element.spPr
        if white_bg:
            sf = etree.SubElement(spPr, qn("a:solidFill"))
            etree.SubElement(sf, qn("a:srgbClr"), val="FFFFFF")
        eff = etree.fromstring(
            '<a:effectLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
            '<a:outerShdw blurRad="254000" dist="50800" dir="5400000" algn="t" rotWithShape="0">'
            '<a:srgbClr val="202124"><a:alpha val="22000"/></a:srgbClr></a:outerShdw></a:effectLst>')
        spPr.append(eff)

    def _picture(self, slide, idx, path, mode="contain"):
        """Insert a picture into a picture placeholder. mode: contain (no crop) | cover (crop to fill)."""
        ph = self._ph(slide, idx)
        x, y, w, h = ph.left, ph.top, ph.width, ph.height
        pic = ph.insert_picture(str(path))
        if mode == "contain":
            pic.crop_left = pic.crop_right = pic.crop_top = pic.crop_bottom = 0
            iw, ih = pic.image.size
            s = min(w / iw, h / ih)
            nw, nh = int(iw * s), int(ih * s)
            pic.left, pic.top, pic.width, pic.height = x + (w - nw) // 2, y + (h - nh) // 2, nw, nh
        return pic

    def warn(self, slide_no: int, kind: str, msg: str):
        self.warnings.append(f"slide {slide_no:2d} ({kind}): {msg}")

    # ---------------------------------------------------------------- slides
    def title_slide(self, subtitle: str | None = None, title: str | None = None, author: str | None = None,
                    date: str | None = None, notes: str | None = None):
        s = self._new("Title", "title", footer=False)
        self._title(s, title or self.title)
        keep = {0, 13, 14}
        if subtitle:
            self._ph(s, 1).text_frame.text = subtitle
            keep.add(1)
        self._ph(s, 13).text_frame.text = author if author is not None else self.author
        self._ph(s, 14).text_frame.text = self.format_date(date) if date else self.date
        self._cleanup(s, keep)
        self._notes(s, notes)
        return s

    def agenda(self, items: list | None = None, title: str | None = None, current: int | None = None, notes=None):
        """items: list of str or (title, description). None -> filled from chapter() calls on save()."""
        s = self._new("Agenda", "agenda")
        self._title(s, title or I18N[self.lang]["agenda"])
        self._agenda_slides.append((s, items, current))
        self._notes(s, notes)
        return s

    def _fill_agenda(self, s, items, current):
        tf = self._ph(s, 1).text_frame
        tf.clear()
        first = True
        for k, it in enumerate(items):
            name, desc = (it, None) if isinstance(it, str) else (it[0], it[1] if len(it) > 1 else None)
            dim = current is not None and k + 1 != current
            p = tf.paragraphs[0] if first else tf.add_paragraph()
            first = False
            self._runs(p, f"{k + 1:02d}", color="muted" if dim else COLORS[k % 4], bold=True, font="Google Sans")
            self._runs(p, "\t" + name, color="muted" if dim else ("accent_text" if current == k + 1 else None))
            if desc:
                d = tf.add_paragraph()
                d.level = 1
                self._runs(d, desc)

    def chapter(self, title: str, subtitle: str | None = None, image: str | None = None,
                color: str | None = None, number: str | None = None, notes=None):
        self._chapter_count += 1
        color = color or COLORS[(self._chapter_count - 1) % 4]
        s = self._new(f"Chapter {color.capitalize()}", "chapter", footer=False)
        self.chapters.append((title, subtitle))
        self._title(s, title)
        keep = {0, 16}
        if subtitle:
            self._ph(s, 1).text_frame.text = subtitle
            keep.add(1)
        num = self._ph(s, 16)
        num.text_frame.text = number or f"{self._chapter_count:02d}"
        if image:
            self._picture(s, 15, image, mode="cover")
            keep.add(15)
            # move the number into a badge on the edge of the image (as in the HTML version)
            num.left, num.top, num.width, num.height = px(340), px(440), px(200), px(200)
            spPr = num._element.spPr
            geom = etree.SubElement(spPr, qn("a:prstGeom"), prst="ellipse")
            etree.SubElement(geom, qn("a:avLst"))
            sf = etree.SubElement(spPr, qn("a:solidFill"))
            etree.SubElement(sf, qn("a:schemeClr"), val=SCHEME_VAL[color])
            ln = etree.SubElement(spPr, qn("a:ln"), w=str(12 * EMU))
            lf = etree.SubElement(ln, qn("a:solidFill"))
            etree.SubElement(lf, qn("a:schemeClr"), val="bg1")
            num.text_frame.paragraphs[0].line_spacing = Pt(48)
            for r in num.text_frame.paragraphs[0].runs:
                r.font.size = Pt(44)
            # keep the number above the picture
            num._element.getparent().remove(num._element)
            s.shapes._spTree.append(num._element)
        self._cleanup(s, keep)
        self._notes(s, notes)
        return s

    def content(self, title: str, bullets: list | None = None, image: str | None = None, notes=None):
        """Default slide: bullets left, optional picture right."""
        if image:
            s = self._new("Content with Picture", "default")
            self._picture(s, 2, image)
            keep = {0, 1, 2}
        else:
            s = self._new("Content", "default")
            keep = {0, 1}
        self._title(s, title)
        self._fill_bullets(self._ph(s, 1).text_frame, bullets)
        s._gdg_text_width = 860 if image else 1640
        self._cleanup(s, keep)
        self._notes(s, notes)
        return s

    def full(self, title: str, bullets: list | None = None, table: list[list[str]] | None = None,
             col_widths: list[float] | None = None, notes=None):
        """Full-width slide with bullets or a table (first row = header)."""
        s = self._new("Content", "full")
        self._title(s, title)
        body = self._ph(s, 1)
        if table:
            x, y, w = body.left, body.top, body.width
            self._remove(body)
            rows, cols = len(table), max(len(r) for r in table)
            row_h = px(76)
            gf = s.shapes.add_table(rows, cols, x, y, w, row_h * rows)
            tbl = gf.table
            tblPr = tbl._tbl.tblPr
            tblPr.set("firstRow", "1")
            tblPr.set("bandRow", "1")
            sid = tblPr.find(qn("a:tableStyleId"))
            if sid is None:
                sid = etree.SubElement(tblPr, qn("a:tableStyleId"))
            sid.text = TABLE_STYLE_ID
            if col_widths:
                tot = sum(col_widths)
                for i, cw in enumerate(col_widths):
                    tbl.columns[i].width = int(w * cw / tot)
            for r, row in enumerate(table):
                tbl.rows[r].height = row_h
                for c in range(cols):
                    cell = tbl.cell(r, c)
                    cell.margin_left = cell.margin_right = px(24)
                    cell.margin_top = cell.margin_bottom = px(12)
                    tf = cell.text_frame
                    tf.text = ""
                    self._runs(tf.paragraphs[0], str(row[c]) if c < len(row) else "", size=20, bold=True if r == 0 else None)
            s._gdg_table = table
        else:
            self._fill_bullets(body.text_frame, bullets)
            s._gdg_text_width = 1640
        self._notes(s, notes)
        return s

    def two_columns(self, title: str, left: tuple[str, list], right: tuple[str, list], notes=None):
        s = self._new("Two Columns", "two-columns")
        self._title(s, title)
        for (head, items), hidx, bidx in ((left, 13, 1), (right, 14, 2)):
            self._ph(s, hidx).text_frame.text = head
            self._fill_bullets(self._ph(s, bidx).text_frame, items)
        s._gdg_text_width = 750
        self._notes(s, notes)
        return s

    def image(self, title: str, image: str, caption: str | None = None, notes=None):
        s = self._new("Picture", "image")
        self._title(s, title)
        self._picture(s, 1, image)
        keep = {0, 1}
        if caption:
            self._ph(s, 13).text_frame.text = caption
            keep.add(13)
        self._cleanup(s, keep)
        self._notes(s, notes)
        return s

    def code(self, title: str, code: str, notes=None):
        s = self._new("Code", "code")
        self._title(s, title)
        tf = self._ph(s, 1).text_frame
        tf.clear()
        lines = code.strip("\n").split("\n")
        for i, line in enumerate(lines):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            r = p.add_run()
            r.text = line.replace("\t", "    ")
            r.font.name = "Google Sans Code"
        s._gdg_code = lines
        self._notes(s, notes)
        return s

    def facts(self, title: str, facts: list[tuple[str, str]], notes=None):
        """2-4 big numbers with a short label each."""
        if len(facts) == 3:
            s = self._new("Key Facts", "facts")
            self._title(s, title)
            for k, (num, label) in enumerate(facts):
                tf = self._ph(s, 13 + k).text_frame
                tf.clear()
                tf.paragraphs[0].text = num
                p = tf.add_paragraph()
                p.level = 1
                self._runs(p, label)
        else:
            s = self._new("Title Only", "facts")
            self._title(s, title)
            n, gap = len(facts), 48
            w = (1696 - (n - 1) * gap) / n
            for k, (num, label) in enumerate(facts):
                x = 112 + k * (w + gap)
                c = COLORS[k % 4]
                card = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, px(x), px(324), px(w), px(420))
                card.fill.solid()
                card.fill.fore_color.theme_color = MSO_THEME_COLOR.BACKGROUND_2
                card.line.fill.background()
                bar = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, px(x), px(312), px(w), px(12))
                bar.fill.solid()
                bar.fill.fore_color.theme_color = THEME_SLOT[c]
                bar.line.fill.background()
                tb = s.shapes.add_textbox(px(x + 40), px(372), px(w - 80), px(340))
                tf = tb.text_frame
                tf.word_wrap = True
                self._runs(tf.paragraphs[0], num, color=c, bold=True, size=66)
                p = tf.add_paragraph()
                self._runs(p, label, color="text", size=20)
        s._gdg_facts = facts
        self._notes(s, notes)
        return s

    def statement(self, text: str, attribution: str | None = None, notes=None):
        s = self._new("Statement", "statement")
        self._runs(self._ph(s, 1).text_frame.paragraphs[0], text)
        keep = {1}
        if attribution:
            self._ph(s, 13).text_frame.text = attribution
            keep.add(13)
        self._cleanup(s, keep)
        self._notes(s, notes)
        return s

    def title_only(self, title: str, notes=None):
        """Empty canvas with title - add your own shapes/pictures (content area x 112-1808, y 324-940 px)."""
        s = self._new("Title Only", "canvas")
        self._title(s, title)
        self._notes(s, notes)
        return s

    def references(self, refs: list[str], title: str | None = None, notes=None):
        """refs: strings; URLs become links. More than 8 -> split across slides automatically."""
        chunks = [refs[i:i + 8] for i in range(0, len(refs), 8)] or [[]]
        start = 1
        last = None
        for chunk in chunks:
            s = self._new("References", "references")
            self._title(s, title or I18N[self.lang]["references"])
            tf = self._ph(s, 1).text_frame
            tf.clear()
            for i, ref in enumerate(chunk):
                p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                self._runs(p, f"[{start + i}]", color="accent_text", bold=True)
                self._runs(p, "\t")
                for part in re.split(r"(https?://\S+)", ref):
                    if not part:
                        continue
                    r = p.add_run()
                    r.text = part
                    if part.startswith("http"):
                        r.hyperlink.address = part.rstrip(".,;)")
                        r.font.color.theme_color = MSO_THEME_COLOR.ACCENT_5
            start += len(chunk)
            self._notes(s, notes)
            last = s
        return last

    def questions(self, title: str | None = None, subtitle: str | None = None, contact: list[str] | None = None,
                  qr: str | None = None, qr_caption: str | None = None, notes=None):
        s = self._new("Questions", "outro", footer=False)
        t = title or I18N[self.lang]["questions"]
        tf = s.shapes.title.text_frame
        tf.text = ""
        if t.rstrip().endswith(("?", "？")):
            self._runs(tf.paragraphs[0], t.rstrip()[:-1])
            self._runs(tf.paragraphs[0], t.rstrip()[-1], color="blue")
        else:
            self._runs(tf.paragraphs[0], t)
        self._ph(s, 1).text_frame.text = subtitle if subtitle is not None else I18N[self.lang]["thanks"]
        keep = {0, 1}
        if contact:
            ctf = self._ph(s, 13).text_frame
            ctf.clear()
            for i, line in enumerate(contact):
                p = ctf.paragraphs[0] if i == 0 else ctf.add_paragraph()
                if i == 0:
                    self._runs(p, line, size=24, font="Google Sans Medium")
                else:
                    self._runs(p, line)
            keep.add(13)
        if qr:
            self._shadow(self._picture(s, 14, qr), white_bg=True)
            self._ph(s, 15).text_frame.text = qr_caption or ""
            keep |= {14, 15}
        self._cleanup(s, keep)
        self._notes(s, notes)
        return s

    def next_event(self, title: str, date: str, place: str, description: str = "", poster: str | None = None,
                   qr: str | None = None, qr_caption: str | None = None, eyebrow: str | None = None, notes=None):
        s = self._new("Next Event", "next-event", footer=False)
        self._title(s, title)
        self._ph(s, 14).text_frame.text = eyebrow or I18N[self.lang]["next_event"]
        self._ph(s, 15).text_frame.text = date
        self._ph(s, 16).text_frame.text = place
        keep = {0, 14, 15, 16, 19}
        if description:
            self._ph(s, 17).text_frame.text = description
            keep.add(17)
        if poster:
            self._shadow(self._picture(s, 13, poster, mode="cover"))
            keep.add(13)
        if qr:
            self._picture(s, 18, qr)
            keep.add(18)
        self._ph(s, 19).text_frame.text = qr_caption or I18N[self.lang]["register"]
        self._cleanup(s, keep)
        s._gdg_desc = description
        self._notes(s, notes)
        return s

    # ---------------------------------------------------------------- lint
    def lint(self) -> list[str]:
        self.warnings = []
        for n, s in enumerate(self.prs.slides, 1):
            kind = getattr(s, "_gdg_kind", "default")
            title = s.shapes.title.text_frame.text if s.shapes.title is not None else ""
            body_words, bullets = 0, []
            for shp in s.shapes:
                if shp.is_placeholder and shp.placeholder_format.type in (
                        PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE, PP_PLACEHOLDER.FOOTER, PP_PLACEHOLDER.SLIDE_NUMBER):
                    continue
                if shp.has_text_frame:
                    if kind == "code" and shp.is_placeholder and shp.placeholder_format.idx == 1:
                        continue
                    body_words += words(shp.text_frame.text)
                    if shp.is_placeholder and shp.placeholder_format.idx in (1, 2) and kind in ("default", "full", "two-columns"):
                        bullets.append([(p.level, "".join(r.text for r in p.runs)) for p in shp.text_frame.paragraphs])
                if shp.has_table:
                    body_words += sum(words(c.text_frame.text) for r in shp.table.rows for c in r.cells)
            mx = RULES["max_words"].get(kind, 60)
            if body_words > mx:
                self.warn(n, kind, f"too much text: {body_words} words (recommended max. {mx}) - move details to the notes or split the slide")
            has_visual = any(sh.shape_type is not None and sh.shape_type in (13, 19) or getattr(sh, "has_table", False)
                             or (sh.is_placeholder and sh.placeholder_format.type == PP_PLACEHOLDER.PICTURE)
                             for sh in s.shapes)
            if kind in ("default", "full") and not has_visual and body_words < RULES["min_words_without_visual"]:
                self.warn(n, kind, f"very little text ({body_words} words) and no visual - readers of the PDF may miss context")
            if title and kind not in ("outro",):
                tw = words(title)
                if tw > RULES["max_title_words"]:
                    self.warn(n, kind, f"title has {tw} words (max. {RULES['max_title_words']})")
                if kind not in ("title", "chapter", "next-event") and len(title) > 88:
                    self.warn(n, kind, f"title probably wraps to more than 2 lines ({len(title)} characters, max. ~88)")
            # bullets & estimated height
            width = getattr(s, "_gdg_text_width", None)
            for lst in bullets:
                top = [t for lvl, t in lst if lvl == 0]
                if len(top) > RULES["max_bullets"]:
                    self.warn(n, kind, f"{len(top)} bullet points (max. {RULES['max_bullets']})")
                for k, t in enumerate(top):
                    if words(t) > RULES["max_words_per_bullet"]:
                        self.warn(n, kind, f"bullet {k + 1} has {words(t)} words (max. {RULES['max_words_per_bullet']})")
                if any(lvl >= RULES["max_list_depth"] for lvl, _ in lst):
                    self.warn(n, kind, f"list nesting deeper than {RULES['max_list_depth']} levels")
                if width:
                    h = 0.0
                    for lvl, t in lst:
                        size = 48 if lvl == 0 else 40
                        cpl = max(10, (width - (0 if lvl == 0 else 48)) / (size * 0.5))
                        h += math.ceil(max(1, len(t)) / cpl) * size * 1.35 + size * 0.4
                    if h > 616:
                        self.warn(n, kind, f"text probably overflows the slide (estimated {int(h)}px of 616px) - shorten or split")
            if kind == "full" and getattr(s, "_gdg_table", None):
                rows = len(s._gdg_table)
                if rows > 8:
                    self.warn(n, kind, f"table has {rows} rows - max. ~8 rows fit at 20 pt")
            if kind == "code":
                lines = getattr(s, "_gdg_code", [])
                if len(lines) > RULES["max_code_lines"]:
                    self.warn(n, kind, f"{len(lines)} code lines (max. {RULES['max_code_lines']})")
                long = max((len(l) for l in lines), default=0)
                if long > RULES["max_code_line_chars"]:
                    self.warn(n, kind, f"code line with {long} characters (max. {RULES['max_code_line_chars']})")
            if kind == "next-event" and len(getattr(s, "_gdg_desc", "")) > 190:
                self.warn(n, kind, "description longer than ~3 lines - shorten it")
            if kind == "facts" and not 2 <= len(getattr(s, "_gdg_facts", [])) <= 4:
                self.warn(n, kind, "use 2-4 key facts")
        return self.warnings

    # ---------------------------------------------------------------- save
    def save(self, path: str | pathlib.Path, pdf: bool = False, quiet: bool = False) -> pathlib.Path:
        for s, items, current in self._agenda_slides:
            self._fill_agenda(s, items if items is not None else self.chapters, current)
        path = pathlib.Path(path)
        path.parent.mkdir(parents=True, exist_ok=True)
        self.prs.core_properties.title = self.title
        self.prs.core_properties.author = self.author
        self.prs.save(str(path))
        warnings = self.lint()
        if not quiet:
            print(f"wrote {path} ({len(self.prs.slides)} slides)")
            print("  lint: no issues found" if not warnings else "\n".join("  WARN  " + w for w in warnings))
        if pdf:
            out = export_pdf(path)
            if not quiet:
                print(f"wrote {out}" if out else "  (LibreOffice not found - no PDF written)")
        return path

    # ---------------------------------------------------------------- spec
    @classmethod
    def from_spec(cls, spec: dict, base: pathlib.Path | None = None) -> "Deck":
        base = base or pathlib.Path.cwd()
        deck = cls(title=spec["title"], author=spec.get("author", ""), date=spec.get("date", ""),
                   theme=spec.get("theme", "light"), lang=spec.get("lang", "en"), footer=spec.get("footer"))

        def res(p):
            return str((base / p).resolve()) if p else None

        for sl in spec["slides"]:
            sl = dict(sl)
            kind = sl.pop("type")
            for key in ("image", "poster", "qr"):
                if key in sl:
                    sl[key] = res(sl[key])
            if kind == "two-columns":
                sl["left"] = tuple(sl["left"])
                sl["right"] = tuple(sl["right"])
            if kind == "facts":
                sl["facts"] = [tuple(f) for f in sl["facts"]]
            method = {"title": deck.title_slide, "agenda": deck.agenda, "chapter": deck.chapter,
                      "content": deck.content, "default": deck.content, "full": deck.full,
                      "two-columns": deck.two_columns, "image": deck.image, "code": deck.code,
                      "facts": deck.facts, "statement": deck.statement, "canvas": deck.title_only,
                      "references": deck.references, "questions": deck.questions, "outro": deck.questions,
                      "next-event": deck.next_event}[kind]
            method(**sl)
        return deck


def soffice_cmd() -> list[str] | None:
    for exe in ("soffice", "libreoffice"):
        if shutil.which(exe):
            return [exe]
    mac = pathlib.Path("/Applications/LibreOffice.app/Contents/MacOS/soffice")
    if mac.exists():
        return [str(mac)]
    if shutil.which("flatpak") and subprocess.run(["flatpak", "info", "org.libreoffice.LibreOffice"],
                                                  capture_output=True).returncode == 0:
        return ["flatpak", "run", "--command=soffice", "org.libreoffice.LibreOffice"]
    return None


def export_pdf(pptx: pathlib.Path) -> pathlib.Path | None:
    cmd = soffice_cmd()
    if not cmd:
        return None
    pptx = pathlib.Path(pptx).resolve()
    subprocess.run(cmd + ["--headless", "--convert-to", "pdf", "--outdir", str(pptx.parent), str(pptx)],
                   check=True, capture_output=True, timeout=300)
    return pptx.with_suffix(".pdf")


def main(argv=None):
    import argparse
    ap = argparse.ArgumentParser(description="Build a GDG UDE PowerPoint deck from a JSON spec.")
    ap.add_argument("spec", help="deck spec (.json)")
    ap.add_argument("out", help="output .pptx")
    ap.add_argument("--pdf", action="store_true", help="also export a PDF with LibreOffice")
    ap.add_argument("--theme", choices=["light", "dark"], help="override the theme of the spec")
    a = ap.parse_args(argv)
    spec_path = pathlib.Path(a.spec)
    spec = json.loads(spec_path.read_text())
    if a.theme:
        spec["theme"] = a.theme
    deck = Deck.from_spec(spec, base=spec_path.parent)
    deck.save(a.out, pdf=a.pdf)


if __name__ == "__main__":
    main()
