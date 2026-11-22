#!/usr/bin/env python3
"""
Build the PowerPoint and LibreOffice Impress templates of the GDG UDE slide design.

    python tools/build_templates.py            # -> templates/*.potx (+ *.otp if LibreOffice is available)
    python tools/build_templates.py --no-odp   # PowerPoint templates only

The .potx package is written from scratch (theme, slide master, layouts, sample
slides) so that every layout, placeholder, colour and font is native and
editable in PowerPoint. LibreOffice converts the result into .otp templates.

Geometry is given in px on the 1920x1080 design canvas (1 px = 6350 EMU),
matching html/gdg-slides.css and design/tokens.json.
"""
from __future__ import annotations

import argparse
import datetime as _dt
import json
import pathlib
import shutil
import subprocess
import sys
import tempfile
import zipfile
from xml.sax.saxutils import escape

ROOT = pathlib.Path(__file__).resolve().parent.parent
TOKENS = json.loads((ROOT / "design" / "tokens.json").read_text())
EMU = 6350  # per px
SLIDE_W, SLIDE_H = 1920, 1080

NS = ('xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
      'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"')
XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
RT = {
    "master": f"{REL_NS}/slideMaster",
    "layout": f"{REL_NS}/slideLayout",
    "slide": f"{REL_NS}/slide",
    "theme": f"{REL_NS}/theme",
    "image": f"{REL_NS}/image",
    "presProps": f"{REL_NS}/presProps",
    "viewProps": f"{REL_NS}/viewProps",
    "tableStyles": f"{REL_NS}/tableStyles",
}

TABLE_STYLE_ID = "{6E2F6D3A-4C1B-4B7E-9C55-1D2A3B4C5D6E}"

# semantic colour -> theme slot (see theme(): slots are filled per variant)
SCHEME = {
    "blue": "accent1", "red": "accent2", "yellow": "accent3", "green": "accent4",
    "accent_text": "accent5", "line": "accent6",
    "text": "tx1", "bg": "bg1", "muted": "tx2", "surface": "bg2",
}
CHAPTER_COLORS = ["blue", "red", "yellow", "green"]

LOGO_RATIO = 1684 / 208  # trimmed logo width / height


def E(px: float) -> int:
    return int(round(px * EMU))


# ---------------------------------------------------------------------------
# small XML builders
# ---------------------------------------------------------------------------
def clr(c: str, alpha: float | None = None) -> str:
    a = f'<a:alpha val="{int(alpha * 100000)}"/>' if alpha is not None else ""
    if c.startswith("#"):
        return f'<a:srgbClr val="{c[1:].upper()}">{a}</a:srgbClr>'
    return f'<a:schemeClr val="{SCHEME.get(c, c)}">{a}</a:schemeClr>'


def fill(c: str | None, alpha: float | None = None) -> str:
    return "<a:noFill/>" if c is None else f"<a:solidFill>{clr(c, alpha)}</a:solidFill>"


def rpr(sz=None, color=None, b=None, i=None, font=None, cap=None, spc=None, tag="a:rPr", lang=True) -> str:
    attrs = []
    if lang and tag == "a:rPr":
        attrs.append('lang="en-US"')
    if sz:
        attrs.append(f'sz="{int(sz * 100)}"')
    if b is not None:
        attrs.append(f'b="{1 if b else 0}"')
    if i is not None:
        attrs.append(f'i="{1 if i else 0}"')
    if cap:
        attrs.append(f'cap="{cap}"')
    if spc is not None:
        attrs.append(f'spc="{int(spc)}"')
    if tag == "a:rPr":
        attrs.append('dirty="0"')
    inner = fill(color) if color else ""
    if font:
        inner += f'<a:latin typeface="{font}"/><a:ea typeface="{font}"/><a:cs typeface="{font}"/>'
    return f'<{tag} {" ".join(attrs)}>{inner}</{tag}>'


def ppr(lvl=1, algn="l", marL=0, indent=0, ln=None, bef=0, aft=0, bullet=None, defrpr="", tabs=None, tag=None) -> str:
    """Paragraph properties. ln/bef/aft in pt. bullet: None | (char, color, size_pct)."""
    tag = tag or f"a:lvl{lvl}pPr"
    kids = ""
    if ln:
        kids += f'<a:lnSpc><a:spcPts val="{int(ln * 100)}"/></a:lnSpc>'
    kids += f'<a:spcBef><a:spcPts val="{int(bef * 100)}"/></a:spcBef>'
    kids += f'<a:spcAft><a:spcPts val="{int(aft * 100)}"/></a:spcAft>'
    if bullet:
        ch, bc, size = bullet
        kids += f'<a:buClr>{clr(bc)}</a:buClr><a:buSzPct val="{int(size * 1000)}"/><a:buFont typeface="Arial"/><a:buChar char="{ch}"/>'
    else:
        kids += "<a:buNone/>"
    if tabs:
        kids += "<a:tabLst>" + "".join(f'<a:tab pos="{E(t)}" algn="l"/>' for t in tabs) + "</a:tabLst>"
    kids += defrpr
    return f'<{tag} marL="{E(marL)}" indent="{E(indent)}" algn="{algn}" rtl="0">{kids}</{tag}>'


def body_pr(anchor="t", ins=(0, 0, 0, 0), autofit="norm", numCol=1, spcCol=0, wrap="square") -> str:
    l, t, r, b = (E(v) for v in ins)
    col = f' numCol="{numCol}" spcCol="{E(spcCol)}"' if numCol > 1 else ""
    fit = {"norm": "<a:normAutofit/>", "none": "<a:noAutofit/>", "shape": "<a:spAutoFit/>"}[autofit]
    return f'<a:bodyPr vert="horz" wrap="{wrap}" lIns="{l}" tIns="{t}" rIns="{r}" bIns="{b}" rtlCol="0" anchor="{anchor}"{col}>{fit}</a:bodyPr>'


def run(text: str, **kw) -> str:
    return f"<a:r>{rpr(**kw)}<a:t>{escape(text)}</a:t></a:r>"


def para(content: str = "", lvl: int = 0, pprx: str = "", end_sz=None) -> str:
    """content = runs xml. lvl is 0-based like python-pptx."""
    attrs = f' lvl="{lvl}"' if lvl else ""
    p = f"<a:pPr{attrs}>{pprx}</a:pPr>" if (attrs or pprx) else ""
    end = f'<a:endParaRPr lang="en-US"{f" sz={chr(34)}{int(end_sz * 100)}{chr(34)}" if end_sz else ""} dirty="0"/>'
    return f"<a:p>{p}{content}{end}</a:p>"


def xfrm(x, y, w, h, rot=0, tag="a:xfrm") -> str:
    r = f' rot="{int(rot * 60000)}"' if rot else ""
    return f'<{tag}{r}><a:off x="{E(x)}" y="{E(y)}"/><a:ext cx="{E(w)}" cy="{E(h)}"/></{tag}>'


def prst(name="rect", adj: dict | None = None) -> str:
    av = "".join(f'<a:gd name="{k}" fmla="val {v}"/>' for k, v in (adj or {}).items())
    return f'<a:prstGeom prst="{name}"><a:avLst>{av}</a:avLst></a:prstGeom>'


def cust_half_disc(w, h) -> str:
    """Right half of a circle whose centre lies on the left edge of the box."""
    W, H = E(w), E(h)
    return (f'<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/>'
            f'<a:pathLst><a:path w="{W}" h="{H}"><a:moveTo><a:pt x="0" y="0"/></a:moveTo>'
            f'<a:arcTo wR="{W}" hR="{H // 2}" stAng="16200000" swAng="10800000"/><a:close/></a:path></a:pathLst></a:custGeom>')


def cust_block_bl(w, h, r) -> str:
    """Rectangle with a rounded bottom-left corner (poster style)."""
    W, H, R = E(w), E(h), E(r)
    return (f'<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/>'
            f'<a:pathLst><a:path w="{W}" h="{H}"><a:moveTo><a:pt x="0" y="0"/></a:moveTo>'
            f'<a:lnTo><a:pt x="{W}" y="0"/></a:lnTo><a:lnTo><a:pt x="{W}" y="{H}"/></a:lnTo>'
            f'<a:lnTo><a:pt x="{R}" y="{H}"/></a:lnTo>'
            f'<a:arcTo wR="{R}" hR="{R}" stAng="5400000" swAng="5400000"/><a:close/></a:path></a:pathLst></a:custGeom>')


SHADOW = ('<a:effectLst><a:outerShdw blurRad="254000" dist="50800" dir="5400000" algn="t" rotWithShape="0">'
          '<a:srgbClr val="202124"><a:alpha val="22000"/></a:srgbClr></a:outerShdw></a:effectLst>')


# ---------------------------------------------------------------------------
# shape tree
# ---------------------------------------------------------------------------
class Tree:
    """Collects the shapes of one slide / layout / master and its image relations."""

    def __init__(self, images: dict[str, str]):
        self.items: list[str] = []
        self.next_id = 2
        self.images = images          # key -> media file name (shared)
        self.rels: list[tuple[str, str, str]] = []  # (rId, type, target)

    def nid(self) -> int:
        i = self.next_id
        self.next_id += 1
        return i

    def rel(self, rtype: str, target: str) -> str:
        for rid, t, tg in self.rels:
            if t == rtype and tg == target:
                return rid
        rid = f"rId{len(self.rels) + 1}"
        self.rels.append((rid, rtype, target))
        return rid

    def xml(self) -> str:
        return ('<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
                '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
                + "".join(self.items) + "</p:spTree>")

    # -- static drawing ----------------------------------------------------
    def shape(self, x, y, w, h, color="blue", geom=None, name="Decoration", alpha=None, line=None,
              effect="", text="", user=True):
        geom = geom or prst("rect")
        ln = f'<a:ln w="{E(line[1])}">{fill(line[0])}</a:ln>' if line else "<a:ln><a:noFill/></a:ln>"
        tx = text or ""
        self.items.append(
            f'<p:sp><p:nvSpPr><p:cNvPr id="{self.nid()}" name="{escape(name)}"/><p:cNvSpPr/>'
            f'<p:nvPr{" userDrawn=" + chr(34) + "1" + chr(34) if user else ""}/></p:nvSpPr>'
            f'<p:spPr>{xfrm(x, y, w, h)}{geom}{fill(color, alpha)}{ln}{effect}</p:spPr>{tx}</p:sp>')

    def ellipse(self, x, y, d, color, **kw):
        self.shape(x, y, d, d, color, prst("ellipse"), **kw)

    def dots(self, x, y, cols, rows, color, step=40, r=7, name="Dots"):
        gid = self.nid()
        kids = []
        for row in range(rows):
            for col in range(cols):
                cx, cy = x + step / 2 + col * step - r, y + step / 2 + row * step - r
                kids.append(
                    f'<p:sp><p:nvSpPr><p:cNvPr id="{self.nid()}" name="Dot"/><p:cNvSpPr/><p:nvPr userDrawn="1"/></p:nvSpPr>'
                    f'<p:spPr>{xfrm(cx, cy, 2 * r, 2 * r)}{prst("ellipse")}{fill(color)}<a:ln><a:noFill/></a:ln></p:spPr></p:sp>')
        w, h = cols * step, rows * step
        grp = (f'<a:xfrm><a:off x="{E(x)}" y="{E(y)}"/><a:ext cx="{E(w)}" cy="{E(h)}"/>'
               f'<a:chOff x="{E(x)}" y="{E(y)}"/><a:chExt cx="{E(w)}" cy="{E(h)}"/></a:xfrm>')
        self.items.append(
            f'<p:grpSp><p:nvGrpSpPr><p:cNvPr id="{gid}" name="{name}"/><p:cNvGrpSpPr/><p:nvPr userDrawn="1"/></p:nvGrpSpPr>'
            f'<p:grpSpPr>{grp}</p:grpSpPr>{"".join(kids)}</p:grpSp>')

    def rainbow(self, x, y, seg=32, h=8):
        for k, c in enumerate(CHAPTER_COLORS):
            self.shape(x + k * seg, y, seg, h, c, name="Colour bar")

    def textbox(self, x, y, w, h, paras: str, anchor="t", name="Text"):
        tx = f"<p:txBody>{body_pr(anchor, autofit='none')}<a:lstStyle/>{paras}</p:txBody>"
        self.items.append(
            f'<p:sp><p:nvSpPr><p:cNvPr id="{self.nid()}" name="{escape(name)}"/><p:cNvSpPr txBox="1"/><p:nvPr userDrawn="1"/></p:nvSpPr>'
            f'<p:spPr>{xfrm(x, y, w, h)}{prst("rect")}<a:noFill/></p:spPr>{tx}</p:sp>')

    def picture(self, x, y, w, h, key, name="Picture", descr=""):
        rid = self.rel(RT["image"], f"../media/{self.images[key]}")
        self.items.append(
            f'<p:pic><p:nvPicPr><p:cNvPr id="{self.nid()}" name="{escape(name)}" descr="{escape(descr)}"/>'
            f'<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr userDrawn="1"/></p:nvPicPr>'
            f'<p:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>'
            f'<p:spPr>{xfrm(x, y, w, h)}{prst("rect")}</p:spPr></p:pic>')

    # -- placeholders --------------------------------------------------------
    def ph(self, name, phtype=None, idx=None, box=None, bpr="", lst="", prompt="", geom=None, sp_fill=None,
           line=None, effect="", sz=None):
        """Placeholder for layouts / masters. box=(x,y,w,h). prompt = paragraphs xml."""
        t = f' type="{phtype}"' if phtype else ""
        i = f' idx="{idx}"' if idx is not None else ""
        s = f' sz="{sz}"' if sz else ""
        sp = ""
        if box:
            sp += xfrm(*box)
        if geom:
            sp += geom
        if sp_fill:
            sp += fill(sp_fill)
        if line:
            sp += f'<a:ln w="{E(line[1])}">{fill(line[0])}</a:ln>'
        sp += effect
        lock = '<a:spLocks noGrp="1"/>' if phtype != "pic" else '<a:spLocks noGrp="1"/>'
        txb = f"<p:txBody>{bpr or '<a:bodyPr/>'}<a:lstStyle>{lst}</a:lstStyle>{prompt or para()}</p:txBody>"
        self.items.append(
            f'<p:sp><p:nvSpPr><p:cNvPr id="{self.nid()}" name="{escape(name)}"/><p:cNvSpPr>{lock}</p:cNvSpPr>'
            f'<p:nvPr><p:ph{t}{s}{i}/></p:nvPr></p:nvSpPr><p:spPr>{sp}</p:spPr>{txb}</p:sp>')

    def slide_ph(self, name, phtype=None, idx=None, paras=""):
        """Placeholder instance on a slide (inherits geometry and style from the layout)."""
        t = f' type="{phtype}"' if phtype else ""
        i = f' idx="{idx}"' if idx is not None else ""
        self.items.append(
            f'<p:sp><p:nvSpPr><p:cNvPr id="{self.nid()}" name="{escape(name)}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
            f'<p:nvPr><p:ph{t}{i}/></p:nvPr></p:nvSpPr><p:spPr/>'
            f'<p:txBody><a:bodyPr/><a:lstStyle/>{paras or para()}</p:txBody></p:sp>')


# ---------------------------------------------------------------------------
# text styles
# ---------------------------------------------------------------------------
FONT_SANS = "+mn-lt"
FONT_MED = "+mj-lt"            # theme major font = Google Sans Medium
FONT_MONO = "Google Sans Code"


def defrpr(sz, color="text", b=None, font=FONT_SANS, cap=None, spc=None):
    return rpr(sz=sz, color=color, b=b, font=font, cap=cap, spc=spc, tag="a:defRPr", lang=False)


def body_levels(sz1=24, sz2=20, bullets=True):
    """List style for bullet bodies (level 1 + 2, levels 3-9 like level 2)."""
    b1 = ("●", "blue", 60) if bullets else None
    b2 = ("●", "muted", 55) if bullets else None
    m1 = 56 if bullets else 0
    out = ppr(1, marL=m1, indent=-m1 if bullets else 0, ln=sz1 * 1.35, bef=0, aft=sz1 * 0.45, bullet=b1, defrpr=defrpr(sz1))
    for lvl in range(2, 10):
        m = m1 + 48 * (lvl - 1)
        out += ppr(lvl, marL=m, indent=-40 if bullets else 0, ln=sz2 * 1.35, bef=0, aft=sz2 * 0.3,
                   bullet=b2, defrpr=defrpr(sz2, "muted"))
    return out


def simple_style(sz, color="text", b=None, font=FONT_SANS, algn="l", ln=None, cap=None, spc=None, aft=0):
    return ppr(1, algn=algn, ln=ln or sz * 1.3, aft=aft, defrpr=defrpr(sz, color, b, font, cap, spc))


def master_txstyles():
    title = ppr(1, ln=36 * 1.15, defrpr=defrpr(36, "text", font=FONT_MED), tag="a:lvl1pPr")
    other = ppr(1, defrpr=defrpr(18, "text"))
    return (f"<p:txStyles><p:titleStyle>{title}</p:titleStyle>"
            f"<p:bodyStyle>{body_levels()}</p:bodyStyle>"
            f"<p:otherStyle><a:defPPr><a:defRPr lang=\"en-US\"/></a:defPPr>{other}</p:otherStyle></p:txStyles>")


# ---------------------------------------------------------------------------
# geometry constants (px) - keep in sync with gdg-slides.css
# ---------------------------------------------------------------------------
M = 112
TITLE_BOX = (M, 120, 1600, 172)
BODY_TOP = 324
BODY_H = 616
FOOTER_TXT = (484, 1004, 1220, 40)
SLDNUM_BOX = (1824, 1000, 96, 64)


def title_ph(t: Tree, box=TITLE_BOX, prompt_text="Slide title: one clear message, max. 2 lines"):
    t.ph("Title", "title", box=box, bpr=body_pr("t", autofit="norm"),
         prompt=para(run(prompt_text)))


def footer_phs(t: Tree):
    t.ph("Footer", "ftr", idx=11, sz="quarter", box=FOOTER_TXT, bpr=body_pr("ctr", autofit="none", wrap="none"),
         lst=simple_style(12, "muted"), prompt=para(run("Presentation title · Author · Date")))
    t.ph("Slide Number", "sldNum", idx=12, sz="quarter", box=SLDNUM_BOX, bpr=body_pr("ctr", autofit="none", wrap="none"),
         lst=simple_style(14, "#FFFFFF", b=True, algn="ctr"),
         prompt=f'<a:p><a:fld id="{{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}}" type="slidenum">{rpr()}<a:t>‹#›</a:t></a:fld><a:endParaRPr lang="en-US" dirty="0"/></a:p>')


def base_chrome(t: Tree, variant: str, bar=True):
    """Decorations shared by all base slides: corner bubbles, logo, footer, number bubble."""
    t.ellipse(1860, -60, 120, "yellow", name="Bubble yellow")
    t.ellipse(1808, 36, 28, "green", name="Bubble green")
    t.ellipse(1784, 944, 272, "blue", name="Number bubble")
    t.picture(M, 1004, 40 * LOGO_RATIO, 40, f"logo_{variant}", "GDG UDE logo",
              "Google Developer Group University of Duisburg-Essen")
    t.shape(460, 1008, 2, 32, "line", name="Footer separator")
    if bar:
        t.rainbow(M, 88)


def poster_cluster(t: Tree, outro=False):
    """Bubbles of the du+tech poster, used on title and outro slides."""
    t.shape(1620, 0, 300, 300, "yellow", cust_block_bl(300, 300, 150), name="Block yellow")
    t.ellipse(1780, 70, 230, "blue", name="Bubble blue")
    t.dots(1620, 380, 4, 4, "green")
    t.ellipse(1570, 710, 520, "red", name="Bubble red")
    t.ellipse(1370, 910, 400, "yellow", name="Bubble yellow")
    t.ellipse(1414, 210, 36, "red", name="Dot red")
    t.dots(1140, 896, 4, 3, "blue")


# ---------------------------------------------------------------------------
# layouts
# ---------------------------------------------------------------------------
def pic_prompt(text):
    return para(run(text))


def layout_defs(variant: str):
    """Returns list of (name, type, builder(tree))."""
    L = []

    def title_slide(t):
        poster_cluster(t)
        t.picture(M, 96, 72 * LOGO_RATIO, 72, f"logo_{variant}", "GDG UDE logo",
                  "Google Developer Group University of Duisburg-Essen")
        t.ph("Title", "ctrTitle", box=(M, 250, 1300, 280), bpr=body_pr("b"),
             lst=ppr(1, ln=54 * 1.08, defrpr=defrpr(54, "text", b=True)),
             prompt=para(run("Presentation title")))
        t.ph("Subtitle", "subTitle", idx=1, box=(M, 548, 1200, 80), bpr=body_pr("t"),
             lst=simple_style(28, "muted"), prompt=para(run("Optional subtitle")))
        t.rainbow(M, 652)
        t.ph("Author", "body", idx=13, box=(M, 692, 1200, 52), bpr=body_pr("t"),
             lst=simple_style(24, "text", font=FONT_MED), prompt=para(run("Speaker name")))
        t.ph("Date", "body", idx=14, box=(M, 748, 1200, 44), bpr=body_pr("t"),
             lst=simple_style(20, "muted"), prompt=para(run("22 November 2026")))
    L.append(("Title", "title", title_slide))

    def agenda(t):
        base_chrome(t, variant)
        title_ph(t, prompt_text="Agenda")
        lst = (ppr(1, marL=112, indent=-112, ln=28 * 1.2, bef=18, aft=0, tabs=[112],
                   defrpr=defrpr(28, "text", font=FONT_MED))
               + ppr(2, marL=112, indent=0, ln=20 * 1.3, bef=4, defrpr=defrpr(20, "muted")))
        prompt = (para(run("01", b=True, color="blue", font=FONT_SANS) + run("\tChapter title"))
                  + para(run("Optional one-line description"), lvl=1))
        t.ph("Agenda", "body", idx=1, box=(M, BODY_TOP, 1696, BODY_H),
             bpr=body_pr("t", numCol=2, spcCol=96), lst=lst, prompt=prompt)
        footer_phs(t)
    L.append(("Agenda", None, agenda))

    for k, c in enumerate(CHAPTER_COLORS):
        def chapter(t, c=c):
            on = "#1E1E1E" if c == "yellow" else "#FFFFFF"
            t.dots(1920 - M - 160, 96, 4, 4, c)
            t.ellipse(1760, 920, 260, c, alpha=0.18, name="Bubble")
            t.shape(0, 100, 440, 880, c, cust_half_disc(440, 880), name="Chapter half circle")
            t.ph("Chapter image", "pic", idx=15, box=(0, 100, 440, 880), geom=cust_half_disc(440, 880),
                 bpr=body_pr("ctr"), lst=simple_style(16, on, algn="ctr"), prompt=pic_prompt("Optional image"))
            t.ph("Chapter number", "body", idx=16, box=(0, 100, 416, 880), bpr=body_pr("ctr", autofit="none"),
                 lst=ppr(1, algn="ctr", ln=100, defrpr=defrpr(100, on, b=True)), prompt=para(run(f"0{k + 1}")))
            t.ph("Title", "title", box=(640, 270, 1168, 310), bpr=body_pr("b"),
                 lst=ppr(1, ln=54 * 1.08, defrpr=defrpr(54, "text", b=True, font=FONT_SANS)),
                 prompt=para(run("Chapter title")))
            t.shape(640, 620, 128, 8, c, name="Chapter bar")
            t.ph("Subtitle", "body", idx=1, box=(640, 660, 1168, 160), bpr=body_pr("t"),
                 lst=simple_style(28, "muted"), prompt=para(run("Optional subtitle")))
        L.append((f"Chapter {c.capitalize()}", "secHead", chapter))

    def content_image(t):
        base_chrome(t, variant)
        title_ph(t)
        t.ph("Content", None, idx=1, box=(M, BODY_TOP, 916, BODY_H), bpr=body_pr("t"),
             lst=body_levels(), prompt=para(run("Short bullet point (one line)")) + para(run("Second level"), lvl=1))
        t.ph("Picture", "pic", idx=2, box=(1108, BODY_TOP, 700, BODY_H), bpr=body_pr("ctr"),
             lst=simple_style(16, "muted", algn="ctr"), prompt=pic_prompt("Optional picture"))
        footer_phs(t)
    L.append(("Content with Picture", None, content_image))

    def content(t):
        base_chrome(t, variant)
        title_ph(t)
        t.ph("Content", None, idx=1, box=(M, BODY_TOP, 1696, BODY_H), bpr=body_pr("t"),
             lst=body_levels(), prompt=para(run("Text, bullet points, a table or a chart")) + para(run("Second level"), lvl=1))
        footer_phs(t)
    L.append(("Content", "obj", content))

    def two_columns(t):
        base_chrome(t, variant)
        title_ph(t)
        for k, (x, c) in enumerate([(M, "blue"), (1000, "red")]):
            t.shape(x, 318, 808, 8, c, name="Column bar")
            t.ph(f"Column heading {k + 1}", "body", idx=13 + k, box=(x, 342, 808, 60), bpr=body_pr("t"),
                 lst=simple_style(24, "text", font=FONT_MED), prompt=para(run("Column heading")))
            t.ph(f"Column {k + 1}", None, idx=1 + k, box=(x, 414, 808, 520), bpr=body_pr("t"),
                 lst=body_levels(), prompt=para(run("Bullet point")))
        footer_phs(t)
    L.append(("Two Columns", "twoTxTwoObj", two_columns))

    def image(t):
        base_chrome(t, variant)
        title_ph(t)
        t.ph("Picture", "pic", idx=1, box=(M, BODY_TOP, 1696, 548), bpr=body_pr("ctr"),
             lst=simple_style(16, "muted", algn="ctr"), prompt=pic_prompt("Picture, diagram or screenshot"))
        t.ph("Caption", "body", idx=13, box=(M, 886, 1696, 48), bpr=body_pr("t"),
             lst=simple_style(16, "muted"), prompt=para(run("Optional caption or source")))
        footer_phs(t)
    L.append(("Picture", "picTx", image))

    def code(t):
        base_chrome(t, variant)
        title_ph(t)
        # background + bar live on the layout (slide content is always drawn above layout shapes)
        t.shape(M, BODY_TOP, 1696, BODY_H, "surface", prst("roundRect", {"adj": 2600}), name="Code background",
                line=("line", 2))
        t.shape(M, BODY_TOP, 8, BODY_H, "blue", name="Code bar")
        t.ph("Code", "body", idx=1, box=(M, BODY_TOP, 1696, BODY_H),
             bpr=body_pr("t", ins=(48, 32, 40, 32), autofit="norm"),
             lst=ppr(1, ln=18 * 1.45, defrpr=defrpr(18, "text", font=FONT_MONO)),
             prompt=para(run("const answer = 42;", font=FONT_MONO)))
        footer_phs(t)
    L.append(("Code", None, code))

    def facts(t):
        base_chrome(t, variant)
        title_ph(t)
        w = (1696 - 2 * 48) / 3
        for k in range(3):
            x = M + k * (w + 48)
            c = CHAPTER_COLORS[k]
            t.ph(f"Fact {k + 1}", "body", idx=13 + k, box=(x, BODY_TOP, w, 420), sp_fill="surface",
                 bpr=body_pr("t", ins=(40, 48, 40, 40)),
                 lst=ppr(1, ln=66, aft=12, defrpr=defrpr(66, c, b=True)) + ppr(2, ln=20 * 1.3, defrpr=defrpr(20, "text")),
                 prompt=para(run("42")) + para(run("What the number means"), lvl=1))
            t.shape(x, BODY_TOP - 12, w, 12, c, name="Fact bar")
        footer_phs(t)
    L.append(("Key Facts", None, facts))

    def statement(t):
        base_chrome(t, variant, bar=False)
        t.rainbow(240, 330)
        t.ph("Statement", "body", idx=1, box=(240, 370, 1440, 380), bpr=body_pr("t"),
             lst=ppr(1, ln=44 * 1.2, defrpr=defrpr(44, "text", font=FONT_MED)),
             prompt=para(run("One key message or quote, max. ~20 words")))
        t.ph("Attribution", "body", idx=13, box=(240, 770, 1440, 60), bpr=body_pr("t"),
             lst=simple_style(24, "muted"), prompt=para(run("Optional source or attribution")))
        footer_phs(t)
    L.append(("Statement", None, statement))

    def title_only(t):
        base_chrome(t, variant)
        title_ph(t)
        footer_phs(t)
    L.append(("Title Only", "titleOnly", title_only))

    def references(t):
        base_chrome(t, variant)
        title_ph(t, prompt_text="References")
        t.ph("References", "body", idx=1, box=(M, BODY_TOP, 1696, BODY_H), bpr=body_pr("t"),
             lst=ppr(1, marL=72, indent=-72, ln=16 * 1.35, aft=10, tabs=[72], defrpr=defrpr(16, "text")),
             prompt=para(run("[1]", b=True, color="accent_text") + run("\tAuthor. Title. Publisher, Year. URL")))
        footer_phs(t)
    L.append(("References", None, references))

    def outro(t):
        poster_cluster(t)
        t.picture(M, 96, 72 * LOGO_RATIO, 72, f"logo_{variant}", "GDG UDE logo",
                  "Google Developer Group University of Duisburg-Essen")
        t.ph("Title", "title", box=(M, 300, 1180, 220), bpr=body_pr("b", autofit="none"),
             lst=ppr(1, ln=80, defrpr=defrpr(80, "text", b=True, font=FONT_SANS)),
             prompt=para(run("Any questions") + run("?", color="blue")))
        t.ph("Subtitle", "body", idx=1, box=(M, 540, 1150, 70), bpr=body_pr("t"),
             lst=simple_style(28, "muted"), prompt=para(run("Thank you for your attention!")))
        t.rainbow(M, 640)
        t.ph("Contact", "body", idx=13, box=(M, 680, 1150, 230), bpr=body_pr("t"),
             lst=ppr(1, ln=20 * 1.4, defrpr=defrpr(20, "text")),
             prompt=para(run("Speaker name", font=FONT_MED, sz=24)) + para(run("e-mail · GitHub · LinkedIn")) + para(run("Instagram: @gdg_ude")))
        t.ph("QR code", "pic", idx=14, box=(1340, 440, 280, 280),
             bpr=body_pr("ctr"), lst=simple_style(16, "#5F6368", algn="ctr"), prompt=pic_prompt("Optional QR code"))
        t.ph("QR caption", "body", idx=15, box=(1340, 736, 280, 44), bpr=body_pr("t"),
             lst=simple_style(16, "muted", algn="ctr", font=FONT_MED), prompt=para(run("Slides & code")))
    L.append(("Questions", None, outro))

    def next_event(t):
        t.shape(1700, 0, 220, 220, "yellow", cust_block_bl(220, 220, 110), name="Block yellow")
        t.ellipse(1820, 50, 170, "blue", name="Bubble blue")
        t.dots(1500, 72, 3, 3, "green")
        t.ellipse(1680, 860, 360, "red", name="Bubble red")
        t.picture(830, 100, 56 * LOGO_RATIO, 56, f"logo_{variant}", "GDG UDE logo",
                  "Google Developer Group University of Duisburg-Essen")
        t.ph("Poster", "pic", idx=13, box=(M, 100, 622, 880), geom=prst("roundRect", {"adj": 2600}),
             bpr=body_pr("ctr"), lst=simple_style(16, "muted", algn="ctr"),
             prompt=pic_prompt("DIN A3 poster (portrait, 1 : 1.414)"))
        t.ph("Eyebrow", "body", idx=14, box=(830, 212, 978, 40), bpr=body_pr("t", autofit="none"),
             lst=simple_style(16, "accent_text", b=True, cap="all", spc=128), prompt=para(run("Next event")))
        t.ph("Title", "title", box=(830, 256, 978, 190), bpr=body_pr("t"),
             lst=ppr(1, ln=40 * 1.08, defrpr=defrpr(40, "text", b=True, font=FONT_SANS)),
             prompt=para(run("Event title")))
        t.picture(830, 471, 44, 44, "icon_calendar", "Calendar icon")
        t.ph("Date and time", "body", idx=15, box=(894, 468, 914, 50), bpr=body_pr("ctr", autofit="none"),
             lst=simple_style(20, "text"), prompt=para(run("Wednesday, 3 December 2026 · 18:00")))
        t.picture(830, 531, 44, 44, "icon_place", "Location icon")
        t.ph("Location", "body", idx=16, box=(894, 528, 914, 50), bpr=body_pr("ctr", autofit="none"),
             lst=simple_style(20, "text"), prompt=para(run("Room, address")))
        t.ph("Description", "body", idx=17, box=(830, 600, 640, 200), bpr=body_pr("t"),
             lst=simple_style(20, "muted", ln=20 * 1.35), prompt=para(run("Short description, max. 3 lines (~30 words).")))
        t.textbox(826, 860, 720, 90, para(
            run("DU ", sz=40, b=True, color="text", font=FONT_SANS) + run("+", sz=40, b=True, color="blue", font=FONT_SANS)
            + run(" TECH = ", sz=40, b=True, color="text", font=FONT_SANS) + run("?", sz=40, b=True, color="blue", font=FONT_SANS),
            pprx=f'<a:lnSpc><a:spcPts val="4400"/></a:lnSpc>'), anchor="b", name="Tagline DU + TECH = ?")
        t.shape(1558, 690, 250, 290, "#FFFFFF", prst("roundRect", {"adj": 9600}), name="QR card", effect=SHADOW)
        t.ph("QR code", "pic", idx=18, box=(1578, 710, 210, 210), bpr=body_pr("ctr"),
             lst=simple_style(14, "#5F6368", algn="ctr"), prompt=pic_prompt("QR code"))
        t.ph("QR caption", "body", idx=19, box=(1558, 926, 250, 40), bpr=body_pr("t", autofit="none"),
             lst=simple_style(16, "#1E1E1E", b=True, algn="ctr"), prompt=para(run("Register now")))
    L.append(("Next Event", None, next_event))

    def blank(t):
        base_chrome(t, variant, bar=False)
        footer_phs(t)
    L.append(("Blank with Footer", None, blank))
    return L


# ---------------------------------------------------------------------------
# sample slides (instructions for manual use)
# ---------------------------------------------------------------------------
FOOTER_SAMPLE = "Presentation title  ·  Speaker name  ·  22 November 2026"


def samples():
    """(layout name, [(ph type, idx, paras)], footer?)"""
    P = lambda *runs_: para("".join(runs_))  # noqa: E731
    R = run
    S = []
    S.append(("Title", [("ctrTitle", None, P(R("Presentation title"))),
                        ("subTitle", 1, P(R("Optional subtitle"))),
                        (None, 13, P(R("Speaker name"))),
                        (None, 14, P(R("22 November 2026")))], False))
    S.append(("Agenda", [("title", None, P(R("Agenda"))),
                         (None, 1, "".join(
                             para(R(f"0{i + 1}", b=True, color=c) + R("\t" + name)) + (para(R(desc), lvl=1) if desc else "")
                             for i, (c, name, desc) in enumerate([
                                 ("blue", "First chapter", "One-line description"),
                                 ("red", "Second chapter", "One-line description"),
                                 ("yellow", "Third chapter", ""),
                                 ("green", "Q&A", "")])))], True))
    S.append(("Chapter Blue", [("title", None, P(R("Chapter title"))), (None, 1, P(R("Optional subtitle"))),
                               (None, 16, P(R("01")))], False))
    S.append(("Content with Picture", [("title", None, P(R("Slide titles state the message of the slide"))),
                                       (None, 1, "".join([
                                           para(R("3–6 short bullet points per slide")),
                                           para(R("One line each, max. ~12 words")),
                                           para(R("Details go into the speaker notes"), lvl=1),
                                           para(R("Optional picture on the right"))]))], True))
    S.append(("Content", [("title", None, P(R("Full-width content: text, table or chart"))),
                          (None, 1, "".join([para(R("Use this layout for tables, charts and longer bullet lists")),
                                             para(R("Body text 24 pt, second level 20 pt, never below 16 pt"))]))], True))
    S.append(("Two Columns", [("title", None, P(R("Compare two things side by side"))),
                              (None, 13, P(R("Option A"))), (None, 1, P(R("Advantage")) + P(R("Another point"))),
                              (None, 14, P(R("Option B"))), (None, 2, P(R("Advantage")) + P(R("Another point")))], True))
    S.append(("Picture", [("title", None, P(R("A picture says more than bullet points"))),
                          (None, 13, P(R("Caption or source of the picture")))], True))
    S.append(("Code", [("title", None, P(R("Show only the code that matters"))),
                       (None, 1, "".join(para(R(line, font=FONT_MONO)) for line in [
                           "// Google Sans Code, 18 pt, max. ~15 lines",
                           "const greeting = \"Hello GDG UDE!\";",
                           "console.log(greeting);"]))], True))
    S.append(("Key Facts", [("title", None, P(R("Three numbers people should remember")))] + [
        (None, 13 + k, para(R(n)) + para(R(lbl), lvl=1)) for k, (n, lbl) in
        enumerate([("3", "short label for the first number"), ("1", "short label for the second number"),
                   ("0", "short label for the third number")])], True))
    S.append(("Statement", [(None, 1, P(R("One strong statement per slide is easier to remember."))),
                            (None, 13, P(R("Optional attribution")))], True))
    S.append(("References", [("title", None, P(R("References"))),
                             (None, 1, "".join(para(R(f"[{i}]", b=True, color="accent_text") + R("\t" + t)) for i, t in enumerate([
                                 "Author, A. Title of the article. Journal, Year. https://example.com",
                                 "Organisation. Title of the web page. https://example.com"], 1)))], True))
    S.append(("Questions", [("title", None, P(R("Any questions") + R("?", color="blue"))),
                            (None, 1, P(R("Thank you for your attention!"))),
                            (None, 13, para(R("Speaker name", font=FONT_MED, sz=24)) + P(R("name@example.com")) + P(R("Instagram: @gdg_ude"))),
                            (None, 15, P(R("Slides & code")))], False))
    S.append(("Next Event", [(None, 14, P(R("Next event"))), ("title", None, P(R("Event title"))),
                             (None, 15, P(R("Wednesday, 3 December 2026 · 18:00"))),
                             (None, 16, P(R("SH-403, Schützenbahn 70, Essen"))),
                             (None, 17, P(R("Short description of the event, max. three lines."))),
                             (None, 19, P(R("Register now")))], False))
    return S


# ---------------------------------------------------------------------------
# package parts
# ---------------------------------------------------------------------------
def theme_xml(variant: str) -> str:
    c = TOKENS["colors"]
    v = c[variant]
    b = c["brand"]
    # Text/background slots: dk1 = text, lt1 = background (also in dark mode, so that the
    # standard colour mapping works in every application).
    slots = {
        "dk1": v["text"], "lt1": v["background"], "dk2": v["text_muted"], "lt2": v["surface"],
        "accent1": b["blue"], "accent2": b["red"], "accent3": b["yellow"], "accent4": b["green"],
        "accent5": v["accent_text"], "accent6": v["line"],
        "hlink": v["accent_text"], "folHlink": "#A142F4" if variant == "light" else "#D7AEFB",
    }
    cs = "".join(f'<a:{k}><a:srgbClr val="{h[1:].upper()}"/></a:{k}>' for k, h in slots.items())
    fonts = TOKENS["fonts"]
    font = lambda f: f'<a:latin typeface="{f}"/><a:ea typeface=""/><a:cs typeface=""/>'  # noqa: E731
    fmt = ('<a:fmtScheme name="GDG UDE"><a:fillStyleLst>'
           '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
           '<a:solidFill><a:schemeClr val="phClr"><a:tint val="60000"/></a:schemeClr></a:solidFill>'
           '<a:solidFill><a:schemeClr val="phClr"><a:shade val="80000"/></a:schemeClr></a:solidFill>'
           '</a:fillStyleLst><a:lnStyleLst>'
           '<a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>'
           '<a:ln w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>'
           '<a:ln w="38100" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>'
           '</a:lnStyleLst><a:effectStyleLst>'
           '<a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle>'
           '<a:effectStyle><a:effectLst><a:outerShdw blurRad="254000" dist="50800" dir="5400000" algn="t" rotWithShape="0">'
           '<a:srgbClr val="000000"><a:alpha val="22000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle>'
           '</a:effectStyleLst><a:bgFillStyleLst>'
           '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
           '<a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/></a:schemeClr></a:solidFill>'
           '<a:solidFill><a:schemeClr val="phClr"><a:shade val="90000"/></a:schemeClr></a:solidFill>'
           '</a:bgFillStyleLst></a:fmtScheme>')
    name = f"GDG UDE {variant.capitalize()}"
    return (XML_HEAD + f'<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="{name}">'
            f'<a:themeElements><a:clrScheme name="{name}">{cs}</a:clrScheme>'
            f'<a:fontScheme name="GDG UDE"><a:majorFont>{font(fonts["sans_medium"])}</a:majorFont>'
            f'<a:minorFont>{font(fonts["sans"])}</a:minorFont></a:fontScheme>{fmt}</a:themeElements>'
            '<a:objectDefaults/><a:extraClrSchemeLst/></a:theme>')


def table_styles_xml() -> str:
    none = "<a:ln><a:noFill/></a:ln>"
    thin = '<a:ln w="12700"><a:solidFill><a:schemeClr val="accent6"/></a:solidFill></a:ln>'
    thick = '<a:ln w="38100"><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:ln>'
    txt = '<a:tcTxStyle{b}><a:fontRef idx="minor"><a:prstClr val="black"/></a:fontRef><a:schemeClr val="tx1"/></a:tcTxStyle>'
    whole = (f'<a:wholeTbl>{txt.format(b="")}<a:tcStyle><a:tcBdr>'
             f'<a:left>{none}</a:left><a:right>{none}</a:right><a:top>{none}</a:top><a:bottom>{thin}</a:bottom>'
             f'<a:insideH>{thin}</a:insideH><a:insideV>{none}</a:insideV></a:tcBdr>'
             f'<a:fill><a:noFill/></a:fill></a:tcStyle></a:wholeTbl>')
    band2 = '<a:band2H><a:tcStyle><a:tcBdr/><a:fill><a:solidFill><a:schemeClr val="bg2"/></a:solidFill></a:fill></a:tcStyle></a:band2H>'
    first = (f'<a:firstRow>{txt.format(b=" b=" + chr(34) + "on" + chr(34))}<a:tcStyle><a:tcBdr><a:bottom>{thick}</a:bottom></a:tcBdr>'
             f'<a:fill><a:noFill/></a:fill></a:tcStyle></a:firstRow>')
    return (XML_HEAD + f'<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{TABLE_STYLE_ID}">'
            f'<a:tblStyle styleId="{TABLE_STYLE_ID}" styleName="GDG UDE">{whole}{band2}{first}</a:tblStyle></a:tblStyleLst>')


def rels_xml(rels) -> str:
    body = "".join(f'<Relationship Id="{rid}" Type="{t}" Target="{escape(tg)}"/>' for rid, t, tg in rels)
    return XML_HEAD + f'<Relationships xmlns="{PKG_REL_NS}">{body}</Relationships>'


def build_package(variant: str, out: pathlib.Path, template=True, with_samples=True):
    images = {
        f"logo_{variant}": "image1.png",
        "icon_calendar": "image2.png",
        "icon_place": "image3.png",
    }
    image_src = {
        "image1.png": ROOT / TOKENS["colors"][variant]["logo"],
        "image2.png": ROOT / "assets/icons/calendar.png",
        "image3.png": ROOT / "assets/icons/place.png",
    }
    parts: dict[str, str | bytes] = {}

    # layouts
    layout_names = []
    for n, (name, ltype, builder) in enumerate(layout_defs(variant), 1):
        t = Tree(images)
        t.rel(RT["master"], "../slideMasters/slideMaster1.xml")
        builder(t)
        typ = f' type="{ltype}"' if ltype else ""
        parts[f"ppt/slideLayouts/slideLayout{n}.xml"] = (
            XML_HEAD + f'<p:sldLayout {NS}{typ} preserve="1" userDrawn="1">'
            f'<p:cSld name="{escape(name)}">{t.xml()}</p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>')
        parts[f"ppt/slideLayouts/_rels/slideLayout{n}.xml.rels"] = rels_xml(t.rels)
        layout_names.append(name)

    # master
    mt = Tree(images)
    title_ph(mt)
    mt.ph("Text", "body", idx=1, box=(M, BODY_TOP, 1696, BODY_H), bpr=body_pr("t"),
          prompt=para(run("First level")) + para(run("Second level"), lvl=1))
    footer_phs(mt)
    lid = "".join(f'<p:sldLayoutId id="{2147483649 + i}" r:id="rId{i + 1}"/>' for i in range(len(layout_names)))
    mrels = [(f"rId{i + 1}", RT["layout"], f"../slideLayouts/slideLayout{i + 1}.xml") for i in range(len(layout_names))]
    mrels.append((f"rId{len(layout_names) + 1}", RT["theme"], "../theme/theme1.xml"))
    parts["ppt/slideMasters/slideMaster1.xml"] = (
        XML_HEAD + f'<p:sldMaster {NS}><p:cSld><p:bg><p:bgPr>{fill("bg")}<a:effectLst/></p:bgPr></p:bg>{mt.xml()}</p:cSld>'
        '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" '
        'accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>'
        f'<p:sldLayoutIdLst>{lid}</p:sldLayoutIdLst><p:hf hdr="0" dt="0"/>{master_txstyles()}</p:sldMaster>')
    parts["ppt/slideMasters/_rels/slideMaster1.xml.rels"] = rels_xml(mrels)
    parts["ppt/theme/theme1.xml"] = theme_xml(variant)

    # sample slides
    slide_ids = []
    if with_samples:
        for n, (lname, phs, footer) in enumerate(samples(), 1):
            t = Tree(images)
            t.rel(RT["layout"], f"../slideLayouts/slideLayout{layout_names.index(lname) + 1}.xml")
            for phtype, idx, paras in phs:
                t.slide_ph(f"Placeholder {idx or phtype}", phtype, idx, paras)
            if footer:
                t.slide_ph("Footer", "ftr", 11, para(run(FOOTER_SAMPLE)))
                t.slide_ph("Slide Number", "sldNum", 12,
                           f'<a:p><a:fld id="{{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}}" type="slidenum">{rpr()}<a:t>{n}</a:t></a:fld></a:p>')
            parts[f"ppt/slides/slide{n}.xml"] = (
                XML_HEAD + f'<p:sld {NS}><p:cSld>{t.xml()}</p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>')
            parts[f"ppt/slides/_rels/slide{n}.xml.rels"] = rels_xml(t.rels)
            slide_ids.append(n)

    # presentation
    prels = [("rId1", RT["master"], "slideMasters/slideMaster1.xml")]
    sld = ""
    for n in slide_ids:
        rid = f"rId{n + 1}"
        prels.append((rid, RT["slide"], f"slides/slide{n}.xml"))
        sld += f'<p:sldId id="{255 + n}" r:id="{rid}"/>'
    k = len(prels)
    prels += [(f"rId{k + 1}", RT["presProps"], "presProps.xml"), (f"rId{k + 2}", RT["viewProps"], "viewProps.xml"),
              (f"rId{k + 3}", RT["theme"], "theme/theme1.xml"), (f"rId{k + 4}", RT["tableStyles"], "tableStyles.xml")]
    deftxt = "".join(ppr(i, defrpr=defrpr(18, "text"), tag=f"a:lvl{i}pPr") for i in range(1, 10))
    parts["ppt/presentation.xml"] = (
        XML_HEAD + f'<p:presentation {NS} saveSubsetFonts="1"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>'
        + (f"<p:sldIdLst>{sld}</p:sldIdLst>" if sld else "")
        + f'<p:sldSz cx="{E(SLIDE_W)}" cy="{E(SLIDE_H)}"/><p:notesSz cx="6858000" cy="9144000"/>'
        f'<p:defaultTextStyle><a:defPPr><a:defRPr lang="en-US"/></a:defPPr>{deftxt}</p:defaultTextStyle></p:presentation>')
    parts["ppt/_rels/presentation.xml.rels"] = rels_xml(prels)
    parts["ppt/presProps.xml"] = XML_HEAD + f"<p:presentationPr {NS}/>"
    guides = "".join(f'<p:guide orient="{o}" pos="{int(px * 4)}"/>' for o, px in
                     [("vert", M), ("vert", 1808), ("horz", BODY_TOP), ("horz", BODY_TOP + BODY_H)])
    parts["ppt/viewProps.xml"] = (
        XML_HEAD + f'<p:viewPr {NS}><p:normalViewPr><p:restoredLeft sz="15620"/><p:restoredTop sz="94660"/></p:normalViewPr>'
        '<p:slideViewPr><p:cSldViewPr snapToGrid="0"><p:cViewPr varScale="1"><p:scale><a:sx n="100" d="100"/><a:sy n="100" d="100"/></p:scale>'
        f'<p:origin x="0" y="0"/></p:cViewPr><p:guideLst>{guides}</p:guideLst></p:cSldViewPr></p:slideViewPr>'
        '<p:gridSpacing cx="76200" cy="76200"/></p:viewPr>')
    parts["ppt/tableStyles.xml"] = table_styles_xml()

    now = _dt.datetime.now(_dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    parts["docProps/core.xml"] = (
        XML_HEAD + '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
        'xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        f'<dc:title>GDG UDE slide template ({variant})</dc:title><dc:creator>GDG on Campus University of Duisburg-Essen</dc:creator>'
        f'<dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified>'
        '</cp:coreProperties>')
    parts["docProps/app.xml"] = (
        XML_HEAD + '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
        'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
        f'<Application>GDG UDE build_templates.py</Application><Slides>{len(slide_ids)}</Slides><PresentationFormat>Widescreen</PresentationFormat></Properties>')
    parts["_rels/.rels"] = rels_xml([
        ("rId1", f"{REL_NS}/officeDocument", "ppt/presentation.xml"),
        ("rId2", "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties", "docProps/core.xml"),
        ("rId3", f"{REL_NS}/extended-properties", "docProps/app.xml")])

    main = ("application/vnd.openxmlformats-officedocument.presentationml.template.main+xml" if template
            else "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml")
    pml = "application/vnd.openxmlformats-officedocument.presentationml"
    ov = [("/ppt/presentation.xml", main), ("/ppt/slideMasters/slideMaster1.xml", f"{pml}.slideMaster+xml"),
          ("/ppt/theme/theme1.xml", "application/vnd.openxmlformats-officedocument.theme+xml"),
          ("/ppt/presProps.xml", f"{pml}.presProps+xml"), ("/ppt/viewProps.xml", f"{pml}.viewProps+xml"),
          ("/ppt/tableStyles.xml", f"{pml}.tableStyles+xml"),
          ("/docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml"),
          ("/docProps/app.xml", "application/vnd.openxmlformats-officedocument.extended-properties+xml")]
    ov += [(f"/ppt/slideLayouts/slideLayout{i + 1}.xml", f"{pml}.slideLayout+xml") for i in range(len(layout_names))]
    ov += [(f"/ppt/slides/slide{n}.xml", f"{pml}.slide+xml") for n in slide_ids]
    ct = (XML_HEAD + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>'
          + "".join(f'<Override PartName="{p}" ContentType="{c}"/>' for p, c in ov) + "</Types>")

    out.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        for name in ["_rels/.rels"] + sorted(k for k in parts if k != "_rels/.rels"):
            z.writestr(name, parts[name])
        for media, src in image_src.items():
            z.write(src, f"ppt/media/{media}")
    return layout_names


# ---------------------------------------------------------------------------
# LibreOffice
# ---------------------------------------------------------------------------
def soffice_cmd() -> list[str] | None:
    for exe in ("soffice", "libreoffice"):
        if shutil.which(exe):
            return [exe]
    if shutil.which("flatpak"):
        r = subprocess.run(["flatpak", "info", "org.libreoffice.LibreOffice"], capture_output=True)
        if r.returncode == 0:
            return ["flatpak", "run", "--command=soffice", "org.libreoffice.LibreOffice"]
    return None


def convert(src: pathlib.Path, fmt: str, outdir: pathlib.Path) -> pathlib.Path:
    cmd = soffice_cmd()
    if not cmd:
        raise RuntimeError("LibreOffice not found")
    outdir.mkdir(parents=True, exist_ok=True)
    subprocess.run(cmd + ["--headless", "--convert-to", fmt, "--outdir", str(outdir), str(src)],
                   check=True, capture_output=True, timeout=300)
    return outdir / (src.stem + "." + fmt.split(":")[0])


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--no-odp", action="store_true", help="skip the LibreOffice Impress templates")
    ap.add_argument("--out", default=str(ROOT / "templates"))
    args = ap.parse_args()
    out = pathlib.Path(args.out)
    for variant in ("light", "dark"):
        potx = out / f"gdg-ude-{variant}.potx"
        names = build_package(variant, potx, template=True)
        print(f"wrote {potx.relative_to(ROOT)} ({len(names)} layouts)")
        if args.no_odp:
            continue
        if not soffice_cmd():
            print("LibreOffice not found - skipping .otp", file=sys.stderr)
            continue
        # LibreOffice imports .pptx more reliably than .potx: build a .pptx twin first.
        build_dir = ROOT / "build"
        tmp = build_dir / f"gdg-ude-{variant}.pptx"
        build_package(variant, tmp, template=False)
        otp = convert(tmp, "otp:impress8_template", build_dir)
        shutil.move(otp, out / f"gdg-ude-{variant}.otp")
        print(f"wrote templates/gdg-ude-{variant}.otp")


if __name__ == "__main__":
    main()
