#!/usr/bin/env python3
"""Build the UUI x Meta strategy deck (37 slides) as a PPTX.

Design coordinates are expressed in pixels on a 1280x720 canvas and mapped
1:1 to a 13.333 x 7.5 inch widescreen slide (96 px per inch).
"""

import os
from lxml import etree
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
A = lambda *p: os.path.join(HERE, *p)

# ---------------------------------------------------------------- palette
TEXT = RGBColor(0xF4, 0xF6, 0xFB)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
MUTED = RGBColor(0x98, 0xA3, 0xBA)
MUTED2 = RGBColor(0x6F, 0x7B, 0x93)
RED = RGBColor(0xCE, 0x20, 0x20)
REDB = RGBColor(0xFF, 0x4D, 0x5E)
BLUE = RGBColor(0x0A, 0x7C, 0xFF)
BLUEB = RGBColor(0x38, 0xB6, 0xFF)
GREEN = RGBColor(0x25, 0xD3, 0x66)
AMBER = RGBColor(0xFF, 0xC2, 0x4B)
LINEC = RGBColor(0xFF, 0xFF, 0xFF)

FONT = "Arial"
MONO = "Menlo"

# ---------------------------------------------------------------- geometry
def IN(px):
    return Inches(px / 96.0)

def FS(px):
    """CSS px -> points."""
    return Pt(px * 0.75)

# ---------------------------------------------------------------- low level
def _set_alpha(color_parent, alpha):
    """Add/replace <a:alpha> under an <a:srgbClr> element."""
    if alpha is None or alpha >= 1:
        return
    srgb = color_parent.find(qn("a:srgbClr"))
    if srgb is None:
        return
    node = srgb.find(qn("a:alpha"))
    if node is None:
        node = etree.SubElement(srgb, qn("a:alpha"))
    node.set("val", str(int(round(alpha * 100000))))


def fill(shape, rgb, alpha=1.0):
    shape.fill.solid()
    shape.fill.fore_color.rgb = rgb
    if alpha < 1:
        spPr = shape._element.spPr
        sf = spPr.find(qn("a:solidFill"))
        _set_alpha(sf, alpha)


def stroke(shape, rgb=None, alpha=1.0, width=0.75):
    if rgb is None:
        shape.line.fill.background()
        return
    shape.line.color.rgb = rgb
    shape.line.width = Pt(width)
    if alpha < 1:
        spPr = shape._element.spPr
        ln = spPr.find(qn("a:ln"))
        sf = ln.find(qn("a:solidFill"))
        _set_alpha(sf, alpha)


def no_shadow(shape):
    try:
        shape.shadow.inherit = False
    except Exception:
        pass


def shape(s, x, y, w, h, kind=MSO_SHAPE.ROUNDED_RECTANGLE,
          bg=None, bg_alpha=1.0, border=None, border_alpha=1.0, bw=0.75,
          radius=None):
    shp = s.shapes.add_shape(kind, IN(x), IN(y), IN(w), IN(h))
    no_shadow(shp)
    if kind == MSO_SHAPE.ROUNDED_RECTANGLE and radius is not None:
        shp.adjustments[0] = max(0.0, min(0.5, radius / max(1.0, min(w, h))))
    if bg is None:
        shp.fill.background()
    else:
        fill(shp, bg, bg_alpha)
    if border is None:
        shp.line.fill.background()
    else:
        stroke(shp, border, border_alpha, bw)
    shp.text_frame.text = ""
    return shp


def card(s, x, y, w, h, tone="plain", radius=16):
    if tone == "red":
        return shape(s, x, y, w, h, bg=RED, bg_alpha=0.13, border=REDB,
                     border_alpha=0.28, radius=radius)
    if tone == "blue":
        return shape(s, x, y, w, h, bg=BLUE, bg_alpha=0.13, border=BLUEB,
                     border_alpha=0.28, radius=radius)
    return shape(s, x, y, w, h, bg=WHITE, bg_alpha=0.045, border=WHITE,
                 border_alpha=0.10, radius=radius)


ALIGN = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}
ANCHOR = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM}


def _runs(p, content, size, color, bold, italic, font, spc):
    if isinstance(content, str):
        content = [(content, {})]
    for raw in content:
        if isinstance(raw, str):
            txt, opt = raw, {}
        else:
            txt, opt = raw
        r = p.add_run()
        r.text = txt
        f = r.font
        f.name = opt.get("font", font)
        f.size = Pt(opt.get("size", size))
        f.bold = opt.get("bold", bold)
        f.italic = opt.get("italic", italic)
        f.color.rgb = opt.get("color", color)
        sp = opt.get("spc", spc)
        if sp:
            r._r.get_or_add_rPr().set("spc", str(int(sp * 100)))


def text(s, x, y, w, h, content, size=10, color=MUTED, bold=False,
         italic=False, align="l", anchor="t", line_spacing=1.18,
         spc=0, font=FONT, margins=(0, 0, 0, 0), wrap=True):
    tb = s.shapes.add_textbox(IN(x), IN(y), IN(w), IN(h))
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.margin_left = IN(margins[0])
    tf.margin_top = IN(margins[1])
    tf.margin_right = IN(margins[2])
    tf.margin_bottom = IN(margins[3])
    tf.vertical_anchor = ANCHOR[anchor]
    p = tf.paragraphs[0]
    p.alignment = ALIGN[align]
    p.line_spacing = line_spacing
    p.space_before = Pt(0)
    p.space_after = Pt(0)
    _runs(p, content, size, color, bold, italic, font, spc)
    return tb


def paragraph(tf, content, size=10, color=MUTED, bold=False, italic=False,
              align="l", line_spacing=1.18, space_before=0, space_after=0,
              spc=0, font=FONT):
    p = tf.add_paragraph()
    p.alignment = ALIGN[align]
    p.line_spacing = line_spacing
    p.space_before = Pt(space_before)
    p.space_after = Pt(space_after)
    _runs(p, content, size, color, bold, italic, font, spc)
    return p


def textbox(s, x, y, w, h, anchor="t", margins=(0, 0, 0, 0)):
    tb = s.shapes.add_textbox(IN(x), IN(y), IN(w), IN(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = IN(margins[0])
    tf.margin_top = IN(margins[1])
    tf.margin_right = IN(margins[2])
    tf.margin_bottom = IN(margins[3])
    tf.vertical_anchor = ANCHOR[anchor]
    return tf


def hline(s, x, y, w, color=WHITE, alpha=0.10, width=0.75):
    ln = s.shapes.add_connector(1, IN(x), IN(y), IN(x + w), IN(y))
    ln.line.color.rgb = color
    ln.line.width = Pt(width)
    if alpha < 1:
        spPr = ln._element.spPr
        lnel = spPr.find(qn("a:ln"))
        _set_alpha(lnel.find(qn("a:solidFill")), alpha)
    return ln


def pill(s, x, y, w, h, label, color=MUTED, border=WHITE, alpha=0.14,
         size=7.5, bgalpha=0.04):
    shape(s, x, y, w, h, bg=WHITE, bg_alpha=bgalpha, border=border,
          border_alpha=alpha, radius=h / 2)
    text(s, x, y + (h - size * 1.33) / 2 - 1, w, size * 1.6, label, size=size,
         color=color, align="c", spc=0.4)


def footer(s, n, total=37):
    text(s, 64, 686, 620, 16, "UNIVERSITAS UBUDIYAH INDONESIA  ×  META",
         size=6.5, color=MUTED2, spc=1.1)
    text(s, 1050, 686, 166, 16, f"{n:02d} / {total}", size=6.5, color=MUTED2,
         align="r", spc=0.8)


def header(s, eyebrow, title_runs, tag=None, lede=None, y=44):
    text(s, 90, y + 2, 700, 14, eyebrow.upper(), size=7.5, color=MUTED,
         bold=True, spc=1.6)
    shape(s, 64, y + 5, 18, 2, kind=MSO_SHAPE.RECTANGLE, bg=RED)
    text(s, 90, y + 19, 1110, 40, title_runs, size=30, color=TEXT,
         bold=True, line_spacing=1.02, spc=-0.5)
    if tag:
        w = max(104, min(180, 10 + len(tag) * 5.4))
        pill(s, 1216 - w, y + 1, w, 25, tag.upper(), color=MUTED, alpha=0.18,
             size=7)
    if lede:
        text(s, 90, y + 80, 1080, 46, lede, size=10.5, color=MUTED,
             line_spacing=1.35)


def note(s, x, y, w, content, h=42, tone="amber"):
    if tone == "amber":
        shape(s, x, y, w, h, bg=AMBER, bg_alpha=0.075, border=AMBER,
              border_alpha=0.20, radius=9)
        ic_color = AMBER
    else:
        shape(s, x, y, w, h, bg=WHITE, bg_alpha=0.03, border=WHITE,
              border_alpha=0.10, radius=9)
        ic_color = REDB
    text(s, x + 14, y + 9, 20, 18, "→", size=10, color=ic_color, bold=True)
    text(s, x + 34, y + 8, w - 48, h - 12, content, size=8.8, color=MUTED,
         line_spacing=1.3)


# ---------------------------------------------------------------- deck
prs = Presentation()
prs.core_properties.author = "Direktorat Cyber Development Center & Media UBOnTV UUI"
prs.core_properties.title = "UUI x Meta — Strategi Meta Ads & 1 Juta Followers"
prs.core_properties.subject = "Rencana 18 bulan media sosial dan Meta Ads Universitas Ubudiyah Indonesia"
prs.slide_width = Inches(13.3333)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]
BG = {
    "content": A("bg-content.png"),
    "cover": A("bg-cover.png"),
    "divider": A("bg-divider.png"),
}


def new_slide(bg="content"):
    s = prs.slides.add_slide(BLANK)
    s.shapes.add_picture(BG[bg], 0, 0, width=prs.slide_width,
                         height=prs.slide_height)
    return s


TOTAL = 37

ASSET = lambda *p: os.path.join(ROOT, "assets", *p)


def title_lines(s, lines, x=90, y=64, w=1110, size=30, h=100):
    tf = textbox(s, x, y, w, h)
    for i, runs in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.line_spacing = 1.02
        p.space_before = Pt(0)
        p.space_after = Pt(0)
        _runs(p, runs, size, TEXT, True, False, FONT, -0.5)


def header2(s, eyebrow, lines, tag=None, lede=None, y=44, size=30, w=1110):
    text(s, 90, y + 2, 760, 14, eyebrow.upper(), size=7.5, color=MUTED,
         bold=True, spc=1.6)
    shape(s, 64, y + 5, 18, 2, kind=MSO_SHAPE.RECTANGLE, bg=RED)
    title_lines(s, lines, y=y + 19, w=w, size=size)
    if tag:
        tw = max(104, min(190, 12 + len(tag) * 5.2))
        pill(s, 1216 - tw, y + 1, tw, 25, tag.upper(), color=MUTED, alpha=0.18,
             size=7)
    if lede:
        text(s, 90, y + 84, 1080, 46, lede, size=10.5, color=MUTED,
             line_spacing=1.35)


# ============================================================ 01 COVER
s = new_slide("cover")
for (x, y, size, al) in [(880, -140, 760, 0.055), (975, -45, 570, 0.045),
                         (1070, 50, 380, 0.04)]:
    shape(s, x, y, size, size, kind=MSO_SHAPE.OVAL, border=WHITE,
          border_alpha=al, bw=1)

s.shapes.add_picture(ASSET("uui-logo.png"), IN(64), IN(52), height=IN(50))
shape(s, 215, 62, 1, 30, kind=MSO_SHAPE.RECTANGLE, bg=WHITE, bg_alpha=0.18)
s.shapes.add_picture(ASSET("meta-logo-white.png"), IN(234), IN(66), height=IN(22))

pill(s, 1004, 44, 212, 29, "DOKUMEN INTERNAL  ·  2026", color=MUTED, alpha=0.18,
     size=7.2)

shape(s, 64, 274, 42, 2, kind=MSO_SHAPE.RECTANGLE, bg=RED)
text(s, 118, 266, 700, 16, "RENCANA INDUK MEDIA SOSIAL & META ADS", size=7.5,
     color=REDB, bold=True, spc=1.8)

tf = textbox(s, 64, 294, 940, 190)
p = tf.paragraphs[0]
p.line_spacing = 0.95
_runs(p, [("Dari Kampus", {"color": RGBColor(0xE9, 0xEF, 0xFB)})], 60, TEXT,
      True, False, FONT, -1.6)
p2 = tf.add_paragraph()
p2.line_spacing = 0.95
_runs(p2, [("ke Feed.", {"color": RGBColor(0x66, 0x73, 0x8C)})], 60, TEXT, True, False,
      FONT, -1.6)

text(s, 64, 500, 760, 64,
     [("Peta jalan 18 bulan untuk membangun kehadiran Universitas Ubudiyah Indonesia di ekosistem Meta — dari pendaftaran akun iklan hingga ", {}),
      ("1 juta audiens", {"bold": True, "color": WHITE}),
      (" yang benar-benar terlibat.", {})],
     size=11.5, color=MUTED, line_spacing=1.45)

# bottom info
for i, (label, value) in enumerate([
        ("PLATFORM", "Facebook · Instagram · Threads · WhatsApp"),
        ("TARGET", "1.000.000 audiens"),
        ("HORIZON", "18 bulan · 6 kuartal")]):
    x = [64, 382, 556][i]
    text(s, x, 632, 300, 12, label, size=6.8, color=MUTED2, spc=1.4)
    text(s, x, 648, 360, 18, value, size=9, color=TEXT, bold=True)

text(s, 900, 626, 316, 14, "Disusun & dipresentasikan oleh", size=7.2,
     color=MUTED2, align="r")
text(s, 900, 640, 316, 14, "Direktorat Cyber Development Center", size=7.8,
     color=MUTED, bold=True, align="r")
text(s, 900, 654, 316, 14, "& Media UBOnTV UUI", size=7.8,
     color=MUTED, bold=True, align="r")
text(s, 900, 668, 316, 14, "November 2026", size=7.2, color=MUTED2, align="r")


# ============================================================ 02 AGENDA
s = new_slide()
header2(s, "Agenda", [
    [("Lima bab, satu arah:", {})],
    [("dari nol menjadi mesin distribusi.", {"color": REDB})],
], tag="Pembuka")

items = [
    ("01", "Fondasi & ekosistem Meta", "Empat platform, satu portofolio bisnis, dan arsitektur aset yang tidak berantakan."),
    ("02", "Mendaftar Meta Ads", "Sepuluh langkah: dari Business Portfolio dan verifikasi bisnis sampai kampanye pertama tayang."),
    ("03", "Konten & engagement", "Pilar konten, ritme posting, anatomi Reels, dan playbook komunitas."),
    ("04", "Menuju 1 juta followers", "Matematika target, alokasi per platform, kurva pertumbuhan, dan anggaran."),
    ("05", "Eksekusi & tata kelola", "Tim, KPI, ritme kerja, risiko, dan lima keputusan yang diminta hari ini."),
]
y = 228
for num, t, d in items:
    shape(s, 64, y, 584, 64, bg=WHITE, bg_alpha=0.03, border=WHITE,
          border_alpha=0.09, radius=10)
    shape(s, 80, y + 18, 28, 28, bg=RED, radius=7)
    text(s, 80, y + 24, 28, 16, num, size=8, color=WHITE, bold=True, align="c")
    text(s, 122, y + 10, 500, 18, t, size=10, color=TEXT, bold=True)
    text(s, 122, y + 29, 508, 28, d, size=7.8, color=MUTED, line_spacing=1.28)
    y += 71

card(s, 680, 216, 536, 382, radius=16)
text(s, 704, 236, 460, 14, "HASIL YANG DIHARAPKAN", size=7, color=MUTED2,
     bold=True, spc=1.4)
text(s, 704, 258, 488, 70,
     "Satu rencana yang bisa langsung dieksekusi Senin depan: akun terverifikasi, mesin konten berjalan, iklan terukur, dan jalur jelas menuju 1 juta audiens.",
     size=9.2, color=MUTED, line_spacing=1.45)
hline(s, 704, 340, 488, alpha=0.10)
for i, (n, l) in enumerate([("10", "langkah pendaftaran"), ("120", "konten / bulan"),
                            ("4", "platform terintegrasi"), ("18", "bulan menuju target")]):
    cx = 704 + (i % 2) * 248
    cy = 360 + (i // 2) * 92
    text(s, cx, cy, 180, 30, n, size=20, color=TEXT, bold=True)
    text(s, cx, cy + 28, 190, 16, l, size=7.6, color=MUTED)

note(s, 680, 612, 536,
     [("Semua angka adalah asumsi kerja. ", {"bold": True, "color": TEXT}),
      ("Ganti dengan data aktual setelah audit 14 hari pertama.", {})], h=42)
footer(s, 2)


# ============================================================ 03 EXEC SUMMARY
s = new_slide()
header2(s, "Ringkasan eksekutif", [
    [("Satu juta audiens bukan satu kampanye.", {})],
    [("Ini sebuah sistem.", {"color": REDB})],
], tag="Pembuka")

metrics = [("1.000.000", "AUDIENS TERAFILIASI\nDALAM 18 BULAN", REDB),
           ("4", "PLATFORM META\nSATU PORTOFOLIO", BLUEB),
           ("35jt", "MEDIA SPEND REKOMENDASI\nPER BULAN", TEXT)]
for i, (n, l, c) in enumerate(metrics):
    x = 64 + i * 384
    if i:
        shape(s, x - 18, 228, 1, 74, kind=MSO_SHAPE.RECTANGLE, bg=WHITE, bg_alpha=0.10)
    text(s, x, 220, 340, 54, n, size=36, color=c, bold=True, spc=-1.0)
    tf = textbox(s, x, 292, 320, 34)
    for j, part in enumerate(l.split("\n")):
        p = tf.paragraphs[0] if j == 0 else tf.add_paragraph()
        p.line_spacing = 1.1
        _runs(p, part, 6.9, MUTED2, True, False, FONT, 1.2)

cards = [
    ("70/30", "Organik dulu,\niklan mempercepat", "70% pertumbuhan dari konten dan komunitas, 30% dari media berbayar. Iklan tidak menyelamatkan konten yang lemah."),
    ("4×", "Satu ide,\nempat format", "Reels untuk jangkauan, carousel untuk kedalaman, Threads untuk percakapan, WhatsApp untuk retensi."),
    ("120", "Konten\nper bulan", "4 konten/hari lintas platform, 60% video pendek. Volume adalah syarat, bukan pilihan."),
    ("14", "Hari untuk\nmemutuskan", "Setiap kreatif diuji 14 hari. Yang gagal dimatikan, yang menang digandakan 20% per minggu."),
]
for i, (badge, t, d) in enumerate(cards):
    x = 64 + i * 292
    card(s, x, 336, 276, 212)
    shape(s, x + 18, 354, 50, 26, bg=WHITE, bg_alpha=0.07, border=WHITE,
          border_alpha=0.12, radius=7)
    text(s, x + 18, 361, 50, 14, badge, size=7.6, color=TEXT, bold=True, align="c")
    tf = textbox(s, x + 18, 396, 240, 44)
    for j, line in enumerate(t.split("\n")):
        p = tf.paragraphs[0] if j == 0 else tf.add_paragraph()
        p.line_spacing = 1.05
        _runs(p, line, 10.6, TEXT, True, False, FONT, -0.2)
    text(s, x + 18, 444, 240, 86, d, size=7.8, color=MUTED, line_spacing=1.4)

note(s, 64, 574, 1152,
     [("Sikap yang diminta: ", {"bold": True, "color": TEXT}),
      ("target 1 juta dalam 18 bulan adalah target agresif. Rencana ini mengejarnya lewat kurva bertahap, bukan lonjakan tunggal — dan menyiapkan skenario 24 bulan bila realitas pasar menuntut.", {})],
     h=48)
footer(s, 3)


# ============================================================ 04 DIVIDER 01
s = new_slide("divider")
text(s, 84, 214, 330, 250, "01", size=172, color=RGBColor(0x2A, 0x34, 0x4A),
     bold=True, spc=-4)
shape(s, 90, 508, 96, 3, kind=MSO_SHAPE.RECTANGLE, bg=RED)
text(s, 430, 242, 300, 14, "FONDASI", size=7.5, color=REDB, bold=True, spc=2.2)
title_lines(s, [
    [("Ekosistem Meta bukan", {})],
    [("satu aplikasi. Ia ", {}), ("satu sistem", {"color": REDB})],
    [("distribusi.", {})],
], x=430, y=268, w=770, size=42, h=220)
text(s, 430, 488, 620, 50,
     "Sebelum membicarakan iklan, kita rapikan rumahnya: aset, akses, data, dan siapa yang bertanggung jawab.",
     size=11.5, color=MUTED, line_spacing=1.4)
footer(s, 4)


# ============================================================ 05 PLATFORMS
s = new_slide()
header2(s, "Ekosistem", [
    [("Empat permukaan, ", {}), ("satu suara.", {"color": BLUEB})],
], tag="Bab 01 · Fondasi",
    lede="Setiap platform punya pekerjaan yang berbeda. Memperlakukan semuanya sama adalah cara tercepat untuk tidak unggul di mana pun.")

platforms = [
    ("fb", "Facebook Page", "Kredibilitas & jangkauan 25+.", "Pengumuman resmi, live, album kegiatan, dan grup alumni/orang tua.", "KPI: reach · share · klik", "blue"),
    ("ig", "Instagram", "Mesin discovery.", "Reels adalah etalase utama untuk audiens 16–24; carousel untuk kedalaman; story untuk keseharian.", "KPI: reach · save · follow", "red"),
    ("th", "Threads", "Percakapan harian.", "Tempat opini, behind the scene, dan polling. Murah diproduksi, cepat menyebar.", "KPI: reply · repost · mention", "plain"),
    ("wa", "WhatsApp Channel", "Distribusi langsung.", "Pengumuman PMB, reminder beasiswa, dan retensi audiens yang sudah percaya.", "KPI: view · CTR · unsubscribe", "plain"),
]
for i, (ic, t, strong, d, kpi, tone) in enumerate(platforms):
    x = 64 + i * 292
    card(s, x, 250, 276, 306, tone=tone)
    if ic == "fb":
        s.shapes.add_picture(ASSET("2021_Facebook_icon.png"), IN(x + 18), IN(268), width=IN(42), height=IN(42))
    elif ic == "ig":
        s.shapes.add_picture(ASSET("Instagram_logo_2016.png"), IN(x + 18), IN(268), width=IN(42), height=IN(42))
    elif ic == "wa":
        s.shapes.add_picture(ASSET("WhatsApp.png"), IN(x + 18), IN(268), width=IN(42), height=IN(42))
    else:
        shape(s, x + 18, 268, 42, 42, bg=WHITE, bg_alpha=0.08, border=WHITE,
              border_alpha=0.16, radius=12)
        text(s, x + 18, 277, 42, 24, "@", size=16, color=WHITE, bold=True, align="c")

    text(s, x + 18, 326, 240, 18, t, size=11, color=TEXT, bold=True)
    text(s, x + 18, 350, 242, 30, strong, size=8.6, color=WHITE, bold=True, line_spacing=1.25)
    text(s, x + 18, 378, 242, 100, d, size=8.2, color=MUTED, line_spacing=1.4)
    hline(s, x + 18, 490, 240, alpha=0.10)
    text(s, x + 18, 500, 240, 30, kpi, size=7.2, color=MUTED2)

note(s, 64, 590, 1152,
     [("Aturan main: ", {"bold": True, "color": TEXT}),
      ("satu identitas visual, satu nada bicara, empat format berbeda. Jangan copy-paste konten yang sama tanpa menyesuaikan konteks platform.", {})],
     h=44)
footer(s, 5)


def table(s, x, y, w, headers, rows, cols, row_h=44, header_h=26,
          hl=None, size=8.2, head_size=6.5):
    cx = x
    for i, h in enumerate(headers):
        text(s, cx, y, cols[i] - 10, header_h, h.upper(), size=head_size,
             color=MUTED2, bold=True, spc=1.1)
        cx += cols[i]
    hline(s, x, y + header_h, w, alpha=0.16)
    yy = y + header_h
    for r, row in enumerate(rows):
        if hl is not None and r == hl:
            shape(s, x - 10, yy - 2, w + 20, row_h + 4, bg=RED, bg_alpha=0.09,
                  radius=7)
        cx = x
        for i, val in enumerate(row):
            color = MUTED
            bold = False
            if isinstance(val, tuple):
                if len(val) == 3:
                    val, color, bold = val
                else:
                    val, color = val
            if i == 0:
                bold = True
                color = TEXT
            tb = text(s, cx, yy + (row_h - size * 1.5) / 2, cols[i] - 14, row_h,
                      val, size=size, color=color, bold=bold, line_spacing=1.25)
            cx += cols[i]
        if r < len(rows) - 1:
            hline(s, x, yy + row_h, w, alpha=0.06)
        yy += row_h
    return yy


# ============================================================ 06 ARCHITECTURE
s = new_slide()
header2(s, "Arsitektur", [
    [("Bangun rumahnya dulu, ", {}), ("baru pasang iklannya.", {"color": REDB})],
], tag="Bab 01 · Fondasi")

# diagram
shape(s, 64, 238, 760, 252, bg=WHITE, bg_alpha=0.02, border=WHITE,
      border_alpha=0.10, radius=14)
shape(s, 80, 252, 728, 52, bg=RED, bg_alpha=0.16, border=REDB,
      border_alpha=0.35, radius=9)
text(s, 94, 262, 86, 30, "BUSINESS\nPORTFOLIO", size=5.8, color=RGBColor(0xFF, 0x9B, 0xA5),
     bold=True, spc=0.8, line_spacing=1.15)
text(s, 186, 266, 380, 18, "Universitas Ubudiyah Indonesia", size=9.4,
     color=TEXT, bold=True)
text(s, 540, 269, 254, 14, "business.facebook.com · 1 pemilik", size=6.4,
     color=MUTED2, align="r")

arch = [
    ("ASET", ["Facebook Page", "Instagram Profesional", "Threads", "WhatsApp Business"]),
    ("IKLAN", ["Ad Account — Brand", "Ad Account — PMB", "Ad Account — Fakultas", "+ akun cadangan"]),
    ("DATA", ["Dataset / Pixel", "Conversions API", "Katalog konten", "+ UTM & GA4"]),
    ("AKSES", ["Admin (2 orang)", "Advertiser", "Analyst", "Finance (read-only)"]),
]
for i, (label, lines) in enumerate(arch):
    x = 80 + i * 182
    shape(s, x, 318, 174, 122, bg=WHITE, bg_alpha=0.025, border=WHITE,
          border_alpha=0.09, radius=8)
    text(s, x + 10, 326, 150, 12, label, size=5.8, color=REDB, bold=True, spc=1.0)
    tf = textbox(s, x + 10, 342, 155, 108)
    for j, ln in enumerate(lines):
        p = tf.paragraphs[0] if j == 0 else tf.add_paragraph()
        p.line_spacing = 1.55
        _runs(p, ln, 7.2, MUTED2 if ln.startswith("+") else MUTED, False,
              ln.startswith("+"), FONT, 0)
shape(s, 80, 448, 728, 30, bg=WHITE, bg_alpha=0.02, border=None, radius=6)
text(s, 80, 455, 728, 16, "SATU PORTOFOLIO  ·  BANYAK AKUN PER TUJUAN  ·  AKSES BERBASIS PERAN",
     size=6.2, color=MUTED2, align="c", spc=1.0)

rules = [
    ("1", "Jangan lahir di akun pribadi", "Aset yang dibuat di akun staf akan hilang saat staf itu pindah. Semua harus di portofolio institusi."),
    ("2", "Pisahkan akun iklan per tujuan", "Brand, PMB, dan fakultas punya audiens dan anggaran berbeda. Laporan jadi bersih, evaluasi jadi adil."),
    ("3", "Bagikan akses, bukan password", "Undang lewat email dengan peran terbatas. Password bersama adalah celah keamanan terbesar."),
]
y = 238
for num, t, d in rules:
    shape(s, 848, y, 368, 88, bg=WHITE, bg_alpha=0.03, border=WHITE,
          border_alpha=0.09, radius=10)
    shape(s, 862, y + 16, 24, 24, bg=RED, radius=6)
    text(s, 862, y + 21, 24, 12, num, size=7.2, color=WHITE, bold=True, align="c")
    text(s, 898, y + 13, 300, 16, t, size=8.8, color=TEXT, bold=True)
    text(s, 898, y + 31, 304, 48, d, size=7.2, color=MUTED, line_spacing=1.35)
    y += 97
footer(s, 6)


# ============================================================ 07 DIVIDER 02
s = new_slide("divider")
text(s, 84, 214, 330, 250, "02", size=172, color=RGBColor(0x2A, 0x34, 0x4A),
     bold=True, spc=-4)
shape(s, 90, 508, 96, 3, kind=MSO_SHAPE.RECTANGLE, bg=RED)
text(s, 430, 242, 300, 14, "PENDAFTARAN", size=7.5, color=REDB, bold=True, spc=2.2)
title_lines(s, [
    [("Sepuluh langkah.", {})],
    [("Dari nol sampai ", {}), ("tayang.", {"color": REDB})],
], x=430, y=274, w=760, size=44, h=140)
text(s, 430, 436, 620, 60,
     "Urutannya penting. Melompat langsung ke “Boost Post” tanpa verifikasi dan pelacakan adalah cara membakar anggaran tanpa data.",
     size=11.5, color=MUTED, line_spacing=1.4)
footer(s, 7)


# ============================================================ 08 TEN STEPS
s = new_slide()
header2(s, "Peta pendaftaran", [
    [("Sepuluh langkah menuju kampanye pertama.", {})],
], tag="Bab 02 · Meta Ads")

steps = [
    ("01", "Akun Facebook kerja", "Akun personal resmi + 2FA. Bukan akun bot, bukan akun pinjaman."),
    ("02", "Business Portfolio", "Buat di business.facebook.com dengan email domain kampus."),
    ("03", "Hubungkan aset", "Page UUI, Instagram profesional, Threads, WhatsApp Business."),
    ("04", "Verifikasi domain", "uui.ac.id lewat DNS TXT atau meta-tag di halaman utama."),
    ("05", "Verifikasi bisnis", "Unggah dokumen legal. Proses 3–7 hari kerja."),
    ("06", "Buat Ad Account", "Mata uang IDR, zona waktu Asia/Jakarta, per tujuan."),
    ("07", "Metode pembayaran", "Kartu korporat atau transfer bank. Siapkan NPWP institusi."),
    ("08", "Pixel + CAPI", "Pasang di situs, kirim event server-side, uji lewat Test Events."),
    ("09", "Tim & izin", "Admin, advertiser, analyst, finance. Prinsip akses minimum."),
    ("10", "Kampanye pertama", "Tujuan → audiens → placement → anggaran → kreatif → tayang."),
]
for i, (n, t, d) in enumerate(steps):
    x = 64 + (i % 5) * 234
    y = 222 + (i // 5) * 132
    shape(s, x, y, 224, 120, bg=WHITE, bg_alpha=0.035, border=WHITE,
          border_alpha=0.09, radius=10)
    text(s, x + 14, y + 12, 40, 20, n, size=13, color=RGBColor(0x55, 0x5F, 0x75), bold=True)
    hline(s, x + 52, y + 22, 155, alpha=0.16)
    text(s, x + 14, y + 40, 194, 18, t, size=8.5, color=TEXT, bold=True)
    text(s, x + 14, y + 60, 194, 52, d, size=7.0, color=MUTED2, line_spacing=1.4)

info = [
    ("ESTIMASI WAKTU", [("7–14 hari", {"bold": True, "color": TEXT}),
                        (" — 1 hari setup, 3–7 hari verifikasi, 2 hari pelacakan & uji.", {})]),
    ("YANG PALING SERING GAGAL", [("Nama bisnis tidak konsisten dengan dokumen, domain belum terverifikasi, dan tidak ada metode pembayaran aktif.", {})]),
    ("YANG WAJIB ADA DI HARI PERTAMA", [("Akses admin ganda, 2FA, email domain kampus, dan dokumen legal dalam bentuk PDF.", {})]),
]
for i, (h, body) in enumerate(info):
    x = 64 + i * 390
    shape(s, x, 500, 372, 92, bg=WHITE, bg_alpha=0.022, border=WHITE,
          border_alpha=0.08, radius=10)
    text(s, x + 16, 512, 330, 12, h, size=6.2, color=MUTED2, bold=True, spc=1.1)
    text(s, x + 16, 530, 338, 52, body, size=8.2, color=MUTED, line_spacing=1.35)
footer(s, 8)


# ============================================================ 09 STEPS 1-5
s = new_slide()
header2(s, "Langkah 1–5", [
    [("Menyiapkan identitas dan kepemilikan aset.", {})],
], tag="Bab 02 · Meta Ads")

left = [
    ("1", "Akun Facebook kerja yang aman", "Gunakan akun resmi atas nama jabatan (mis. Digital Lead UUI), aktifkan 2FA, dan simpan recovery code di tempat terkunci. Hindari akun yang terikat nomor pribadi staf."),
    ("2", "Buat Business Portfolio", "Buka business.facebook.com → Create account. Nama: “Universitas Ubudiyah Indonesia”. Email kerja berdomain uui.ac.id. Ini menjadi rumah semua aset, bukan akun iklan itu sendiri."),
    ("3", "Hubungkan Page, Instagram, Threads, WhatsApp", "Business Settings → Accounts. Instagram harus akun profesional (Creator/Business) dan ditautkan ke Page. Threads otomatis mengikuti akun IG. WhatsApp Business dihubungkan lewat nomor institusi."),
]
right = [
    ("4", "Verifikasi domain uui.ac.id", "Brand Safety → Domains → Add. Pilih DNS TXT (paling stabil) atau unggah meta-tag ke bagian head halaman utama. Verifikasi domain membuka pelacakan lintas perangkat dan pengukuran konversi yang akurat."),
    ("5", "Verifikasi bisnis (Business Verification)", "Security Center → Start verification. Siapkan NIB, akta pendirian, SK, dan bukti alamat. Nama, alamat, dan nomor telepon harus identik di semua dokumen dan di halaman Facebook."),
]
for col, items, x in [(0, left, 64), (1, right, 660)]:
    y = 214
    for num, t, d in items:
        h = 96 if num != "3" else 108
        shape(s, x, y, 556, h, bg=WHITE, bg_alpha=0.03, border=WHITE,
              border_alpha=0.09, radius=10)
        shape(s, x + 14, y + 14, 24, 24, bg=RED, radius=6)
        text(s, x + 14, y + 19, 24, 12, num, size=7.2, color=WHITE, bold=True, align="c")
        text(s, x + 50, y + 11, 490, 16, t, size=8.8, color=TEXT, bold=True)
        text(s, x + 50, y + 28, 490, h - 36, d, size=7.6, color=MUTED, line_spacing=1.38)
        y += h + 10

note(s, 660, 434, 556,
     [("Jangan lewati verifikasi. ", {"bold": True, "color": TEXT}),
      ("Tanpa verifikasi, akun rentan dibatasi, tidak bisa menjalankan format iklan tertentu, dan sulit mengajukan banding saat terjadi peninjauan.", {})],
     h=76)
footer(s, 9)


# ============================================================ 10 STEPS 6-10
s = new_slide()
header2(s, "Langkah 6–10", [
    [("Menyiapkan uang, data, dan orang.", {})],
], tag="Bab 02 · Meta Ads")

left = [
    ("6", "Buat Ad Account per tujuan", "Mata uang IDR, zona waktu Asia/Jakarta. Pisahkan: Brand & Engagement, PMB, dan Fakultas. Timezone tidak bisa diubah setelah akun dipakai — pastikan sejak awal."),
    ("7", "Atur metode pembayaran", "Kartu kredit/korporat adalah pilihan paling stabil. Alternatif: transfer bank atau invoice untuk institusi. Siapkan NPWP dan alamat penagihan yang sesuai."),
    ("8", "Pasang Pixel + Conversions API", "Events Manager → Dataset. Pasang base code lewat Google Tag Manager, lalu kirim event server-side (CAPI) untuk mengurangi kehilangan data akibat pemblokir iklan. Uji dengan Test Events sebelum menyalakan iklan."),
]
for num, t, d in left:
    y = 214 + (int(num) - 6) * 118
    shape(s, 64, y, 556, 106, bg=WHITE, bg_alpha=0.03, border=WHITE,
          border_alpha=0.09, radius=10)
    shape(s, 78, y + 14, 24, 24, bg=RED, radius=6)
    text(s, 78, y + 19, 24, 12, num, size=7.2, color=WHITE, bold=True, align="c")
    text(s, 114, y + 11, 490, 16, t, size=8.8, color=TEXT, bold=True)
    text(s, 114, y + 28, 490, 72, d, size=7.6, color=MUTED, line_spacing=1.38)

right = [
    ("9", "Atur tim dan izin", "Business Settings → People. Admin maksimal dua orang, advertiser untuk tim performa, analyst read-only untuk pelaporan, finance hanya untuk penagihan. Tinjau akses setiap kuartal."),
    ("10", "Jalankan kampanye pertama", "Ads Manager → Create. Pilih tujuan yang benar (bukan sekadar “Boost”), tentukan audiens, placement, anggaran, dan 3 varian kreatif. Tinjauan Meta umumnya < 24 jam."),
]
for i, (num, t, d) in enumerate(right):
    y = 214 + i * 118
    shape(s, 660, y, 556, 106, bg=WHITE, bg_alpha=0.03, border=WHITE,
          border_alpha=0.09, radius=10)
    shape(s, 674, y + 14, 24, 24, bg=RED, radius=6)
    text(s, 674, y + 19, 24, 12, num, size=7.2, color=WHITE, bold=True, align="c")
    text(s, 710, y + 11, 490, 16, t, size=8.8, color=TEXT, bold=True)
    text(s, 710, y + 28, 490, 72, d, size=7.6, color=MUTED, line_spacing=1.38)

shape(s, 660, 452, 556, 116, bg=RED, bg_alpha=0.13, border=REDB,
      border_alpha=0.30, radius=10)
text(s, 680, 464, 500, 14, "KESALAHAN PEMULA", size=6.4, color=RGBColor(0xFF, 0xB3, 0xBA),
     bold=True, spc=1.4)
text(s, 680, 484, 516, 74,
     "Menekan “Boost Post” untuk semua konten. Boost mengoptimalkan interaksi, bukan pertumbuhan follower atau pendaftaran. Gunakan Ads Manager untuk tujuan yang terukur.",
     size=8.4, color=RGBColor(0xF3, 0xD7, 0xDA), line_spacing=1.4)
footer(s, 10)


# ============================================================ 11 VERIFICATION
s = new_slide()
header2(s, "Verifikasi", [
    [("Dokumen yang harus siap ", {}), ("sebelum mulai.", {"color": REDB})],
], tag="Bab 02 · Meta Ads")

text(s, 64, 222, 500, 14, "CHECKLIST DOKUMEN", size=6.4, color=MUTED2, bold=True, spc=1.4)
checks = [
    ("NIB / izin operasional", " — nomor dan nama badan hukum harus terbaca jelas."),
    ("Akta pendirian & SK", " — menunjukkan legalitas dan pengurus yang berwenang."),
    ("Bukti alamat operasional", " — tagihan listrik/internet 3 bulan terakhir."),
    ("Email domain kampus", " — bukan Gmail pribadi; mempercepat verifikasi."),
    ("Nomor telepon bisnis", " — aktif dan dijawab saat Meta melakukan konfirmasi."),
    ("Website dengan nama & alamat", " — halaman kontak dan tentang kami harus lengkap."),
    ("Surat kuasa", " — bila pendaftar bukan pemegang jabatan resmi."),
]
y = 248
for strong, rest in checks:
    shape(s, 64, y + 2, 15, 15, bg=GREEN, bg_alpha=0.14, border=GREEN,
          border_alpha=0.30, radius=4)
    text(s, 67, y + 4, 12, 10, "✓", size=6.5, color=GREEN, bold=True)
    text(s, 88, y, 516, 22,
         [(strong, {"bold": True, "color": TEXT}), (rest, {})],
         size=8.1, color=MUTED, line_spacing=1.25)
    y += 34

vtl = [
    ("HARI 1", "Kumpulkan & samakan data", "Nama bisnis, alamat, telepon, dan domain harus konsisten di dokumen, Page, dan situs."),
    ("HARI 2", "Ajukan verifikasi", "Security Center → Business Verification. Unggah PDF, bukan foto miring."),
    ("HARI 3–7", "Respons pertanyaan Meta", "Balas cepat. Satu dokumen kurang bisa mengulang proses dari awal."),
    ("SETELAHNYA", "Aktifkan Meta Verified", "Langganan centang biru untuk Page & IG (bila tersedia): kepercayaan publik, dukungan akun, dan perlindungan penyamaran."),
]
y = 230
for when, t, d in vtl:
    shape(s, 660, y + 6, 7, 7, kind=MSO_SHAPE.OVAL, bg=RED)
    if when != "SETELAHNYA":
        shape(s, 663, y + 16, 1, 62, kind=MSO_SHAPE.RECTANGLE, bg=WHITE, bg_alpha=0.16)
    text(s, 682, y, 420, 12, when, size=6.2, color=MUTED2, bold=True, spc=1.2)
    text(s, 682, y + 15, 510, 16, t, size=8.8, color=TEXT, bold=True)
    text(s, 682, y + 32, 510, 48, d, size=7.3, color=MUTED, line_spacing=1.35)
    y += 78

note(s, 64, 610, 1152,
     "Verifikasi domain dan verifikasi bisnis adalah dua hal berbeda. Keduanya dibutuhkan: yang pertama untuk pengukuran, yang kedua untuk kepercayaan dan stabilitas akun iklan.",
     h=42)
footer(s, 11)


# ============================================================ 12 ROLES
s = new_slide()
header2(s, "Tata kelola akun", [
    [("Akses minimum, ", {}), ("akuntabilitas maksimum.", {"color": BLUEB})],
], tag="Bab 02 · Meta Ads")

shape(s, 64, 218, 750, 292, bg=WHITE, bg_alpha=0.02, border=WHITE,
      border_alpha=0.09, radius=14)
table(s, 88, 232, 702,
      ["Peran", "Siapa", "Izin"],
      [["Portfolio Admin", "Kepala Humas + Digital Lead", "Penuh · 2 orang"],
       ["Advertiser", "Performance marketer", "Buat & edit kampanye"],
       ["Analyst", "Staf data / reporting", "Read-only"],
       ["Finance", "Bagian keuangan", "Billing & invoice"],
       ["Partner", "Agensi / vendor kreatif", "Aset terbatas"]],
      [190, 300, 212], row_h=44, header_h=28, size=8.4)

right = [
    ("WAJIB", "2FA untuk semua admin · recovery code tercetak · audit akses tiap kuartal · email domain institusi.", "plain"),
    ("DILARANG", "Password dibagikan via chat · akun iklan di profil pribadi · satu orang memegang semua kendali.", "plain"),
    ("RENCANA KONTINJENSI", "Minimal dua admin aktif. Bila akun dibatasi, admin kedua bisa mengajukan banding tanpa menunggu satu orang.", "red"),
]
y = 218
for h, body, tone in right:
    shape(s, 840, y, 376, 98 if tone == "plain" else 104,
          bg=RED if tone == "red" else WHITE,
          bg_alpha=0.10 if tone == "red" else 0.03,
          border=REDB if tone == "red" else WHITE,
          border_alpha=0.28 if tone == "red" else 0.09, radius=10)
    text(s, 858, y + 14, 330, 12, h, size=6.2, color=RGBColor(0xFF, 0xB3, 0xBA) if tone == "red" else MUTED2,
         bold=True, spc=1.2)
    text(s, 858, y + 32, 340, 58, body, size=7.6, color=MUTED, line_spacing=1.38)
    y += (98 if tone == "plain" else 104) + 10
footer(s, 12)


# ============================================================ 13 PIXEL
s = new_slide()
header2(s, "Pengukuran", [
    [("Kalau tidak diukur, ", {}), ("itu bukan iklan.", {"color": REDB})],
], tag="Bab 02 · Meta Ads")

def flow_node(x, y, w, h, kicker, title, accent=False):
    shape(s, x, y, w, h, bg=RED if accent else WHITE,
          bg_alpha=0.13 if accent else 0.035,
          border=REDB if accent else WHITE,
          border_alpha=0.32 if accent else 0.10, radius=8)
    text(s, x + 12, y + 7, w - 24, 10, kicker.upper(), size=5.6, color=RGBColor(0xFF, 0x9B, 0xA5) if accent else MUTED2, bold=True, spc=1.0)
    text(s, x + 12, y + 19, w - 24, 14, title, size=7.4, color=TEXT, bold=True)

flow_node(64, 218, 540, 42, "Situs & aplikasi", "uui.ac.id · landing PMB")
text(s, 64, 260, 540, 16, "↓", size=10, color=MUTED2, align="c")
flow_node(64, 278, 264, 42, "Browser", "Meta Pixel")
flow_node(340, 278, 264, 42, "Server", "Conversions API")
text(s, 64, 320, 540, 16, "↓", size=10, color=MUTED2, align="c")
flow_node(64, 338, 540, 42, "Events Manager", "Dataset UUI · deduplikasi event_id")
text(s, 64, 380, 540, 16, "↓", size=10, color=MUTED2, align="c")
flow_node(64, 398, 540, 44, "Optimasi", "Ads Manager belajar dari konversi nyata", True)

note(s, 64, 464, 540,
     "Pasang consent banner sesuai UU PDP. Jangan kirim data pribadi sensitif (NIK, nomor HP) sebagai parameter event.",
     h=62)

text(s, 640, 218, 576, 14, "EVENT YANG WAJIB DILACAK", size=6.4, color=MUTED2, bold=True, spc=1.4)
rows = [["PageView", "Semua halaman", "Dasar"],
        ["ViewContent", "Halaman program studi", "Minat"],
        ["Lead", "Formulir PMB terkirim", "Tinggi"],
        ["Contact", "Klik WhatsApp / telepon", "Tinggi"],
        ["CompleteRegistration", "Akun pendaftaran dibuat", "Konversi"],
        ["Subscribe", "Newsletter / WA Channel", "Retensi"]]
table(s, 640, 240, 576, ["Event", "Kapan", "Nilai"], rows,
      [200, 250, 126], row_h=42, header_h=26, size=8.2)
text(s, 640, 540, 560, 28, "Uji dengan Test Events sebelum kampanye. Bandingkan jumlah event browser vs server setiap minggu.",
     size=6.8, color=MUTED2, line_spacing=1.35)
footer(s, 13)


# ============================================================ 14 CAMPAIGN
s = new_slide()
header2(s, "Kampanye", [
    [("Pilih tujuan yang tepat, ", {}), ("bukan yang termudah.", {"color": BLUEB})],
], tag="Bab 02 · Meta Ads")

rows = [
    ["Awareness", "Kampanye kampus, Reels, momen wisuda", "Reach / video views", "CPM Rp 15–30 rb"],
    ["Engagement", "Membangun social proof & percakapan", "Post engagement", "Rp 1.500–3.500 / interaksi"],
    ["Traffic", "Mengarahkan ke profil & landing PMB", "Link clicks / profile visits", "Rp 800–2.500 / klik"],
    ["Leads", "Formulir PMB, unduh brosur, webinar", "Lead", "Rp 15–45 rb / lead"],
    ["Conversions", "Pendaftaran & pembayaran", "Event konversi", "Biaya per pendaftar"],
]
shape(s, 64, 218, 1152, 268, bg=WHITE, bg_alpha=0.02, border=WHITE,
      border_alpha=0.09, radius=14)
table(s, 88, 232, 1104,
      ["Tujuan", "Dipakai untuk", "Optimasi", "Benchmark"],
      rows, [190, 390, 280, 244], row_h=44, header_h=28, hl=2, size=8.4)

cards = [
    ("Struktur minimal", "1 kampanye · 2–3 ad set (audiens) · 3 kreatif per ad set. Jangan 20 ad set dengan 1 kreatif."),
    ("Aturan belanja", "70% ke yang terbukti · 20% untuk scaling · 10% eksperimen format baru."),
    ("Yang dipantau", "Frekuensi < 3, CPM stabil, CTR > 1%, dan biaya per hasil turun dalam 14 hari."),
]
for i, (h, d) in enumerate(cards):
    x = 64 + i * 390
    shape(s, x, 512, 372, 92, bg=WHITE, bg_alpha=0.022, border=WHITE,
          border_alpha=0.08, radius=10)
    text(s, x + 16, 524, 330, 12, h.upper(), size=6.2, color=MUTED2, bold=True, spc=1.1)
    text(s, x + 16, 542, 340, 52, d, size=8.0, color=MUTED, line_spacing=1.36)
footer(s, 14)


# ============================================================ 15 GO-LIVE
s = new_slide()
header2(s, "Sebelum menekan publish", [
    [("Dua belas centang sebelum uang keluar.", {})],
], tag="Bab 02 · Meta Ads")

cols = [
    ["Business Portfolio aktif & terverifikasi", "Domain uui.ac.id terverifikasi", "Pixel menyala di semua halaman penting", "CAPI mengirim event & terdeduplikasi"],
    ["Metode pembayaran aktif & limit cukup", "Dua admin dengan 2FA", "UTM konsisten di semua link", "Landing page cepat & ramah mobile"],
    ["Tiga varian kreatif per ad set", "Caption & subtitle sudah diperiksa", "Consent banner & kebijakan privasi", "Dashboard pelaporan siap"],
]
for c, items in enumerate(cols):
    x = 64 + c * 390
    for r, it in enumerate(items):
        y = 230 + r * 40
        shape(s, x, y + 3, 16, 16, bg=GREEN, bg_alpha=0.14, border=GREEN,
              border_alpha=0.30, radius=4)
        text(s, x + 3, y + 5, 12, 10, "✓", size=6.8, color=GREEN, bold=True)
        text(s, x + 26, y, 340, 24, it, size=8.4, color=MUTED, line_spacing=1.25)

cards = [
    ("Batas aman", "Mulai Rp 150–300 rb/hari per kampanye. Naikkan maksimal 20% per 3 hari agar pembelajaran tidak ter-reset.", "red"),
    ("Frekuensi", "Ganti kreatif saat frekuensi > 3 atau CTR turun 30% dalam 7 hari. Kelelahan kreatif adalah pemborosan paling umum.", "plain"),
    ("Eksperimen", "Satu variabel per uji: hook, audiens, atau format. Lebih dari satu variabel membuat hasil tidak bisa dibaca.", "plain"),
]
for i, (h, d, tone) in enumerate(cards):
    x = 64 + i * 390
    shape(s, x, 418, 372, 112, bg=RED if tone == "red" else WHITE,
          bg_alpha=0.11 if tone == "red" else 0.025,
          border=REDB if tone == "red" else WHITE,
          border_alpha=0.28 if tone == "red" else 0.08, radius=10)
    text(s, x + 16, 432, 330, 12, h.upper(), size=6.2,
         color=RGBColor(0xFF, 0xB3, 0xBA) if tone == "red" else MUTED2, bold=True, spc=1.1)
    text(s, x + 16, 450, 340, 68, d, size=8.0, color=MUTED, line_spacing=1.36)
footer(s, 15)

# ============================================================ 16 DIVIDER 03
s = new_slide("divider")
text(s, 84, 214, 330, 250, "03", size=172, color=RGBColor(0x2A, 0x34, 0x4A),
     bold=True, spc=-4)
shape(s, 90, 508, 96, 3, kind=MSO_SHAPE.RECTANGLE, bg=RED)
text(s, 430, 242, 300, 14, "KONTEN & ENGAGEMENT", size=7.5, color=REDB, bold=True, spc=2.2)
title_lines(s, [
    [("Algoritma tidak jatuh cinta", {})],
    [("pada kampus. Manusia jatuh", {})],
    [("cinta pada ", {}), ("manusia.", {"color": REDB})],
], x=430, y=272, w=790, size=38, h=200)
text(s, 430, 482, 620, 60,
     "Konten institusi yang kaku akan dilewati. Konten yang terasa seperti kehidupan nyata mahasiswa akan dibagikan.",
     size=11.5, color=MUTED, line_spacing=1.4)
footer(s, 16)


# ============================================================ 17 PILLARS
s = new_slide()
header2(s, "Pilar konten", [
    [("Lima pilar, satu kepribadian: ", {}), ("UUI yang hidup.", {"color": REDB})],
], tag="Bab 03 · Konten")

pillars = [
    ("30%", "Kehidupan mahasiswa", "Day in the life, UKM, kost, organisasi, drama tugas. Wajah manusia, bukan gedung.", "Reels · story · Threads", "red"),
    ("20%", "Bukti & prestasi", "Wisuda, prestasi, alumni sukses, akreditasi, kerja sama industri.", "Carousel · FB · artikel", "blue"),
    ("20%", "Edukasi & tips", "Cara lolos seleksi, beasiswa, karier, AI untuk mahasiswa, literasi digital.", "Carousel · Threads · live", "amber"),
    ("15%", "Aceh & lokal", "Budaya, kuliner, wisata, komunitas. Akar lokal adalah pembeda dari kampus lain.", "Reels · kolaborasi", "green"),
    ("15%", "Humor & tren", "Meme kampus, POV dosen, audio viral. Cepat diproduksi, tinggi jangkauan.", "Reels · Threads", "plain"),
]
for i, (pct, t, d, fmt, tone) in enumerate(pillars):
    x = 64 + i * 232
    shape(s, x, 232, 220, 196, bg=WHITE, bg_alpha=0.03, border=WHITE,
          border_alpha=0.09, radius=12)
    pcol = {"red": REDB, "blue": BLUEB, "amber": AMBER, "green": GREEN, "plain": MUTED}[tone]
    shape(s, x + 16, 248, 44, 20, bg=pcol, bg_alpha=0.13, border=pcol,
          border_alpha=0.28, radius=10)
    text(s, x + 16, 252, 44, 12, pct, size=6.8, color=pcol, bold=True, align="c")
    text(s, x + 16, 280, 188, 32, t, size=9.6, color=TEXT, bold=True, line_spacing=1.1)
    text(s, x + 16, 316, 188, 78, d, size=7.3, color=MUTED, line_spacing=1.4)
    text(s, x + 16, 400, 188, 20, fmt, size=6.3, color=MUTED2)

shape(s, 64, 446, 564, 40, bg=AMBER, bg_alpha=0.07, border=AMBER,
      border_alpha=0.18, radius=8)
text(s, 80, 457, 530, 18,
     [("70% hiburan & nilai", {"bold": True, "color": TEXT}),
      (" — alasan orang mau menonton dan membagikan.", {})],
     size=8.2, color=MUTED)
shape(s, 640, 446, 576, 40, bg=RED, bg_alpha=0.09, border=REDB,
      border_alpha=0.22, radius=8)
text(s, 656, 457, 544, 18,
     [("30% bukti & promosi", {"bold": True, "color": TEXT}),
      (" — alasan orang mau mendaftar. Jangan dibalik.", {})],
     size=8.2, color=MUTED)
footer(s, 17)


# ============================================================ 18 CADENCE
s = new_slide()
header2(s, "Ritme", [
    [("Konsisten setiap hari, ", {}), ("bukan heroik sekali sebulan.", {"color": BLUEB})],
], tag="Bab 03 · Konten")

rows = [
    ["Instagram", "Reels 9:16, carousel, story", "1 Reels/hari · 3 story/hari · 3 carousel/mgg", "11.00 · 19.00", "Reach · save · follow"],
    ["Facebook", "Reels, link post, live", "2 post/hari · 1 live/mgg", "12.00 · 20.00", "Reach · share · klik"],
    ["Threads", "Teks, foto, repost", "3–5 post/hari", "08.00–22.00", "Reply · repost"],
    ["WhatsApp", "Pengumuman, voice note", "2–3 update/hari", "07.00 · 16.00", "View · CTR"],
]
shape(s, 64, 220, 1152, 244, bg=WHITE, bg_alpha=0.02, border=WHITE,
      border_alpha=0.09, radius=14)
table(s, 88, 234, 1104,
      ["Platform", "Format utama", "Frekuensi", "Jam tayang", "KPI utama"],
      rows, [160, 270, 350, 170, 154], row_h=46, header_h=28, size=8.0)

cards = [
    ("Produksi batch", "Satu hari syuting = 12–15 klip. Editing dua hari. Ini menjaga volume tanpa lembur harian."),
    ("Bank konten", "Simpan minimal 14 hari konten di depan. Buffer adalah pertahanan terhadap minggu sibuk."),
    ("Repurpose", "1 Reels → 1 carousel → 3 Threads → 1 story → 1 potongan untuk WhatsApp Channel."),
]
for i, (h, d) in enumerate(cards):
    x = 64 + i * 390
    shape(s, x, 492, 372, 92, bg=WHITE, bg_alpha=0.022, border=WHITE,
          border_alpha=0.08, radius=10)
    text(s, x + 16, 504, 330, 12, h.upper(), size=6.2, color=MUTED2, bold=True, spc=1.1)
    text(s, x + 16, 522, 340, 56, d, size=8.0, color=MUTED, line_spacing=1.36)
footer(s, 18)


# ============================================================ 19 REEL ANATOMY
s = new_slide()
header2(s, "Format unggulan", [
    [("Tiga detik pertama ", {}), ("menentukan segalanya.", {"color": REDB})],
], tag="Bab 03 · Konten")

# phone
shape(s, 168, 148, 214, 418, bg=RGBColor(0x0B, 0x0F, 0x19), border=WHITE,
      border_alpha=0.18, radius=28)
screen = shape(s, 177, 157, 196, 400, bg=RGBColor(0x17, 0x22, 0x36), border=None,
               radius=22)
screen.fill.gradient()
stops = screen.fill.gradient_stops
stops[0].color.rgb = RGBColor(0x3A, 0x12, 0x20)
stops[0].position = 0.0
stops[1].color.rgb = RGBColor(0x0D, 0x12, 0x20)
stops[1].position = 1.0
screen.fill.gradient_angle = 120
shape(s, 247, 157, 58, 15, bg=RGBColor(0x05, 0x07, 0x0C), border=None, radius=8)
text(s, 189, 176, 80, 10, "9:41", size=5.2, color=RGBColor(0xCC, 0xD6, 0xE8))
text(s, 285, 176, 80, 10, "UUI · Reels", size=5.2, color=RGBColor(0xCC, 0xD6, 0xE8), align="r")

# play affordance
shape(s, 250, 300, 48, 48, kind=MSO_SHAPE.OVAL, bg=WHITE, bg_alpha=0.10,
      border=WHITE, border_alpha=0.22)
shape(s, 268, 314, 0, 0, kind=MSO_SHAPE.ISOSCELES_TRIANGLE, bg=WHITE, bg_alpha=0.85)
play = s.shapes[-1]
play.width = IN(14); play.height = IN(18); play.left = IN(269); play.top = IN(314)
play.rotation = 90

text(s, 191, 444, 138, 60,
     [("“Aku pikir kuliah di Aceh itu sepi…”", {"bold": True, "color": WHITE})],
     size=6.6, color=WHITE, line_spacing=1.3)
text(s, 191, 476, 138, 40,
     "POV: minggu pertama di UUI. Ternyata…", size=6.0,
     color=RGBColor(0xD8, 0xDF, 0xEE), line_spacing=1.3)
for i in range(3):
    shape(s, 340, 420 + i * 30, 20, 20, bg=WHITE, bg_alpha=0.10, border=WHITE,
          border_alpha=0.18, radius=5)
shape(s, 240, 550, 70, 3, bg=WHITE, bg_alpha=0.30, radius=2)

beats = [
    ("0–1s", "Hook visual + teks", "Satu klaim, satu wajah, satu gerakan. Tanpa intro logo."),
    ("1–3s", "Janji yang jelas", "“Ini yang tidak dikatakan brosur kampus.” Beri alasan untuk bertahan."),
    ("3–18s", "Isi yang padat", "3–5 beat cepat. Subtitle wajib — 85% menonton tanpa suara."),
    ("18–25s", "Payoff & loop", "Akhir yang bisa diputar ulang tanpa terasa mengulang."),
    ("Akhir", "CTA spesifik", "“Simpan untuk besok” · “Komentar UUI” · “Ikuti untuk part 2”."),
]
y = 178
for tag, t, d in beats:
    shape(s, 430, y, 786, 58, bg=WHITE, bg_alpha=0.03, border=WHITE,
          border_alpha=0.09, radius=9)
    shape(s, 444, y + 15, 40, 28, bg=RED, radius=6)
    text(s, 444, y + 22, 40, 14, tag, size=6.2, color=WHITE, bold=True, align="c")
    text(s, 496, y + 10, 700, 16, t, size=8.6, color=TEXT, bold=True)
    text(s, 496, y + 28, 700, 24, d, size=7.2, color=MUTED, line_spacing=1.3)
    y += 66

metrics = [("3-detik hold", "≥ 65%"), ("Completion rate", "≥ 35%"),
           ("Share rate", "≥ 1,5%"), ("Follow / 1.000 views", "≥ 8")]
for i, (l, v) in enumerate(metrics):
    x = 64 + i * 292
    shape(s, x, 590, 276, 56, bg=WHITE, bg_alpha=0.025, border=WHITE,
          border_alpha=0.08, radius=9)
    text(s, x + 14, 599, 240, 12, l, size=6.2, color=MUTED2)
    text(s, x + 14, 614, 240, 24, v, size=13, color=TEXT, bold=True)
footer(s, 19)


# ============================================================ 20 ENGAGEMENT
s = new_slide()
header2(s, "Engagement", [
    [("Engagement bukan tombol suka. ", {}), ("Ia percakapan.", {"color": BLUEB})],
], tag="Bab 03 · Konten",
    lede="Algoritma memberi ruang pada konten yang memicu interaksi bermakna — komentar, simpan, bagikan, dan balasan.")

eng = [
    ("00–30 mnt", "Golden hour", "Balas setiap komentar dalam 30 menit pertama. Ajukan pertanyaan balik, bukan hanya “terima kasih”.", "red"),
    ("Komunitas", "Grup per angkatan", "WhatsApp/Telegram per angkatan dan program studi. 20 ambassador mahasiswa menjawab pertanyaan calon.", "blue"),
    ("UGC", "Kampus dalam lensa", "Repost story mahasiswa tiap hari. Beri kredit. Audiens lebih percaya pada teman daripada brosur.", "green"),
    ("Kolaborasi", "2 collab post/minggu", "UKM, alumni, kreator Aceh, bisnis lokal. Satu konten menjangkau dua audiens sekaligus.", "amber"),
    ("Live", "Satu live per minggu", "Q&A PMB, tur kampus, wawancara alumni. Live membangun kedekatan yang tidak bisa ditiru carousel.", "plain"),
    ("Giveaway", "Hadiah berbasis nilai", "Bootcamp, beasiswa mini, merchandise. Hindari skema “follow untuk menang” — melanggar kebijakan dan menarik audiens palsu.", "plain"),
]
for i, (tag, t, d, tone) in enumerate(eng):
    x = 64 + (i % 3) * 384
    y = 226 + (i // 3) * 126
    col = {"red": REDB, "blue": BLUEB, "green": GREEN, "amber": AMBER, "plain": MUTED}[tone]
    shape(s, x, y, 360, 114, bg=WHITE, bg_alpha=0.03, border=WHITE,
          border_alpha=0.09, radius=11)
    shape(s, x + 16, y + 14, max(54, 14 + len(tag) * 4.6), 20, bg=col,
          bg_alpha=0.13, border=col, border_alpha=0.28, radius=10)
    text(s, x + 16, y + 18, max(54, 14 + len(tag) * 4.6), 12, tag, size=6.4,
         color=col, bold=True, align="c")
    text(s, x + 16, y + 42, 326, 16, t, size=9.0, color=TEXT, bold=True)
    text(s, x + 16, y + 62, 328, 46, d, size=7.2, color=MUTED, line_spacing=1.35)

note(s, 64, 492, 1152,
     [("Efek compounding: ", {"bold": True, "color": TEXT}),
      ("balasan cepat → komentar naik → distribusi naik → follower naik → lebih banyak komentar. Ini loop yang harus dijaga setiap hari.", {})],
     h=48)
footer(s, 20)


# ============================================================ 21 CALENDAR
s = new_slide()
header2(s, "Kalender", [
    [("Satu minggu di feed UUI.", {})],
], tag="Bab 03 · Konten")

days = [
    ("Sen", "Karier", "“Template CV yang dilirik HRD.”", "Carousel · 11.00"),
    ("Sel", "Kehidupan", "“06.30 di kost, 08.00 di kelas.”", "Reels · 19.00"),
    ("Rab", "Bukti", "“Dari IPK 3,2 ke perusahaan teknologi.”", "Reels + FB · 20.00"),
    ("Kam", "Tren", "“POV: dosen buka presensi.”", "Reels · 12.00"),
    ("Jum", "Edukasi", "“Beasiswa yang jarang diketahui.”", "Carousel + Live · 15.00"),
    ("Sab", "Aceh", "“Kopi Aceh vs janji temu.”", "Reels · 10.00"),
    ("Min", "Komunitas", "“Foto kalian, kami posting.”", "Story + WA · 08.00"),
]
for i, (d, t, hook, fmt) in enumerate(days):
    x = 64 + i * 165
    shape(s, x, 222, 157, 152, bg=WHITE, bg_alpha=0.03, border=WHITE,
          border_alpha=0.09, radius=10)
    text(s, x + 12, 234, 60, 12, d.upper(), size=6.2, color=REDB, bold=True, spc=1.2)
    text(s, x + 12, 250, 130, 14, t, size=8.4, color=TEXT, bold=True)
    text(s, x + 12, 272, 130, 54, hook, size=7.0, color=MUTED, line_spacing=1.3)
    hline(s, x + 12, 340, 133, alpha=0.08)
    text(s, x + 12, 348, 132, 16, fmt, size=5.9, color=MUTED2)

ops = [
    ("SENIN", "Rapat 30 menit: evaluasi minggu lalu, kunci tema, bagi tugas produksi."),
    ("RABU", "Syuting batch 12–15 klip untuk dua minggu ke depan."),
    ("JUMAT", "Live 30 menit + dorong konten PMB. Ini hari konversi, bukan hari eksperimen."),
]
for i, (h, d) in enumerate(ops):
    x = 64 + i * 390
    shape(s, x, 396, 372, 88, bg=WHITE, bg_alpha=0.022, border=WHITE,
          border_alpha=0.08, radius=10)
    text(s, x + 16, 408, 330, 12, h, size=6.2, color=MUTED2, bold=True, spc=1.1)
    text(s, x + 16, 426, 340, 50, d, size=8.0, color=MUTED, line_spacing=1.35)

note(s, 64, 504, 1152,
     [("Output mingguan: ", {"bold": True, "color": TEXT}),
      ("7 Reels · 7 carousel/story · 21–35 Threads · 14 update WhatsApp — sekitar 28 aset unik yang dipotong ulang menjadi 300+ tayangan di empat platform.", {})],
     h=48)
footer(s, 21)


# ============================================================ 22 DIVIDER 04
s = new_slide("divider")
text(s, 84, 214, 330, 250, "04", size=172, color=RGBColor(0x2A, 0x34, 0x4A),
     bold=True, spc=-4)
shape(s, 90, 508, 96, 3, kind=MSO_SHAPE.RECTANGLE, bg=RED)
text(s, 430, 242, 300, 14, "TARGET 1 JUTA", size=7.5, color=REDB, bold=True, spc=2.2)
title_lines(s, [
    [("1.000.000 bukan", {})],
    [("keberuntungan. Itu hasil", {})],
    [("dari ", {}), ("kurva", {"color": REDB}), (" yang dihitung.", {})],
], x=430, y=270, w=800, size=42, h=200)
text(s, 430, 488, 620, 60,
     "Kita pecah angka besar menjadi angka harian yang bisa dikelola, lalu memberi setiap platform pekerjaan yang jelas.",
     size=11.5, color=MUTED, line_spacing=1.4)
footer(s, 22)

# ============================================================ 23 MATH
s = new_slide()
header2(s, "Matematika", [
    [("Angka besar, ", {}), ("dipecah jadi harian.", {"color": REDB})],
], tag="Bab 04 · 1 Juta")

mets = [("987.600", "FOLLOWER BARU\nYANG HARUS DIDAPAT", REDB),
        ("1.805", "RATA-RATA\nPER HARI", TEXT),
        ("54.867", "RATA-RATA\nPER BULAN", BLUEB)]
for i, (n, l, c) in enumerate(mets):
    x = 64 + i * 190
    if i:
        shape(s, x - 14, 232, 1, 66, kind=MSO_SHAPE.RECTANGLE, bg=WHITE, bg_alpha=0.10)
    text(s, x, 226, 180, 42, n, size=27, color=c, bold=True, spc=-0.8)
    tf = textbox(s, x, 272, 170, 30)
    for j, part in enumerate(l.split("\n")):
        p = tf.paragraphs[0] if j == 0 else tf.add_paragraph()
        p.line_spacing = 1.1
        _runs(p, part, 6.2, MUTED2, True, False, FONT, 1.0)

text(s, 64, 330, 540, 52,
     [("Perhitungan: ", {}), ("1.000.000 − 12.400 baseline = 987.600", {"color": WHITE, "font": MONO}),
      (" dibagi 547 hari. Angka ini bukan target harian yang datar — pertumbuhan mengikuti kurva: lambat di awal, cepat saat mesin konten dan iklan matang.", {})],
     size=7.4, color=MUTED, line_spacing=1.45)
note(s, 64, 398, 540,
     [("Baseline 12.400 adalah asumsi. ", {"bold": True, "color": TEXT}),
      ("Audit akun di hari pertama dan masukkan angka riil ke kalkulator versi web.", {})],
     h=64)

shape(s, 660, 205, 556, 250, bg=WHITE, bg_alpha=0.04, border=WHITE,
      border_alpha=0.10, radius=14)
text(s, 680, 220, 300, 18, "Kalkulator target", size=10.5, color=TEXT, bold=True)
pill(s, 1110, 219, 86, 22, "ILUSTRASI", color=MUTED, alpha=0.16, size=6)

fields = [("FOLLOWERS SAAT INI", "12.400"), ("TARGET FOLLOWERS", "1.000.000"),
          ("DURASI (BULAN)", "18")]
for i, (lab, val) in enumerate(fields):
    x = 680 + i * 172
    text(s, x, 254, 150, 10, lab, size=5.8, color=MUTED2, spc=0.6)
    shape(s, x, 268, 158, 34, bg=WHITE, bg_alpha=0.05, border=WHITE,
          border_alpha=0.12, radius=7)
    text(s, x + 9, 278, 140, 16, val, size=8.2, color=TEXT, bold=True, font=MONO)

hline(s, 680, 318, 516, alpha=0.10)
outs = [("1.805", "per hari"), ("54.867", "per bulan"), ("296.280", "via iklan (30%)")]
for i, (n, l) in enumerate(outs):
    x = 680 + i * 172
    text(s, x, 330, 160, 24, n, size=13.5, color=TEXT, bold=True, spc=-0.3)
    text(s, x, 354, 160, 12, l, size=6.2, color=MUTED2)

shape(s, 680, 382, 516, 30, bg=AMBER, bg_alpha=0.10, border=AMBER,
      border_alpha=0.22, radius=7)
text(s, 690, 389, 496, 16, "Sangat agresif · butuh mesin konten harian + paid scale",
     size=6.8, color=RGBColor(0xFF, 0xE0, 0xA3))

footer(s, 23)


# ============================================================ 24 ALLOCATION
s = new_slide()
header2(s, "Alokasi", [
    [("Membagi 1 juta ", {}), ("sesuai pekerjaan tiap platform.", {"color": BLUEB})],
], tag="Bab 04 · 1 Juta")

bars = [("Instagram", "480 rb", 48, 0.48, REDB, RED),
        ("Facebook", "250 rb", 25, 0.25, BLUEB, BLUE),
        ("Threads", "170 rb", 17, 0.17, RGBColor(0xB0, 0x8C, 0xFF), RGBColor(0x6B, 0x4B, 0xFF)),
        ("WhatsApp", "100 rb", 10, 0.10, GREEN, RGBColor(0x12, 0xA1, 0x50))]
y = 250
for name, val, pct, ratio, c1, c2 in bars:
    shape(s, 64, y + 9, 9, 9, bg=c1, radius=2)
    text(s, 82, y + 4, 120, 16, name, size=8.2, color=TEXT, bold=True)
    shape(s, 200, y + 7, 380, 10, bg=WHITE, bg_alpha=0.07, border=None, radius=5)
    shape(s, 200, y + 7, max(10, 380 * ratio), 10, bg=c1, border=None, radius=5)
    text(s, 590, y + 3, 110, 16, f"{val}  ·  {pct}%", size=8.2, color=TEXT, bold=True)
    y += 42

hline(s, 64, 432, 620, alpha=0.10)
for i, (n, l) in enumerate([("1.000.000", "Total audiens terafiliasi"),
                            ("4 platform", "Satu portofolio Meta"),
                            ("18 bulan", "6 kuartal pertumbuhan")]):
    x = 64 + i * 206
    text(s, x, 450, 200, 24, n, size=14, color=TEXT, bold=True)
    text(s, x, 474, 200, 16, l, size=6.8, color=MUTED2)

cards = [
    ("INSTAGRAM · 480 RB", "Mesin akuisisi utama. Reels menanggung 70% follower baru; carousel mengubah penonton menjadi pendaftar.", REDB),
    ("FACEBOOK · 250 RB", "Jangkauan 25+ dan orang tua calon mahasiswa. Reels lintas-posting menambah reach murah.", BLUEB),
    ("THREADS · 170 RB", "Pertumbuhan cepat karena biaya produksi rendah dan distribusi percakapan. Kunci: konsisten harian.", RGBColor(0xC3, 0xAA, 0xFF)),
    ("WHATSAPP · 100 RB", "Bukan untuk viral, tetapi untuk retensi dan konversi. Channel yang aktif mencegah kehilangan audiens.", RGBColor(0x8F, 0xF0, 0xB5)),
]
y = 226
for h, d, c in cards:
    shape(s, 720, y, 496, 84, bg=WHITE, bg_alpha=0.03, border=WHITE,
          border_alpha=0.09, radius=10)
    text(s, 738, y + 12, 450, 12, h, size=6.4, color=c, bold=True, spc=1.0)
    text(s, 738, y + 29, 460, 46, d, size=7.4, color=MUTED, line_spacing=1.36)
    y += 92
footer(s, 24)


# ============================================================ 25 GROWTH CURVE
s = new_slide()
header2(s, "Kurva", [
    [("Tiga fase, ", {}), ("satu lintasan.", {"color": REDB})],
], tag="Bab 04 · 1 Juta")

CL, CR, CT, CB = 120, 1120, 168, 424

def cx(m):
    return CL + (m / 18.0) * (CR - CL)

def cy(v):
    return CB - (v / 1000.0) * (CB - CT)

# phase bands
for m1, m2, label in [(0, 6, "FONDASI"), (6, 12, "AKSELERASI"), (12, 18, "SKALA")]:
    shape(s, cx(m1), CT, cx(m2) - cx(m1), CB - CT, bg=WHITE, bg_alpha=0.015,
          border=None, radius=0)

for v in [0, 250, 500, 750, 1000]:
    hline(s, CL, cy(v), CR - CL, alpha=0.08)
    text(s, CL - 14, cy(v) - 6, 52, 14,
         ("0" if v == 0 else ("1 jt" if v == 1000 else f"{v} rb")),
         size=6.4, color=MUTED2, align="r")

for m in [0, 3, 6, 9, 12, 15, 18]:
    shape(s, cx(m), CT, 1, CB - CT, kind=MSO_SHAPE.RECTANGLE, bg=WHITE, bg_alpha=0.04)
    text(s, cx(m) - 30, CB + 10, 60, 12, ("Mulai" if m == 0 else f"Bln {m}"),
         size=6.4, color=MUTED2, align="c")

milestones = [(0, 12.4), (1, 15), (3, 25), (6, 95), (9, 220), (12, 420), (15, 680), (18, 1000)]

def catmull(pts, samples=300):
    out = []
    for i in range(len(pts) - 1):
        p0 = pts[i - 1] if i > 0 else pts[i]
        p1 = pts[i]
        p2 = pts[i + 1]
        p3 = pts[i + 2] if i + 2 < len(pts) else p2
        for j in range(samples):
            t = j / samples
            t2, t3 = t * t, t * t * t
            x = 0.5 * ((2 * p1[0]) + (-p0[0] + p2[0]) * t +
                       (2 * p0[0] - 5 * p1[0] + 4 * p2[0] - p3[0]) * t2 +
                       (-p0[0] + 3 * p1[0] - 3 * p2[0] + p3[0]) * t3)
            y = 0.5 * ((2 * p1[1]) + (-p0[1] + p2[1]) * t +
                       (2 * p0[1] - 5 * p1[1] + 4 * p2[1] - p3[1]) * t2 +
                       (-p0[1] + 3 * p1[1] - 3 * p2[1] + p3[1]) * t3)
            out.append((x, y))
    out.append(pts[-1])
    return out

pts = [(cx(m), cy(v)) for m, v in milestones]
curve = catmull(pts)

def emu(v):
    return int(Inches(v / 96.0))

# area
fb = s.shapes.build_freeform(emu(curve[0][0]), emu(curve[0][1]))
fb.add_line_segments([(emu(x), emu(y)) for x, y in curve[1:]], close=False)
fb.add_line_segments([(emu(CR), emu(CB)), (emu(CL), emu(CB))], close=True)
area = fb.convert_to_shape()
fill(area, RED, 0.10)
area.line.fill.background()
no_shadow(area)

# line
fb2 = s.shapes.build_freeform(emu(curve[0][0]), emu(curve[0][1]))
fb2.add_line_segments([(emu(x), emu(y)) for x, y in curve[1:]], close=False)
cl = fb2.convert_to_shape()
cl.line.color.rgb = REDB
cl.line.width = Pt(2.4)
cl.fill.background()
no_shadow(cl)

for m, v in milestones[2:]:
    shape(s, cx(m) - 4, cy(v) - 4, 8, 8, kind=MSO_SHAPE.OVAL, bg=RGBColor(0x07, 0x0B, 0x14),
          border=REDB, border_alpha=1, bw=1.4)
    label = "1 juta" if m == 18 else f"{v} rb"
    text(s, cx(m) - 40, cy(v) - 21, 80, 14, str(label), size=7.0, color=WHITE,
         bold=True, align="c")
text(s, CR - 190, cy(1000) - 36, 190, 12, "TARGET 18 BULAN", size=6.2, color=BLUEB,
     align="r", spc=1.2)

phases = [
    ("BULAN 1–6", "Fondasi & pembuktian", "12 rb → 95 rb. Fokus kualitas, SOP, dan menemukan 3 format pemenang.", "red"),
    ("BULAN 7–12", "Akselerasi", "95 rb → 420 rb. Volume naik, kolaborasi mingguan, iklan mulai diskalakan.", "blue"),
    ("BULAN 13–18", "Skala", "420 rb → 1 juta. Loop viral + paid amplification + komunitas lintas kampus.", "amber"),
]
for i, (tag, t, d, tone) in enumerate(phases):
    x = 64 + i * 390
    shape(s, x, 478, 372, 100, bg=WHITE, bg_alpha=0.025, border=WHITE,
          border_alpha=0.08, radius=9)
    col = {"red": REDB, "blue": BLUEB, "amber": AMBER}[tone]
    pill(s, x + 14, 490, 62, 18, tag, color=col, border=col, alpha=0.25, size=5.8)
    text(s, x + 14, 514, 240, 14, t, size=8.6, color=TEXT, bold=True)
    text(s, x + 14, 532, 344, 38, d, size=7.0, color=MUTED, line_spacing=1.35)
footer(s, 25)


# ============================================================ 26 BUDGET
s = new_slide()
header2(s, "Anggaran", [
    [("Berapa harga ", {}), ("1 juta audiens?", {"color": REDB})],
], tag="Bab 04 · 1 Juta")

shape(s, 64, 212, 1152, 184, bg=WHITE, bg_alpha=0.02, border=WHITE,
      border_alpha=0.09, radius=14)
rows = [
    ["Lean", "Rp 15 jt", "135 rb", "400 rb", "535 rb", ("Tidak cukup", MUTED2)],
    ["Rekomendasi", "Rp 35 jt", "315 rb", "700 rb", "1,02 jt", ("Target tercapai", GREEN)],
    ["Agresif", "Rp 60 jt", "540 rb", "750 rb", "1,29 jt", ("Buffer 30%", BLUEB)],
]
table(s, 88, 224, 1104,
      ["Skenario", "Media / bulan", "Follower berbayar", "Follower organik", "Total 18 bulan", "Status"],
      rows, [170, 150, 180, 180, 160, 264], row_h=44, header_h=28, hl=1, size=8.2)

cards = [
    ("ASUMSI BIAYA", "Rp 2.000 per follower berbayar (rentang Rp 1.200–3.500). Diuji ulang setelah 30 hari."),
    ("MEDIA", "Rp 35 jt/bulan untuk akuisisi, retargeting, dan pengujian kreatif."),
    ("PRODUKSI", "Rp 9 jt/bulan: editor, properti, transport, honor kreator mahasiswa."),
    ("TOOLS", "Rp 2 jt/bulan: penjadwalan, analitik, aset stok, penyimpanan."),
]
for i, (h, d) in enumerate(cards):
    x = 64 + i * 292
    shape(s, x, 412, 276, 96, bg=WHITE, bg_alpha=0.025, border=WHITE,
          border_alpha=0.08, radius=10)
    text(s, x + 14, 424, 240, 12, h, size=6.2, color=MUTED2, bold=True, spc=1.1)
    text(s, x + 14, 442, 248, 58, d, size=7.6, color=MUTED, line_spacing=1.35)

note(s, 64, 526, 1152,
     [("Total yang diminta: ", {"bold": True, "color": TEXT}),
      (" ± Rp 46 jt/bulan (Rp 828 jt untuk 18 bulan). Itu setara Rp 828 per audiens terafiliasi untuk 18 bulan — dan sudah termasuk produksi konten.", {})],
     h=48)
footer(s, 26)


# ============================================================ 27 FUNNEL
s = new_slide()
header2(s, "Funnel", [
    [("Dari 15 juta tayangan ", {}), ("menjadi 55 ribu follower baru.", {"color": BLUEB})],
], tag="Bab 04 · 1 Juta")

stages = [("Reach / tayangan", "15.000.000", "per bulan", 1.00),
          ("Kunjungan profil", "450.000", "3%", 0.82),
          ("Follow baru", "55.000", "12%", 0.64),
          ("Terlibat", "30.000", "55%", 0.46),
          ("Advokat", "6.000", "20%", 0.30)]
y = 226
for i, (label, val, extra, ratio) in enumerate(stages):
    w = 560 * ratio
    x = 64 + (560 - w) / 2
    shape(s, x, y, w, 38, bg=BLUE if i < 2 else (RED if i < 4 else GREEN),
          bg_alpha=0.20 if i < 2 else 0.18, border=WHITE, border_alpha=0.14, radius=8)
    text(s, x + 14, y + 11, w * 0.5, 16, label, size=7.6, color=TEXT, bold=True)
    text(s, x + w - 170, y + 10, 156, 16,
         [(val + "  ", {"bold": True, "color": WHITE}), (extra, {"color": MUTED})],
         size=7.4, align="r")
    y += 46
text(s, 64, 468, 560, 28, "Angka adalah target bulan ke-12, bukan bulan pertama. Konversi diuji dan disesuaikan setiap minggu.",
     size=6.6, color=MUTED2, align="c", line_spacing=1.35)

# doughnut chart
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
cd = CategoryChartData()
cd.categories = ["Organik", "Berbayar"]
cd.add_series("Bauran", (70, 30))
gf = s.shapes.add_chart(XL_CHART_TYPE.DOUGHNUT, IN(700), IN(200), IN(210), IN(210), cd)
chart = gf.chart
chart.has_title = False
chart.has_legend = False
chart.font.size = Pt(7)
chart.font.color.rgb = MUTED
try:
    chart.format.fill.background()
    chart.format.line.fill.background()
except Exception:
    pass
plot = chart.plots[0]
plot.has_data_labels = False
ser = plot.series[0]
for idx, col in [(0, RED), (1, BLUE)]:
    pt = ser.points[idx]
    pt.format.fill.solid()
    pt.format.fill.fore_color.rgb = col
    pt.format.line.fill.background()
text(s, 720, 280, 170, 40, "70/30", size=17, color=TEXT, bold=True, align="c")

text(s, 946, 216, 150, 14, [("● ", {"color": RED}), ("70% organik", {"bold": True, "color": TEXT})], size=7.8)
text(s, 946, 234, 264, 40, "Konten harian, kolaborasi, komunitas, dan loop viral.", size=6.8, color=MUTED, line_spacing=1.3)
text(s, 946, 280, 150, 14, [("● ", {"color": BLUE}), ("30% berbayar", {"bold": True, "color": TEXT})], size=7.8)
text(s, 946, 298, 264, 40, "Amplifikasi konten pemenang, retargeting, dan akuisisi terukur.", size=6.8, color=MUTED, line_spacing=1.3)

hline(s, 700, 402, 516, alpha=0.10)
items = [("70", "Belanja terbukti", "Format yang sudah menurunkan biaya per follow."),
         ("20", "Scaling", "Naikkan 20% per 3 hari pada pemenang."),
         ("10", "Eksperimen", "Format, audiens, atau penawaran baru.")]
y = 416
for n, t, d in items:
    shape(s, 700, y, 516, 46, bg=WHITE, bg_alpha=0.025, border=WHITE,
          border_alpha=0.08, radius=8)
    shape(s, 712, y + 11, 24, 24, bg=RED if n == "70" else (BLUE if n == "20" else MUTED2), radius=6)
    text(s, 712, y + 16, 24, 12, n, size=6.6, color=WHITE, bold=True, align="c")
    text(s, 748, y + 7, 200, 14, t, size=7.4, color=TEXT, bold=True)
    text(s, 748, y + 22, 450, 20, d, size=6.4, color=MUTED)
    y += 52
footer(s, 27)


# ============================================================ 28 VIRAL LOOPS
s = new_slide()
header2(s, "Pertumbuhan", [
    [("Lima loop yang ", {}), ("bekerja saat kita tidur.", {"color": REDB})],
], tag="Bab 04 · 1 Juta")

loops = [
    ("1", "Loop Reels", "Tonton → bagikan → jangkauan baru → follower → lebih banyak yang menonton."),
    ("2", "Loop kolaborasi", "Satu collab post muncul di dua profil. Audiens silang tanpa biaya iklan."),
    ("3", "Loop komunitas", "WA Channel → pengumuman → mahasiswa bagikan → teman bergabung."),
    ("4", "Loop kampus", "50 mahasiswa kreator memproduksi konten harian. Biaya rendah, volume tinggi."),
    ("5", "Loop offline", "Event → liputan → konten → pendaftaran event berikutnya. Roda yang berputar."),
]
for i, (n, t, d) in enumerate(loops):
    x = 64 + i * 232
    shape(s, x, 226, 220, 152, bg=WHITE, bg_alpha=0.03, border=WHITE,
          border_alpha=0.09, radius=10)
    shape(s, x + 14, 240, 24, 24, bg=WHITE, bg_alpha=0.07, border=WHITE,
          border_alpha=0.12, radius=6)
    text(s, x + 14, 245, 24, 12, n, size=7.0, color=TEXT, bold=True, align="c")
    text(s, x + 14, 274, 190, 16, t, size=8.8, color=TEXT, bold=True)
    text(s, x + 14, 294, 192, 72, d, size=7.0, color=MUTED, line_spacing=1.38)

cards = [
    ("Program kreator mahasiswa", "50 kreator", "2 konten/minggu · honor Rp 150 rb/konten + merchandise. Rekrutmen lewat open call."),
    ("KOL & alumni", "10 KOL Aceh + 3 alumni nasional", "per kuartal. Konten otentik, bukan iklan scripted."),
    ("Momen besar", "Wisuda, orientasi, kompetisi, Ramadan.", "Siapkan content sprint 2 minggu sebelum momen."),
]
for i, (h, strong, d) in enumerate(cards):
    x = 64 + i * 390
    shape(s, x, 398, 372, 92, bg=WHITE, bg_alpha=0.022, border=WHITE,
          border_alpha=0.08, radius=10)
    text(s, x + 16, 410, 330, 12, h.upper(), size=6.2, color=MUTED2, bold=True, spc=1.1)
    text(s, x + 16, 428, 340, 54,
         [(strong + " ", {"bold": True, "color": TEXT}), (d, {})],
         size=7.4, color=MUTED, line_spacing=1.35)
footer(s, 28)


# ============================================================ 29 DIVIDER 05
s = new_slide("divider")
text(s, 84, 214, 330, 250, "05", size=172, color=RGBColor(0x2A, 0x34, 0x4A),
     bold=True, spc=-4)
shape(s, 90, 508, 96, 3, kind=MSO_SHAPE.RECTANGLE, bg=RED)
text(s, 430, 242, 300, 14, "EKSEKUSI", size=7.5, color=REDB, bold=True, spc=2.2)
title_lines(s, [
    [("Strategi tanpa ritme", {})],
    [("hanyalah daftar ", {}), ("keinginan.", {"color": REDB})],
], x=430, y=284, w=770, size=42, h=140)
text(s, 430, 446, 620, 60,
     "Tiga puluh hari pertama menentukan apakah rencana ini menjadi sistem atau sekadar dokumen.",
     size=11.5, color=MUTED, line_spacing=1.4)
footer(s, 29)


# ============================================================ 30 90 DAYS
s = new_slide()
header2(s, "90 hari pertama", [
    [("Dari keputusan menjadi ", {}), ("tayangan.", {"color": REDB})],
], tag="Bab 05 · Eksekusi")

phases = [
    ("HARI 1–30", "Fondasi", [
        "Audit 4 akun & tetapkan baseline",
        "Business Portfolio + verifikasi",
        "Pixel, CAPI, UTM, dashboard",
        "5 pilar konten & template visual",
        "Produksi batch 40 konten"], "red"),
    ("HARI 31–60", "Peluncuran", [
        "Kampanye awareness + traffic",
        "2 konten/hari, 1 live/minggu",
        "SOP balas komentar < 30 menit",
        "10 kolaborasi pertama",
        "Laporan mingguan pertama"], "blue"),
    ("HARI 61–90", "Optimasi", [
        "Matikan kreatif > biaya target",
        "Scale pemenang 20%/3 hari",
        "Uji audiens & penawaran baru",
        "Rekrut 50 kreator mahasiswa",
        "Review kuartal & rencana Q2"], "plain"),
]
for i, (tag, title, items, tone) in enumerate(phases):
    x = 64 + i * 390
    shape(s, x, 222, 372, 286, bg=RED if tone == "red" else (BLUE if tone == "blue" else WHITE),
          bg_alpha=0.11 if tone != "plain" else 0.025,
          border=REDB if tone == "red" else (BLUEB if tone == "blue" else WHITE),
          border_alpha=0.30 if tone != "plain" else 0.08, radius=12)
    tagcol = REDB if tone == "red" else (BLUEB if tone == "blue" else AMBER)
    pill(s, x + 18, 238, 74, 19, tag, color=tagcol, border=tagcol, alpha=0.28, size=5.8)
    text(s, x + 18, 266, 300, 20, title, size=12.5, color=TEXT, bold=True)
    y = 300
    for it in items:
        shape(s, x + 18, y + 2, 14, 14, bg=GREEN, bg_alpha=0.14, border=GREEN,
              border_alpha=0.30, radius=3)
        text(s, x + 20, y + 4, 10, 8, "✓", size=5.6, color=GREEN, bold=True)
        text(s, x + 40, y, 310, 24, it, size=7.6, color=MUTED, line_spacing=1.25)
        y += 36

note(s, 64, 536, 1152,
     [("Gerbang keputusan hari ke-90: ", {"bold": True, "color": TEXT}),
      ("bila pertumbuhan < 8.000 follower/bulan, ubah hook dan penawaran — jangan menambah anggaran. Uang tidak memperbaiki pesan yang salah.", {})],
     h=48)
footer(s, 30)

# ============================================================ 31 TEAM
s = new_slide()
header2(s, "Tim", [
    [("Tim kecil, ", {}), ("tanggung jawab jelas.", {"color": BLUEB})],
], tag="Bab 05 · Eksekusi")

shape(s, 64, 214, 760, 310, bg=WHITE, bg_alpha=0.02, border=WHITE,
      border_alpha=0.09, radius=14)
team = [
    ["Social Media Lead", "Strategi, kalender, kualitas, persetujuan", "1"],
    ["Content Strategist", "Riset tren, hook, naskah, pilar konten", "1"],
    ["Videographer / Editor", "Produksi & editing Reels, live", "2"],
    ["Designer", "Carousel, template, identitas visual", "1"],
    ["Copywriter", "Caption, Threads, script, siaran pers", "1"],
    ["Community Manager", "Komentar, DM, grup, ambassador", "1"],
    ["Performance Marketer", "Meta Ads, anggaran, optimasi", "1"],
    ["Data & Reporting", "Dashboard, insight mingguan", "0,5"],
]
table(s, 88, 228, 712, ["Peran", "Tanggung jawab utama", "FTE"],
      team, [210, 410, 92], row_h=34, header_h=26, size=7.8)

shape(s, 840, 214, 376, 104, bg=BLUE, bg_alpha=0.13, border=BLUEB,
      border_alpha=0.30, radius=10)
text(s, 858, 226, 330, 12, "TOTAL", size=6.2, color=RGBColor(0xA9, 0xDB, 0xFF), bold=True, spc=1.2)
text(s, 858, 242, 200, 38, "8,5", size=26, color=TEXT, bold=True)
text(s, 858, 286, 340, 18, "FTE — kombinasi staf tetap, magang, dan vendor.", size=6.8, color=MUTED)

for i, (h, d) in enumerate([
        ("YANG TIDAK BOLEH DIRANGKAP", "Production dan approval. Orang yang membuat konten tidak boleh menjadi satu-satunya yang menyetujui konten."),
        ("AMBASSADOR", "20 mahasiswa berperan sebagai mata dan telinga di kampus. Bukan admin, tetapi sumber cerita dan jawaban.")]):
    y = 330 + i * 100
    shape(s, 840, y, 376, 92, bg=WHITE, bg_alpha=0.025, border=WHITE,
          border_alpha=0.08, radius=10)
    text(s, 858, y + 12, 330, 12, h, size=6.2, color=MUTED2, bold=True, spc=1.1)
    text(s, 858, y + 30, 340, 56, d, size=7.2, color=MUTED, line_spacing=1.35)
footer(s, 31)


# ============================================================ 32 KPI
s = new_slide()
header2(s, "KPI", [
    [("Delapan angka yang ", {}), ("dilihat setiap Senin.", {"color": REDB})],
], tag="Bab 05 · Eksekusi")

kpis = [
    ("Pertumbuhan bersih", "+55 rb", "follower/bulan · bulan 12"),
    ("Reach bulanan", "15 jt", "akun unik · semua platform"),
    ("Engagement rate", "≥ 6%", "per reach · Instagram"),
    ("Simpan + bagikan", "≥ 2%", "sinyal kualitas konten"),
    ("Biaya per follower", "≤ Rp 2.500", "blended · semua kampanye"),
    ("Waktu respons", "< 30 mnt", "komentar & DM jam kerja"),
    ("Sentimen positif", "≥ 85%", "social listening bulanan"),
    ("Pendaftar dari sosial", "+20%", "atribusi UTM · per kuartal"),
]
for i, (l, v, d) in enumerate(kpis):
    x = 64 + (i % 4) * 292
    y = 222 + (i // 4) * 112
    shape(s, x, y, 276, 98, bg=WHITE, bg_alpha=0.03, border=WHITE,
          border_alpha=0.09, radius=10)
    text(s, x + 14, y + 12, 240, 12, l, size=6.6, color=MUTED2)
    text(s, x + 14, y + 30, 248, 30, v, size=16, color=TEXT, bold=True, spc=-0.3)
    text(s, x + 14, y + 64, 248, 26, d, size=6.2, color=MUTED2)

cards = [
    ("DASHBOARD", "Looker Studio: satu halaman, empat platform, data harian. Bukan laporan 30 halaman."),
    ("ATRIBUSI", "UTM wajib: utm_source · medium · campaign · content. Tanpa ini, kontribusi sosial tidak terlihat."),
    ("DEFINISI MENANG", "Bukan follower terbanyak, tetapi audiens yang bertanya, mendaftar, dan merekomendasikan."),
]
for i, (h, d) in enumerate(cards):
    x = 64 + i * 390
    shape(s, x, 456, 372, 90, bg=WHITE, bg_alpha=0.022, border=WHITE,
          border_alpha=0.08, radius=10)
    text(s, x + 16, 468, 330, 12, h, size=6.2, color=MUTED2, bold=True, spc=1.1)
    text(s, x + 16, 486, 340, 52, d, size=7.4, color=MUTED, line_spacing=1.35)
footer(s, 32)


# ============================================================ 33 CADENCE
s = new_slide()
header2(s, "Ritme", [
    [("Irama yang membuat ", {}), ("strategi tetap hidup.", {"color": BLUEB})],
], tag="Bab 05 · Eksekusi")

cad = [
    ("Harian", "15 menit berdiri", "Konten tayang? Komentar terbalas? Iklan sehat? Satu keputusan, satu pemilik.", "08.00 WIB · semua tim", "red"),
    ("Mingguan", "Review performa", "Senin 60 menit: metrik, kreatif pemenang, kill list, rencana konten pekan ini.", "Senin · lead + performance", "blue"),
    ("Bulanan", "Business review", "Follower, biaya, pendaftar, sentimen. Satu eksperimen besar per bulan.", "Awal bulan · pimpinan", "amber"),
    ("Kuartalan", "Reset strategi", "Audit pilar konten, alokasi platform, dan struktur anggaran kuartal berikutnya.", "Setiap 3 bulan · pimpinan", "green"),
]
for i, (tag, t, d, when, tone) in enumerate(cad):
    x = 64 + i * 292
    col = {"red": REDB, "blue": BLUEB, "amber": AMBER, "green": GREEN}[tone]
    shape(s, x, 226, 276, 200, bg=WHITE, bg_alpha=0.03, border=WHITE,
          border_alpha=0.09, radius=11)
    pill(s, x + 16, 242, max(54, 12 + len(tag) * 6), 20, tag, color=col,
         border=col, alpha=0.28, size=6.2)
    text(s, x + 16, 274, 240, 36, t, size=10.5, color=TEXT, bold=True, line_spacing=1.1)
    text(s, x + 16, 316, 244, 70, d, size=7.4, color=MUTED, line_spacing=1.38)
    text(s, x + 16, 392, 244, 24, when, size=6.0, color=MUTED2)

note(s, 64, 452, 1152,
     [("Aturan satu halaman: ", {"bold": True, "color": TEXT}),
      ("setiap laporan mingguan muat dalam satu layar — 3 angka naik, 3 angka turun, 3 tindakan minggu depan. Jika tidak muat, itu bukan ringkasan.", {})],
     h=48)
footer(s, 33)


# ============================================================ 34 RISKS
s = new_slide()
header2(s, "Risiko", [
    [("Yang bisa menggagalkan rencana — ", {}), ("dan penawarnya.", {"color": REDB})],
], tag="Bab 05 · Eksekusi")

rows = [
    ["Akun iklan dibatasi", ("Tinggi", REDB), "Verifikasi bisnis, dua admin, kreatif patuh kebijakan, tidak ada klaim berlebihan."],
    ["Follower palsu / engagement dibeli", ("Tinggi", REDB), "Larangan total. Tidak ada bot, tidak ada giveaway follow-only. Audit kualitas audiens bulanan."],
    ["Kelelahan kreatif", ("Sedang", AMBER), "Bank 14 hari konten, 3 varian hook, rotasi format tiap 2 minggu."],
    ["Krisis reputasi", ("Tinggi", REDB), "SOP eskalasi 1 jam, satu juru bicara, tanggapan tertulis sebelum komentar dibuka."],
    ["Anggaran tidak efisien", ("Sedang", AMBER), "Aturan kill/scale mingguan, frekuensi < 3, biaya per follow jadi KPI utama."],
    ["Ketergantungan satu platform", ("Rendah", BLUEB), "Diversifikasi ke WhatsApp Channel, email, dan situs. Bangun audiens yang bisa dimiliki."],
]
shape(s, 64, 214, 1152, 286, bg=WHITE, bg_alpha=0.02, border=WHITE,
      border_alpha=0.09, radius=14)
table(s, 88, 226, 1104, ["Risiko", "Dampak", "Mitigasi"],
      rows, [330, 130, 644], row_h=42, header_h=28, size=7.8)

note(s, 64, 522, 1152,
     [("UU PDP: ", {"bold": True, "color": TEXT}),
      ("setiap pengumpulan data lewat form iklan wajib punya dasar hukum, consent, dan kebijakan privasi yang bisa diakses. Ini bukan urusan tim IT saja.", {})],
     h=48)
footer(s, 34)


# ============================================================ 35 GOVERNANCE
s = new_slide()
header2(s, "Tata kelola", [
    [("Bebas berkreasi, ", {}), ("ketat pada prinsip.", {"color": BLUEB})],
], tag="Bab 05 · Eksekusi")

gov = [
    ("KEBIJAKAN IKLAN", "Tidak ada klaim “kampus terbaik” tanpa bukti. Tidak ada sebelum-sesudah yang menyesatkan. Tidak ada penargetan berbasis atribut pribadi.", "Semua kreatif melewati checklist 6 poin sebelum publish."),
    ("PRIVASI & DATA", "Consent banner, kebijakan privasi, retensi data 12 bulan, dan larangan mengirim NIK/nomor HP ke Meta.", "Selaras dengan UU PDP dan kebijakan Meta."),
    ("KEAMANAN AKUN", "2FA wajib, dua admin, audit akses kuartalan, dan prosedur pemulihan tertulis bila akun dibatasi.", "Tidak ada password yang dibagikan lewat grup chat."),
]
for i, (h, d, foot) in enumerate(gov):
    x = 64 + i * 390
    shape(s, x, 222, 372, 176, bg=WHITE, bg_alpha=0.03, border=WHITE,
          border_alpha=0.09, radius=11)
    text(s, x + 16, 238, 330, 12, h, size=6.4, color=MUTED2, bold=True, spc=1.2)
    text(s, x + 16, 258, 340, 80, d, size=7.8, color=MUTED, line_spacing=1.4)
    hline(s, x + 16, 346, 340, alpha=0.10)
    text(s, x + 16, 356, 340, 30, foot, size=6.4, color=MUTED2, line_spacing=1.3)

flows = [
    ("ALUR PERSETUJUAN KONTEN", "Ide → naskah → produksi → review fakta & kebijakan → jadwal → tayang. Maksimal 2 hari, bukan 2 minggu."),
    ("ALUR KRISIS", "Deteksi < 15 menit → verifikasi fakta → eskalasi ke pimpinan → pernyataan resmi < 2 jam → evaluasi."),
]
for i, (h, d) in enumerate(flows):
    x = 64 + i * 580
    w = 564 if i == 0 else 572
    shape(s, x, 414, w, 90, bg=WHITE, bg_alpha=0.022, border=WHITE,
          border_alpha=0.08, radius=10)
    text(s, x + 16, 426, 500, 12, h, size=6.2, color=MUTED2, bold=True, spc=1.1)
    text(s, x + 16, 444, w - 32, 52, d, size=7.6, color=MUTED, line_spacing=1.35)
footer(s, 35)


# ============================================================ 36 CLOSING
s = new_slide("cover")
header2(s, "Penutup", [
    [("Lima keputusan ", {}), ("yang diminta hari ini.", {"color": REDB})],
])

dec = [
    ("1", "Setujui anggaran Rp 46 jt/bulan", "Rp 35 jt media + Rp 9 jt produksi + Rp 2 jt tools, dievaluasi per kuartal."),
    ("2", "Tunjuk satu pemilik akun", "Social Media Lead dengan mandat lintas biro, bukan komite."),
    ("3", "Lengkapi dokumen verifikasi", "NIB, akta, SK, dan bukti alamat diserahkan dalam 7 hari."),
    ("4", "Aktifkan studio konten", "Ruang, kamera, lampu, dan dua editor. Volume dimulai dari sini."),
    ("5", "Setujui target per kuartal", "Q1 25 rb · Q2 95 rb · Q3 220 rb · Q4 420 rb · Q5 680 rb · Q6 1 juta."),
]
y = 214
for n, t, d in dec:
    shape(s, 64, y, 600, 58, bg=WHITE, bg_alpha=0.03, border=WHITE,
          border_alpha=0.09, radius=9)
    shape(s, 78, y + 15, 26, 26, bg=RED, radius=6)
    text(s, 78, y + 21, 26, 12, n, size=7.4, color=WHITE, bold=True, align="c")
    text(s, 116, y + 9, 530, 16, t, size=8.6, color=TEXT, bold=True)
    text(s, 116, y + 27, 530, 26, d, size=7.0, color=MUTED, line_spacing=1.3)
    y += 66

shape(s, 688, 210, 528, 248, bg=RED, bg_alpha=0.10, border=REDB,
      border_alpha=0.28, radius=12)
text(s, 706, 226, 400, 12, "TUJUH HARI KE DEPAN", size=6.2, color=RGBColor(0xFF, 0xB3, 0xBA), bold=True, spc=1.2)
week = ["Audit keempat akun & tetapkan baseline", "Buat Business Portfolio & ajukan verifikasi",
        "Pasang Pixel, CAPI, dan UTM", "Kunci 5 pilar & template visual",
        "Syuting batch pertama: 30 konten", "Jadwalkan review pertama, Senin"]
y = 250
for it in week:
    shape(s, 706, y + 1, 14, 14, bg=GREEN, bg_alpha=0.14, border=GREEN,
          border_alpha=0.30, radius=3)
    text(s, 708, y + 3, 10, 8, "✓", size=5.6, color=GREEN, bold=True)
    text(s, 728, y - 2, 470, 20, it, size=7.6, color=MUTED)
    y += 30

q = textbox(s, 688, 480, 530, 96)
p = q.paragraphs[0]
p.line_spacing = 1.3
_runs(p, "“Kita tidak mengejar angka.", 14.5, TEXT, False, False, FONT, 0)
p2 = q.add_paragraph()
p2.line_spacing = 1.3
_runs(p2, [("Kita membangun alasan", {"color": REDB}),
           (" bagi orang untuk mengikuti, mempercayai, dan memilih UUI.”", {})],
      14.5, TEXT, False, False, FONT, 0)
footer(s, 36)


# ============================================================ 37 APPENDIX
s = new_slide()
header2(s, "Lampiran", [
    [("Standar kecil yang ", {}), ("menjaga konsistensi.", {"color": BLUEB})],
])

shape(s, 64, 226, 360, 200, bg=WHITE, bg_alpha=0.03, border=WHITE,
      border_alpha=0.09, radius=11)
text(s, 82, 242, 320, 12, "KONVENSI PENAMAAN", size=6.4, color=MUTED2, bold=True, spc=1.2)
tf = textbox(s, 82, 264, 326, 90)
lines = [("[UUI] [Tujuan] [Audiens] [Bulan]", MUTED2),
         ("UUI-Traffic-PMB-SiswaSMA-Nov26", TEXT),
         ("UUI-Engagement-Alumni-Q4-26", MUTED),
         ("UUI-Awareness-Brand-Des26", MUTED)]
for i, (ln, c) in enumerate(lines):
    p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
    p.line_spacing = 1.55
    _runs(p, ln, 6.8, c, False, False, MONO, 0)
hline(s, 82, 366, 324, alpha=0.10)
text(s, 82, 378, 320, 36, "Konsisten = laporan bisa dibaca tanpa bertanya ke pembuatnya.",
     size=6.6, color=MUTED2, line_spacing=1.3)

shape(s, 448, 226, 360, 200, bg=WHITE, bg_alpha=0.03, border=WHITE,
      border_alpha=0.09, radius=11)
text(s, 466, 242, 320, 12, "UTM WAJIB", size=6.4, color=MUTED2, bold=True, spc=1.2)
tf = textbox(s, 466, 264, 326, 90)
for i, ln in enumerate(["utm_source = instagram", "utm_medium = social",
                        "utm_campaign = pmb_2026", "utm_content = reels_hook_a"]):
    p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
    p.line_spacing = 1.55
    _runs(p, ln, 6.8, MUTED, False, False, MONO, 0)
hline(s, 466, 366, 324, alpha=0.10)
text(s, 466, 378, 320, 36, "Setiap link di bio, story, dan iklan memakai UTM yang berbeda.",
     size=6.6, color=MUTED2, line_spacing=1.3)

shape(s, 832, 226, 384, 200, bg=WHITE, bg_alpha=0.03, border=WHITE,
      border_alpha=0.09, radius=11)
text(s, 850, 242, 340, 12, "CHECKLIST ASET MINGGUAN", size=6.4, color=MUTED2, bold=True, spc=1.2)
for i, it in enumerate(["14 konten terjadwal", "3 varian hook per ide", "Subtitle di semua video",
                        "Link bio diperbarui", "Laporan 1 halaman terkirim"]):
    y = 266 + i * 26
    shape(s, 850, y + 2, 14, 14, bg=GREEN, bg_alpha=0.14, border=GREEN,
          border_alpha=0.30, radius=3)
    text(s, 852, y + 4, 10, 8, "✓", size=5.6, color=GREEN, bold=True)
    text(s, 872, y, 330, 20, it, size=7.4, color=MUTED)

hline(s, 64, 456, 1152, alpha=0.10)
text(s, 64, 472, 800, 54,
     "Dokumen ini adalah rencana kerja, bukan jaminan hasil. Setiap angka perlu diverifikasi terhadap data aktual akun, perilaku audiens, dan kondisi pasar. Semua keputusan anggaran tetap melalui persetujuan pimpinan.",
     size=7.2, color=MUTED2, line_spacing=1.42)
s.shapes.add_picture(ASSET("uui-logo.png"), IN(936), IN(500), height=IN(26))
shape(s, 1056, 500, 1, 24, kind=MSO_SHAPE.RECTANGLE, bg=WHITE, bg_alpha=0.18)
s.shapes.add_picture(ASSET("meta-logo-white.png"), IN(1076), IN(505), height=IN(13))
footer(s, 37)


# ---------------------------------------------------------------- speaker notes
import re

def load_narration():
    path = os.path.join(ROOT, "NARASI-PRESENTASI.md")
    if not os.path.exists(path):
        return []
    raw = open(path, encoding="utf-8").read()
    chunks = re.split(r"^## (\d{2}) · ", raw, flags=re.M)[1:]
    out = []
    for i in range(0, len(chunks), 2):
        body = chunks[i + 1]
        m = re.search(r"\*\*Narasi:\*\*(.*?)\*\*Poin kunci:\*\*", body, re.S)
        if not m:
            out.append("")
            continue
        txt = m.group(1)
        txt = re.sub(r"\*\[.*?\]\*", "", txt, flags=re.S)  # stage directions
        txt = txt.replace("**", "")
        lines = [ln.strip() for ln in txt.strip().split("\n")]
        paras = []
        buf = []
        for ln in lines:
            if ln:
                buf.append(ln)
            elif buf:
                paras.append(" ".join(buf))
                buf = []
        if buf:
            paras.append(" ".join(buf))
        out.append("\n\n".join(paras))
    return out

narration = load_narration()
if len(narration) >= len(prs.slides._sldIdLst):
    for slide, note_text in zip(prs.slides, narration):
        if note_text:
            slide.notes_slide.notes_text_frame.text = note_text

prs.save(os.path.join(ROOT, "UUI-Meta-Ads-2026.pptx"))
print("saved", len(prs.slides._sldIdLst), "slides,", len(narration), "notes")
