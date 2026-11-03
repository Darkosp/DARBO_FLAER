# -*- coding: utf-8 -*-
"""DARBO GRUP — A4 flyer (MK / AL), 2 pages."""
from pptx import Presentation
from pptx.util import Cm, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

INK    = "1E2529"
ACCENT = "C97B2A"
LIGHT  = "F2F3F4"
MID    = "6E767B"
DIM    = "D8DCDE"
WHITE  = "FFFFFF"
FONT   = "Arial"

W, H = 21.0, 29.7          # A4 portrait, cm
ML   = 1.30                # left margin
CW   = 18.4                # content width


def rgb(h):
    return RGBColor.from_string(h)


def rect(slide, x, y, w, h, fill=None, line=None, lw=0.6):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Cm(x), Cm(y), Cm(w), Cm(h))
    shp.shadow.inherit = False
    if fill:
        shp.fill.solid(); shp.fill.fore_color.rgb = rgb(fill)
    else:
        shp.fill.background()
    if line:
        shp.line.color.rgb = rgb(line); shp.line.width = Pt(lw)
    else:
        shp.line.fill.background()
    return shp


def tb(slide, x, y, w, h, paras, anchor=MSO_ANCHOR.TOP, wrap=True):
    """paras: list of dicts {runs:[(text, dict(...))], align, sa, sb, ls}"""
    box = slide.shapes.add_textbox(Cm(x), Cm(y), Cm(w), Cm(h))
    tf = box.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for i, p in enumerate(paras):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = p.get("align", PP_ALIGN.LEFT)
        if p.get("sa") is not None:
            para.space_after = Pt(p["sa"])
        if p.get("sb") is not None:
            para.space_before = Pt(p["sb"])
        if p.get("ls") is not None:
            para.line_spacing = p["ls"]
        for text, st in p["runs"]:
            r = para.add_run()
            r.text = text
            f = r.font
            f.name = st.get("font", FONT)
            f.size = Pt(st.get("sz", 8.5))
            f.bold = st.get("b", False)
            f.italic = st.get("i", False)
            f.color.rgb = rgb(st.get("c", INK))
    return box


def hanging(para, cm=0.42):
    pPr = para._p.get_or_add_pPr()
    pPr.set("marL", str(int(Cm(cm))))
    pPr.set("indent", str(int(-Cm(cm))))


def bullet(para, char="•", color=None, cm=0.42):
    hanging(para, cm)
    pPr = para._p.get_or_add_pPr()
    if color:
        buClr = pPr.makeelement(qn("a:buClr"), {})
        srgb = buClr.makeelement(qn("a:srgbClr"), {"val": color})
        buClr.append(srgb); pPr.append(buClr)
    buFont = pPr.makeelement(qn("a:buFont"), {"typeface": FONT, "pitchFamily": "34", "charset": "0"})
    pPr.append(buFont)
    buChar = pPr.makeelement(qn("a:buChar"), {"char": char})
    pPr.append(buChar)


# ----------------------------------------------------------------------------- content
MK = dict(
    head_sub="Лиценциран управител на станбени објекти  ·  Тетово  ·  од 2013",
    lic_top="ЛИЦЕНЦА  Уп.0812/3",
    lic_sub="Регулаторна комисија за домување  ·  важи до 10.03.2028",
    kicker="УПРАВУВАЊЕ СО СТАНБЕНИ ЗГРАДИ",
    headline=["Колку чини вашата зграда", "— на крајот на месецот?"],
    sub="Една цена за управување, чистење и техничко одржување. "
        "Сите останати ставки ги гледате однапред — ставка по ставка, секој месец.",
    big="500 ден.",
    unit=["по стан", "месечно"],
    col1="ШТО Е ВКЛУЧЕНО ВО 500 ден.",
    inc=[
        "Чистење двапати неделно — суво и влажно, со средства за хигиена",
        "Техничка контрола еднаш неделно: електрика, водовод, хидрофор, домофон",
        "Работна рака за замена на сијалици и мали поправки",
        "Организација на месечен сервис на лифтот и годишен технички преглед",
        "Чистење на оџаци наесен",
        "Управување: фактури, договори, годишен извештај, состаноци со сопственици",
        "MyDarbo: сметки, пријава на дефект со фотографија и е-плаќање",
    ],
    col2="ШТО СЕ ПЛАЌА ПОСЕБНО",
    exc=[
        ("Резервен фонд — 1 ден. за м²", "Законска обврска, на посебна сметка на зградата."),
        ("Лифт — по набавна цена", "Договорот е на зградата: околу 3.000 ден. месечно."),
        ("Струја за заеднички простории", "По реална потрошувачка, распределена на сите станови."),
        ("Поправки и материјал", "Со фактура, само кога има потреба."),
    ],
    exc_note="Немаме маржа на овие четири ставки. Нив ги гледате во MyDarbo, ставка по ставка.",
    ex_title="ПРИМЕР: ЗГРАДА СО 30 СТАНА, СТАН ОД 50 м²",
    rows=[
        ("Управување, чистење и техничко одржување", "500"),
        ("Резервен фонд — 50 м² × 1 ден.", "50"),
        ("Лифт — 3.000 ден. ÷ 30 стана", "100"),
        ("Струја за заеднички простории — 5.000 ден. ÷ 30", "167"),
    ],
    total=("ВКУПНО МЕСЕЧНО", "817 ден."),
    stats=[
        ("100+", "згради под\nуправување"),
        ("2.500+", "домаќинства"),
        ("13", "години на\nпазарот"),
        ("15+", "вработени,\nсопствени тимови"),
    ],
    chk_title="ПРАШАЈТЕ ЈА СЕКОЈА ПОНУДА — ПЕТ ПРАШАЊА",
    chk=[
        "Дали „од 399 ден.“ е конечната сметка, или има плус за лифт, струја и резервен фонд?",
        "Кој ја плаќа струјата на ходниците и како се дели?",
        "Колку пати неделно реално доаѓа екипа — и кој го проверува тоа?",
        "Каде седи резервниот фонд, на чие име и кој има пристап?",
        "Има ли управителот активна лиценца од Регулаторната комисија за домување?",
    ],
    fine="Лифтот и струјата по стан зависат од бројот на станови и од фактурите — "
         "затоа ги покажуваме отворено.",
    cta1="Бесплатна проценка",
    cta2="на вашата зграда",
    cta3="и споредба со сегашната понуда — за 3 работни дена.",
    contact=["Дарко Спировски, овластен управител",
             "+389 78 353 765  ·  Viber / WhatsApp",
             "contact@darbo.mk  ·  darbo.mk"],
    qr1="Провери во\nРКД регистарот",
    qr2="Плати сметка\nонлајн",
    footer="Дарбо Груп ДОО  ·  ул. Маршал Тито 166/2, 1200 Тетово  ·  Лиценца Уп.0812/3  ·  darbo.mk",
)

AL = dict(
    head_sub="Menaxher i licencuar i ndërtesave banesore  ·  Tetovë  ·  nga 2013",
    lic_top="LICENCA  Up.0812/3",
    lic_sub="Komisioni Rregullator për Banesa  ·  e vlefshme deri 10.03.2028",
    kicker="MENAXHIM I NDËRTESAVE BANESORE",
    headline=["Sa kushton ndërtesa juaj", "— në fund të muajit?"],
    sub="Një çmim për menaxhim, pastrim dhe mirëmbajtje teknike. "
        "Të gjitha zërat e tjerë i shihni paraprakisht — zë për zë, çdo muaj.",
    big="500 den.",
    unit=["për banesë", "në muaj"],
    col1="ÇKA PËRFSHIHET NË 500 den.",
    inc=[
        "Pastrim dy herë në javë — thatë dhe i lagësht, me mjete higjienike",
        "Kontroll teknik një herë në javë: rryma, uji, hidrofori, interkomi",
        "Punëtor për zëvendësimin e llambave dhe riparime të vogla",
        "Organizim i servisit mujor të ashensorit dhe kontrollit vjetor teknik",
        "Pastrim i oxhaqeve në vjeshtë",
        "Menaxhim: fatura, kontrata, raport vjetor, mbledhje me pronarët",
        "MyDarbo: fatura, raportim defekti me fotografi dhe pagesë online",
    ],
    col2="ÇKA PAGUHET VEÇMAS",
    exc=[
        ("Fondi rezervë — 1 den. për m²", "Detyrim ligjor, në llogari të veçantë të ndërtesës."),
        ("Ashensori — me çmim kosto", "Kontrata është e ndërtesës: rreth 3.000 den. në muaj."),
        ("Rryma e hapësirave të përbashkëta", "Sipas konsumit real, e ndarë mes të gjitha banesave."),
        ("Riparime dhe material", "Me faturë, vetëm kur ka nevojë."),
    ],
    exc_note="Nuk marrim marzhë mbi këta katër zëra. Ata i shihni në MyDarbo, zë për zë.",
    ex_title="SHEMBULL: NDËRTESË ME 30 BANESA, BANESË 50 m²",
    rows=[
        ("Menaxhim, pastrim dhe mirëmbajtje teknike", "500"),
        ("Fondi rezervë — 50 m² × 1 den.", "50"),
        ("Ashensori — 3.000 den. ÷ 30 banesa", "100"),
        ("Rryma e hapësirave të përbashkëta — 5.000 den. ÷ 30", "167"),
    ],
    total=("TOTALI MUJOR", "817 den."),
    stats=[
        ("100+", "ndërtesa nën\nmenaxhim"),
        ("2.500+", "familje"),
        ("13", "vjet në\ntreg"),
        ("15+", "të punësuar,\nekipe të vetat"),
    ],
    chk_title="PYESNI ÇDO OFERTË — PESË PYETJE",
    chk=[
        "A është „nga 399 den.“ fatura finale, apo ka plus për ashensorin, rrymën dhe fondin rezervë?",
        "Kush e paguan rrymën e korridoreve dhe si ndahet?",
        "Sa herë në javë vjen realisht ekipi — dhe kush e kontrollon atë?",
        "Ku mbahet fondi rezervë, në emër të kujt dhe kush ka qasje?",
        "A ka menaxheri licencë aktive nga Komisioni Rregullator për Banesa?",
    ],
    fine="Ashensori dhe rryma për banesë varen nga numri i banesave dhe nga faturat — "
         "prandaj i tregojmë hapur.",
    cta1="Vlerësim falas",
    cta2="për ndërtesën tuaj",
    cta3="dhe krahasim me ofertën aktuale — brenda 3 ditëve pune.",
    contact=["Darko Spirovski, menaxher i autorizuar",
             "+389 78 353 765  ·  Viber / WhatsApp",
             "contact@darbo.mk  ·  darbo.mk"],
    qr1="Verifiko në\nregjistrin e KRB",
    qr2="Paguaj faturën\nonline",
    footer="Darbo Grup DOO  ·  rr. Marshall Tito 166/2, 1200 Tetovë  ·  Licenca Up.0812/3  ·  darbo.mk",
)


def build(c, slide):
    ink, acc, mid, dim, white = INK, ACCENT, MID, DIM, WHITE

    # ── header band ───────────────────────────────────────────────────────────
    rect(slide, 0, 0, W, 2.10, ink)
    tb(slide, ML, 0.40, 10.0, 1.00,
       [{"runs": [("DARBO GRUP", {"sz": 21, "b": True, "c": white})]}], MSO_ANCHOR.MIDDLE)
    tb(slide, ML, 1.36, 12.0, 0.50,
       [{"runs": [(c["head_sub"], {"sz": 8.2, "c": dim})]}])
    tb(slide, 11.60, 0.52, 8.10, 0.45,
       [{"runs": [(c["lic_top"], {"sz": 9.5, "b": True, "c": acc})], "align": PP_ALIGN.RIGHT}])
    tb(slide, 10.20, 1.02, 9.50, 0.45,
       [{"runs": [(c["lic_sub"], {"sz": 7.1, "c": dim})], "align": PP_ALIGN.RIGHT}])
    rect(slide, 0, 2.10, W, 0.16, acc)

    # ── headline ──────────────────────────────────────────────────────────────
    tb(slide, ML, 2.42, CW, 0.42,
       [{"runs": [(c["kicker"], {"sz": 8.0, "b": True, "c": acc})]}])
    tb(slide, ML, 2.76, 17.9, 1.64,
       [{"runs": [(ln, {"sz": 21, "b": True, "c": ink})], "ls": 0.94} for ln in c["headline"]],
       MSO_ANCHOR.BOTTOM)
    tb(slide, ML, 4.52, 16.9, 0.95,
       [{"runs": [(c["sub"], {"sz": 9.6, "c": mid})], "ls": 1.05}])

    # ── hero price ────────────────────────────────────────────────────────────
    tb(slide, ML, 5.34, 7.9, 2.12,
       [{"runs": [(c["big"], {"sz": 45, "b": True, "c": acc})]}], MSO_ANCHOR.MIDDLE)
    rect(slide, 8.62, 5.72, 0.05, 1.36, "C9CDD0")
    tb(slide, 9.00, 5.76, 6.0, 0.55,
       [{"runs": [(c["unit"][0], {"sz": 12.8, "b": True, "c": ink})]}])
    tb(slide, 9.00, 6.28, 6.0, 0.55,
       [{"runs": [(c["unit"][1], {"sz": 12.8, "b": True, "c": mid})]}])

    # ── two columns ───────────────────────────────────────────────────────────
    tb(slide, ML, 7.86, 10.2, 0.42,
       [{"runs": [(c["col1"], {"sz": 9.4, "b": True, "c": ink})]}])
    tb(slide, 11.85, 7.86, 7.85, 0.42,
       [{"runs": [(c["col2"], {"sz": 9.4, "b": True, "c": ink})]}])

    box = slide.shapes.add_textbox(Cm(ML), Cm(8.40), Cm(10.15), Cm(5.90))
    tf = box.text_frame; tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for i, item in enumerate(c["inc"]):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.line_spacing = 1.0
        para.space_after = Pt(6)
        r = para.add_run(); r.text = item
        r.font.name = FONT; r.font.size = Pt(8.9); r.font.color.rgb = rgb(ink)
        bullet(para, "•", acc, 0.40)

    y = 8.40
    for title, expl in c["exc"]:
        rect(slide, 11.85, y + 0.07, 0.13, 0.13, acc)
        tb(slide, 12.22, y, 7.48, 0.36,
           [{"runs": [(title, {"sz": 8.8, "b": True, "c": ink})]}])
        tb(slide, 12.22, y + 0.36, 7.48, 0.46,
           [{"runs": [(expl, {"sz": 8.0, "c": mid})], "ls": 0.98}])
        y += 1.18
    tb(slide, 11.85, y + 0.02, 7.85, 0.78,
       [{"runs": [(c["exc_note"], {"sz": 8.0, "b": True, "c": acc})], "ls": 1.05}])

    # ── example invoice ───────────────────────────────────────────────────────
    y0 = 14.42
    tb(slide, ML, y0, CW, 0.42,
       [{"runs": [(c["ex_title"], {"sz": 9.2, "b": True, "c": ink})]}])
    ry = y0 + 0.52
    rh = 0.68
    for i, (label, amount) in enumerate(c["rows"]):
        if i % 2 == 0:
            rect(slide, ML, ry, CW, rh, LIGHT)
        tb(slide, ML + 0.28, ry, 13.9, rh,
           [{"runs": [(label, {"sz": 8.9, "c": ink})]}], MSO_ANCHOR.MIDDLE)
        tb(slide, 14.30, ry, 5.10, rh,
           [{"runs": [(amount, {"sz": 9.4, "b": True, "c": ink})], "align": PP_ALIGN.RIGHT}],
           MSO_ANCHOR.MIDDLE)
        ry += rh
    th = 0.86
    rect(slide, ML, ry, CW, th, ink)
    tb(slide, ML + 0.28, ry, 12.0, th,
       [{"runs": [(c["total"][0], {"sz": 10.2, "b": True, "c": white})]}], MSO_ANCHOR.MIDDLE)
    tb(slide, 12.60, ry, 6.80, th,
       [{"runs": [(c["total"][1], {"sz": 14.5, "b": True, "c": acc})], "align": PP_ALIGN.RIGHT}],
       MSO_ANCHOR.MIDDLE)
    ry += th

    # ── stats strip ───────────────────────────────────────────────────────────
    sy = ry + 0.30
    rect(slide, ML, sy, CW, 1.16, LIGHT)
    cw = CW / 4.0
    for i, (val, lab) in enumerate(c["stats"]):
        x = ML + i * cw
        tb(slide, x, sy + 0.13, cw, 0.52,
           [{"runs": [(val, {"sz": 15.5, "b": True, "c": acc})], "align": PP_ALIGN.CENTER}])
        tb(slide, x, sy + 0.66, cw, 0.44,
           [{"runs": [(lab.replace("\n", " "), {"sz": 7.2, "c": mid})], "align": PP_ALIGN.CENTER}])

    # ── checklist ─────────────────────────────────────────────────────────────
    ky = sy + 1.50
    tb(slide, ML, ky, CW, 0.42,
       [{"runs": [(c["chk_title"], {"sz": 9.0, "b": True, "c": ink})]}])
    box = slide.shapes.add_textbox(Cm(ML), Cm(ky + 0.50), Cm(CW), Cm(2.90))
    tf = box.text_frame; tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for i, q in enumerate(c["chk"]):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.line_spacing = 1.0
        para.space_after = Pt(4.8)
        rn = para.add_run(); rn.text = f"{i+1}."
        rn.font.name = FONT; rn.font.size = Pt(8.7); rn.font.bold = True
        rn.font.color.rgb = rgb(acc)
        rt = para.add_run(); rt.text = "  " + q
        rt.font.name = FONT; rt.font.size = Pt(8.7); rt.font.color.rgb = rgb(ink)
        hanging(para, 0.42)

    tb(slide, ML, ky + 3.06, CW, 0.55,
       [{"runs": [(c["fine"], {"sz": 7.2, "i": True, "c": mid})], "ls": 1.0}])

    # ── CTA band ──────────────────────────────────────────────────────────────
    cy = 24.55
    rect(slide, 0, cy, W, H - cy - 1.30, ink)
    tb(slide, ML, cy + 0.20, 9.9, 0.62,
       [{"runs": [(c["cta1"], {"sz": 15, "b": True, "c": white})]}])
    tb(slide, ML, cy + 0.82, 9.9, 0.42,
       [{"runs": [(c["cta2"], {"sz": 11.5, "b": True, "c": acc})]}])
    tb(slide, ML, cy + 1.30, 9.9, 0.44,
       [{"runs": [(c["cta3"], {"sz": 8.2, "c": dim})], "ls": 1.0}])
    tb(slide, ML, cy + 1.86, 9.9, 1.15,
       [{"runs": [(c["contact"][0], {"sz": 8.6, "b": True, "c": white})], "sa": 4},
        {"runs": [(c["contact"][1], {"sz": 9.8, "b": True, "c": white})], "sa": 4},
        {"runs": [(c["contact"][2], {"sz": 8.4, "c": dim})]}])
    for (qx, img, cap) in ((13.30, "assets/qr-rkd.png", c["qr1"]),
                           (17.05, "assets/qr-pay.png", c["qr2"])):
        rect(slide, qx, cy + 0.24, 2.40, 2.40, white)
        slide.shapes.add_picture(img, Cm(qx + 0.11), Cm(cy + 0.35), Cm(2.18), Cm(2.18))
        tb(slide, qx - 0.15, cy + 2.74, 2.70, 0.62,
           [{"runs": [(cap, {"sz": 6.6, "c": dim})], "align": PP_ALIGN.CENTER, "ls": 1.0}])

    # ── footer ────────────────────────────────────────────────────────────────
    tb(slide, 0.80, H - 1.06, W - 1.60, 0.40,
       [{"runs": [(c["footer"], {"sz": 6.8, "c": mid})], "align": PP_ALIGN.CENTER}])


def main():
    prs = Presentation()
    prs.slide_width = Cm(W)
    prs.slide_height = Cm(H)
    for content in (MK, AL):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        build(content, slide)
    out = "DARBO_GRUP_flaer_A4_MK-AL.pptx"
    prs.save(out)
    print("saved", out, "slides:", len(prs.slides))


if __name__ == "__main__":
    main()
