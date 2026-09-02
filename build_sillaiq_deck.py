# -*- coding: utf-8 -*-
"""
يبني برزنتيشن PowerPoint تسويقي شامل لتطبيق «صِلَة دوائي» (العراق)
من الاستراتيجية الشاملة. بخطوط عربية RTL وألوان هوية صحية موثوقة.
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from pptx.enum.dml import MSO_FILL

# ---------- الهوية البصرية ----------
PRIMARY   = RGBColor(0x0E, 0x7C, 0x7B)   # تركوازي صحي
DARK      = RGBColor(0x0B, 0x25, 0x45)   # كحلي عميق
ACCENT    = RGBColor(0xE0, 0xA4, 0x58)   # ذهبي دافئ
LIGHT     = RGBColor(0xF4, 0xF7, 0xFA)   # خلفية فاتحة
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
GREY      = RGBColor(0x6B, 0x74, 0x84)
MUTED     = RGBColor(0x50, 0x5A, 0x6A)
TEAL_DARK = RGBColor(0x0A, 0x5C, 0x5B)

FONT = "Segoe UI"
FONT_ARABIC = "Segoe UI"

prs = Presentation()
prs.slide_width  = Inches(13.333)
prs.slide_height = Inches(7.5)
W = prs.slide_width
H = prs.slide_height

SW = prs.slide_width
SH = prs.slide_height


def set_rtl(paragraph):
    """يجعل الفقرة RTL (للعربية)"""
    pPr = paragraph._p.get_or_add_pPr()
    pPr.set('rtl', '1')
    pPr.set('bidi', '1')


def add_text(slide, x, y, w, h, runs, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.TOP,
             rtl=True, line_spacing=1.0, space_after=0):
    """runs: list of paragraphs; each paragraph = dict {text,size,bold,color,font,align,space_before}"""
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    from pptx.util import Pt as Pt_
    # margins to zero
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    first = True
    for para in runs:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = para.get('align', align)
        if para.get('rtl', rtl):
            set_rtl(p)
        p.line_spacing = para.get('line_spacing', line_spacing)
        if para.get('space_after') is not None:
            p.space_after = Pt(para['space_after'])
        if para.get('space_before') is not None:
            p.space_before = Pt(para['space_before'])
        # support 1..n runs per paragraph
        segs = para.get('segments') or [para]
        for seg in segs:
            r = p.add_run()
            r.text = seg.get('text', '')
            f = r.font
            f.size = Pt(seg.get('size', 18))
            f.bold = seg.get('bold', False)
            f.name = seg.get('font', FONT)
            f.color.rgb = seg.get('color', MUTED)
    return tb


def add_rect(slide, x, y, w, h, fill=PRIMARY, line=None, shape=MSO_SHAPE.RECTANGLE, radius=None):
    sp = slide.shapes.add_shape(shape, x, y, w, h)
    sp.fill.solid()
    sp.fill.fore_color.rgb = fill
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = line
        sp.line.width = Pt(1)
    sp.shadow.inherit = False
    if radius is not None and shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        try:
            sp.adjustments[0] = radius
        except Exception:
            pass
    return sp


def add_pill(slide, x, y, w, h, text, fill, tcolor=WHITE, size=13, bold=True):
    sp = add_rect(slide, x, y, w, h, fill=fill, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.5)
    tf = sp.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    set_rtl(p)
    r = p.add_run()
    r.text = text
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.name = FONT
    r.font.color.rgb = tcolor
    return sp


def add_title(slide, title, kicker=None, accent_color=PRIMARY):
    """شريط عنوان موحّد"""
    add_rect(slide, 0, 0, SW, Emu(Inches(0.04).emu * 100 // 100), fill=accent_color)
    add_rect(slide, 0, 0, Inches(1.35), Inches(0.10), fill=accent_color)
    if kicker:
        add_text(slide, Inches(0.6), Inches(0.32), Inches(11.6), Inches(0.5),
                 [{'text': kicker.upper(), 'size': 13, 'bold': True, 'color': accent_color, 'rtl': True}],
                 align=PP_ALIGN.RIGHT)
    add_text(slide, Inches(0.6), Inches(0.62), Inches(11.9), Inches(0.95),
             [{'text': title, 'size': 30, 'bold': True, 'color': DARK}],
             align=PP_ALIGN.RIGHT)
    add_rect(slide, Inches(0.6), Inches(1.7), Inches(1.0), Inches(0.06), fill=accent_color)


def add_footer(slide, page):
    add_text(slide, Inches(0.5), Inches(7.06), Inches(12.3), Inches(0.3),
             [{'text': 'صالَة دوائي — الاستراتيجية التسويقية الشاملة', 'size': 10, 'color': GREY}],
             align=PP_ALIGN.RIGHT)
    add_text(slide, Inches(0.5), Inches(7.06), Inches(12.3), Inches(0.3),
             [{'text': str(page), 'size': 10, 'color': GREY}], align=PP_ALIGN.LEFT, rtl=False)


def slide_bg(slide, color=LIGHT):
    add_rect(slide, 0, 0, SW, SH, fill=color)


# ================== الشريحة 1: الغلاف ==================
s = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(s, DARK)
add_rect(s, 0, 0, SW, Inches(0.35), fill=ACCENT)
# دوائر زخرفية
add_rect(s, SW - Inches(3.2), Inches(1.0), Inches(3.0), Inches(3.0), fill=TEAL_DARK, shape=MSO_SHAPE.OVAL)
add_rect(s, SW - Inches(2.0), Inches(0.0), Inches(1.6), Inches(1.6), fill=PRIMARY, shape=MSO_SHAPE.OVAL)
add_rect(s, -Inches(1.0), Inches(5.2), Inches(2.6), Inches(2.6), fill=TEAL_DARK, shape=MSO_SHAPE.OVAL)

add_text(s, Inches(1.0), Inches(1.35), Inches(9.5), Inches(0.6),
         [{'text': 'صِـلَة دوائي  |  Silla Al-Dawai', 'size': 20, 'bold': True, 'color': ACCENT, 'rtl': True}],
         align=PP_ALIGN.RIGHT)
add_text(s, Inches(1.0), Inches(2.25), Inches(11.0), Inches(2.2),
         [{'text': 'الاستراتيجية التسويقية الشاملة', 'size': 48, 'bold': True, 'color': WHITE, 'rtl': True, 'line_spacing': 1.05},
          {'text': 'منصّة الرعاية الصيدلانية الرقمية في العراق', 'size': 24, 'color': RGBColor(0xC8, 0xE0, 0xDF), 'rtl': True, 'space_before': 10}],
         align=PP_ALIGN.RIGHT)
add_rect(s, Inches(1.0), Inches(4.85), Inches(2.6), Inches(0.10), fill=ACCENT)
add_text(s, Inches(1.0), Inches(5.15), Inches(9.5), Inches(0.6),
         [{'segments': [{'text': 'أونلاين  •  أوفلاين  •  السوق العراقي  •  النمو', 'size': 16, 'bold': True, 'color': WHITE}]}],
         align=PP_ALIGN.RIGHT)
add_text(s, Inches(1.0), Inches(6.35), Inches(11.0), Inches(0.5),
         [{'text': 'دواؤك… يصلك من صيدلية مرخّصة قريبة منك', 'size': 15, 'color': RGBColor(0x9F, 0xB7, 0xC7)}],
         align=PP_ALIGN.RIGHT)


# ================== شريحة 2: خطة العرض (Agenda) ==================
s = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(s)
add_title(s, 'خطة العرض', kicker='Agenda')
items = [
    ('١', 'لماذا صِلَة؟ الرؤية والمشكلة'),
    ('٢', 'السوق العراقي — أرقام تحسم القرار'),
    ('٣', 'المشكلة الكبرى: فوضى الأدوية غير الموثوقة'),
    ('٤', 'تموضع صِلَة وموقع التمايز'),
    ('٥', 'الجماهير المستهدفة'),
    ('٦', 'نموذج النموّ (الـ Flywheel)'),
    ('٧', 'الأونلاين — أفكار إبداعية'),
    ('٨', 'الأوفلاين — أفكار إبداعية'),
    ('٩', 'الميزانية والخارطة والمؤشّرات'),
    ('١٠', 'المخاطر والامتثال — والبدء'),
]
y = Inches(1.95)
for num, txt in items:
    add_rect(s, Inches(0.7), y, Inches(0.55), Inches(0.55), fill=PRIMARY, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.4)
    add_text(s, Inches(0.7), y + Inches(0.02), Inches(0.55), Inches(0.5),
             [{'text': num, 'size': 16, 'bold': True, 'color': WHITE}], align=PP_ALIGN.CENTER)
    add_text(s, Inches(1.5), y + Inches(0.05), Inches(10.6), Inches(0.5),
             [{'text': txt, 'size': 18, 'color': DARK}], align=PP_ALIGN.RIGHT)
    y = y + Inches(0.5)
add_footer(s, 2)


# ================== شريحة 3: الرؤية والمشكلة ==================
s = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(s)
add_title(s, 'لماذا صِلَة؟', kicker='الرؤية')
# بطاقتان
add_rect(s, Inches(0.6), Inches(1.95), Inches(6.0), Inches(4.7), fill=WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.03)
add_rect(s, Inches(0.6), Inches(1.95), Inches(6.0), Inches(0.14), fill=PRIMARY)
add_text(s, Inches(0.9), Inches(2.25), Inches(5.4), Inches(0.5),
         [{'text': 'الرؤية', 'size': 22, 'bold': True, 'color': PRIMARY, 'rtl': True}], align=PP_ALIGN.RIGHT)
add_text(s, Inches(0.9), Inches(2.85), Inches(5.4), Inches(3.5),
         [{'text': 'نبني البنية التحتية الرقمية التي تجعل الدواء يصل إلى كل بيت في العراق — بثقة، وشفافية، وكرامة.', 'size': 16, 'color': MUTED, 'rtl': True, 'line_spacing': 1.3}],
         align=PP_ALIGN.RIGHT)

add_rect(s, Inches(6.9), Inches(1.95), Inches(6.0), Inches(4.7), fill=WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.03)
add_rect(s, Inches(6.9), Inches(1.95), Inches(6.0), Inches(0.14), fill=ACCENT)
add_text(s, Inches(7.2), Inches(2.25), Inches(5.4), Inches(0.5),
         [{'text': 'المشكلة', 'size': 22, 'bold': True, 'color': ACCENT, 'rtl': True}], align=PP_ALIGN.RIGHT)
add_text(s, Inches(7.2), Inches(2.85), Inches(5.4), Inches(3.5),
         [{'segments': [{'text': 'الدواء في العراق يُباع اليوم عبر صفحات تواصل غير موثوقة، بلا ترخيص ولا رقابة. ', 'size': 16, 'color': MUTED, 'rtl': True},
                        {'text': 'صحة المريض في خطر.', 'size': 16, 'bold': True, 'color': DARK, 'rtl': True}]}],
         align=PP_ALIGN.RIGHT)
add_footer(s, 3)


# ================== شريحة 4: السوق العراقي ==================
s = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(s)
add_title(s, 'السوق العراقي — أرقام تحسم القرار', kicker='2025')
rows = [
    ('السكان', '46.5 مليون', '72% في المدن'),
    ('مستخدمو الإنترنت', '38.0 مليون', '81.7% من السكان'),
    ('هويّات التواصل', '34.3 مليون', '73.8% من السكان'),
    ('تيك توك', '~90% من مستخدمي الإنترنت', 'الأوسع انتشاراً'),
    ('فيسبوك', '~20.1 مليون', '75% من البالغين'),
    ('إنستغرام', '~19.0 مليون', '69% من البالغين'),
    ('العمر الوسيط', '20.8 سنة', 'سوق شابّ'),
]
y = Inches(1.95)
add_rect(s, Inches(0.6), y, Inches(4.2), Inches(0.6), fill=DARK)
add_rect(s, Inches(4.9), y, Inches(4.2), Inches(0.6), fill=DARK)
add_rect(s, Inches(9.2), y, Inches(3.6), Inches(0.6), fill=DARK)
for col, txt in zip([('المؤشّر', Inches(1.0)), ('القيمة', Inches(5.2)), ('الدلالة', Inches(9.5))], [('المؤشّر', 0), ('القيمة', 0), ('الدلالة', 0)]):
    add_text(s, col[1], y + Inches(0.05), Inches(3.6), Inches(0.5),
             [{'text': col[0], 'size': 15, 'bold': True, 'color': WHITE}], align=PP_ALIGN.CENTER if col[0] != 'المؤشّر' else PP_ALIGN.RIGHT)
y = y + Inches(0.6)
for i, (a, b, c) in enumerate(rows):
    bg = WHITE if i % 2 == 0 else LIGHT
    add_rect(s, Inches(0.6), y, Inches(4.2), Inches(0.58), fill=bg)
    add_rect(s, Inches(4.9), y, Inches(4.2), Inches(0.58), fill=bg)
    add_rect(s, Inches(9.2), y, Inches(3.6), Inches(0.58), fill=bg)
    add_text(s, Inches(0.9), y + Inches(0.04), Inches(3.6), Inches(0.5),
             [{'text': a, 'size': 15, 'bold': True, 'color': DARK}], align=PP_ALIGN.RIGHT)
    add_text(s, Inches(4.9), y + Inches(0.04), Inches(4.2), Inches(0.5),
             [{'text': b, 'size': 15, 'bold': True, 'color': PRIMARY}], align=PP_ALIGN.CENTER)
    add_text(s, Inches(9.2), y + Inches(0.04), Inches(3.5), Inches(0.5),
             [{'text': c, 'size': 13, 'color': MUTED}], align=PP_ALIGN.RIGHT)
    y = y + Inches(0.58)
add_footer(s, 4)


# ================== شريحة 5: الفوضى غير الموثوقة ==================
s = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(s)
add_title(s, 'المشكلة الكبرى: فوضى الأدوية غير الموثوقة', kicker='فرصة خفيّة', accent_color=ACCENT)
add_text(s, Inches(0.6), Inches(1.95), Inches(12.1), Inches(0.5),
         [{'text': 'نقابة صيادلة العراق تُحذّر رسمياً من شراء الأدوية عبر صفحات التواصل الاجتماعي', 'size': 18, 'bold': True, 'color': DARK, 'rtl': True}],
         align=PP_ALIGN.RIGHT)
add_rect(s, Inches(0.6), Inches(2.6), Inches(12.1), Inches(0.06), fill=ACCENT)
pts = [
    'الأدوية المسجّلة لا يجوز بيعها إلكترونياً خارج صيدلية مرخّصة، وتحتاج نصيحة صيدلي وتوضيح محاذير الاستعمال.',
    'تنتشر حسابات وهمية على فيسبوك تنتحل صفة أطباء عراقيين، وتبيع علاجات وهمية ومنتهية الصلاحية وخلطات عشبية.',
    'المصدر الشرعي الوحيد للدواء بحكم التعريف هو الصيدلية المرخّصة الخاضعة للرقابة.',
]
y = Inches(2.9)
for p in pts:
    add_rect(s, Inches(0.75), y, Inches(0.25), Inches(0.25), fill=ACCENT, shape=MSO_SHAPE.OVAL)
    add_text(s, Inches(1.2), y - Inches(0.02), Inches(11.3), Inches(0.9),
             [{'text': p, 'size': 15, 'color': MUTED, 'rtl': True, 'line_spacing': 1.25}], align=PP_ALIGN.RIGHT)
    y = y + Inches(1.1)
add_rect(s, Inches(0.6), Inches(6.05), Inches(12.1), Inches(0.9), fill=PRIMARY, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.15)
add_text(s, Inches(0.9), Inches(6.25), Inches(11.6), Inches(0.6),
         [{'segments': [{'text': '⬅️ ', 'size': 17, 'color': WHITE, 'rtl': True},
                        {'text': 'موقعُ التمايز الذي لا يملكه منافس عراقي:', 'size': 17, 'color': WHITE, 'rtl': True},
                        {'text': ' «صِلَة دوائي = القناة القانونية المعتمدة»', 'size': 17, 'bold': True, 'color': ACCENT, 'rtl': True}]}],
         align=PP_ALIGN.RIGHT)
add_footer(s, 5)


# ================== شريحة 6: التموضع ==================
s = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(s)
add_title(s, 'تموضع صِلَة وموقع التمايز', kicker='Brand')
# الركائز الثلاث
y = Inches(2.0)
pillars = [
    ('الثقة المرخّصة', 'كل صيدلية وسائق موثّق ومُعتمَد'),
    ('التوفّر الفعلي', 'لا وعود — توفّرٌ حقيقي يعرض أمامك'),
    ('الدفع عند الاستلام', 'معاملتك بيدك حتّى الاستلام'),
]
for i, (t, d) in enumerate(pillars):
    x = Inches(0.6) + i * Inches(4.2)
    add_rect(s, x, y, Inches(3.9), Inches(3.0), fill=WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.04)
    add_rect(s, x, y, Inches(3.9), Inches(1.0), fill=PRIMARY)
    add_text(s, x + Inches(0.3), y + Inches(0.25), Inches(3.3), Inches(0.6),
             [{'text': t, 'size': 20, 'bold': True, 'color': WHITE}], align=PP_ALIGN.CENTER)
    add_text(s, x + Inches(0.3), y + Inches(1.25), Inches(3.3), Inches(1.5),
             [{'text': d, 'size': 15, 'color': MUTED, 'rtl': True, 'line_spacing': 1.25}], align=PP_ALIGN.CENTER)
add_text(s, Inches(0.6), Inches(5.35), Inches(12.1), Inches(1.4),
         [{'segments': [{'text': '«صِلَة ليست تطبيق توصيل، بل ', 'size': 18, 'color': DARK, 'rtl': True},
                        {'text': 'شبكة صحية مُدارة ومرخّصة.', 'size': 18, 'bold': True, 'color': PRIMARY, 'rtl': True}]}],
         align=PP_ALIGN.CENTER)
add_footer(s, 6)


# ================== شريحة 7: الجماهير ==================
s = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(s)
add_title(s, 'الجماهير المستهدفة', kicker='Audience')
cols = [
    ('المرضى (B2C)', PRIMARY, ['مريض مزمن/مسنّ/مقيّد', 'أمّ/ربّ أسرة', 'موظف/عامل', 'شاب رقمي']),
    ('الصيدليات (B2B)', ACCENT, ['صيدلية مستقلة صغيرة', 'صيدلية متوسّطة/سلسلة']),
    ('السائقون (Gig)', TEAL_DARK, ['طلبة', 'مستقلّون', 'أصحاب سيارات']),
]
for i, (title, color, items) in enumerate(cols):
    x = Inches(0.6) + i * Inches(4.2)
    add_rect(s, x, Inches(2.0), Inches(3.9), Inches(4.5), fill=WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.04)
    add_rect(s, x, Inches(2.0), Inches(3.9), Inches(0.85), fill=color)
    add_text(s, x + Inches(0.2), Inches(2.2), Inches(3.5), Inches(0.5),
             [{'text': title, 'size': 19, 'bold': True, 'color': WHITE}], align=PP_ALIGN.CENTER)
    yy = Inches(3.1)
    for it in items:
        add_rect(s, x + Inches(0.3), yy + Inches(0.1), Inches(0.2), Inches(0.2), fill=color, shape=MSO_SHAPE.OVAL)
        add_text(s, x + Inches(0.7), yy, Inches(2.9), Inches(0.5),
                 [{'text': it, 'size': 15, 'color': MUTED, 'rtl': True}], align=PP_ALIGN.RIGHT)
        yy = yy + Inches(0.6)
add_footer(s, 7)


# ================== شريحة 8: نموذج النمو ==================
s = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(s)
add_title(s, 'نموذج النموّ — الـ Flywheel', kicker='Growth')
boxes = [
    ('توفّر عالٍ', Inches(1.2), Inches(2.2)), ('تجربة ممتازة', Inches(1.2), Inches(3.6)),
    ('إحالة وشفاهة', Inches(1.2), Inches(5.0)), ('مزيد زبائن', Inches(5.2), Inches(2.2)),
    ('طلبات أكثر', Inches(5.2), Inches(3.6)), ('عرض أقوى', Inches(5.2), Inches(5.0)),
]
for t, x, y in boxes:
    add_rect(s, x, y, Inches(2.4), Inches(1.0), fill=PRIMARY, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.4)
    add_text(s, x, y + Inches(0.2), Inches(2.4), Inches(0.6),
             [{'text': t, 'size': 16, 'bold': True, 'color': WHITE}], align=PP_ALIGN.CENTER)
add_text(s, Inches(0.6), Inches(2.0), Inches(12.1), Inches(0.5),
         [{'text': 'المبدأ: الكثافة قبل الاتّساع (Density-First) — ابتدِئ بنطاق واحد مكثّف قبل التوسّع.', 'size': 15, 'bold': True, 'color': TEAL_DARK, 'rtl': True}],
         align=PP_ALIGN.RIGHT)
add_rect(s, Inches(8.6), Inches(2.3), Inches(4.0), Inches(4.0), fill=LIGHT, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.05)
add_text(s, Inches(8.8), Inches(2.55), Inches(3.6), Inches(3.4),
         [{'text': 'لماذا يعمل؟', 'size': 18, 'bold': True, 'color': PRIMARY, 'rtl': True, 'space_after': 8},
          {'text': 'كل صيدلية وسائق جديد في نفس النطاق يزيد التوفّر ويقلّل تكلفة الجذب ويصعّد الثقة.', 'size': 14, 'color': MUTED, 'rtl': True, 'line_spacing': 1.3}],
         align=PP_ALIGN.RIGHT)
add_footer(s, 8)


# ================== شريحة 9: الأونلاين (نظرة عامة) ==================
s = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(s)
add_title(s, 'الأونلاين — منظومة القنوات الإبداعية', kicker='Online', accent_color=PRIMARY)
channels = [
    ('تيك توك / ريلز', 'الأوسع انتشاراً — فيديوهات قصيرة'),
    ('فيسبوك + ماسنجر', 'الثقة والجمهور الأوسع'),
    ('إنستغرام', 'بصرية / حياة / جمال'),
    ('واتساب / تيليجرام', 'بيع وخدمة وتأكيد طلبات'),
    ('ASO + SEO', 'نموّ غير مكلّف من المتجر والبحث'),
    ('الإحالة + الولاء', 'أعلى تحويل في القطاع الصحي'),
]
for i, (t, d) in enumerate(channels):
    row = i // 3
    col = i % 3
    x = Inches(0.6) + col * Inches(4.2)
    y = Inches(2.15) + row * Inches(2.3)
    add_rect(s, x, y, Inches(3.9), Inches(2.0), fill=WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.06)
    add_rect(s, x, y, Inches(0.16), Inches(2.0), fill=PRIMARY)
    add_text(s, x + Inches(0.4), y + Inches(0.3), Inches(3.3), Inches(0.6),
             [{'text': t, 'size': 18, 'bold': True, 'color': DARK}], align=PP_ALIGN.RIGHT)
    add_text(s, x + Inches(0.4), y + Inches(1.0), Inches(3.3), Inches(0.9),
             [{'text': d, 'size': 13, 'color': MUTED, 'rtl': True, 'line_spacing': 1.2}], align=PP_ALIGN.RIGHT)
add_footer(s, 9)


# ================== شريحة 10: حملة «صيدليتك لا صفحة مجهولة» ==================
s = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(s)
add_title(s, 'الحملة الكبرى: «صيدليتُك لا صفحةٌ مجهولة»', kicker='Creative', accent_color=ACCENT)
add_rect(s, Inches(0.6), Inches(2.0), Inches(12.1), Inches(0.85), fill=ACCENT, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.2)
add_text(s, Inches(0.9), Inches(2.2), Inches(11.6), Inches(0.5),
         [{'text': 'حملة توعوية وطنية تتصدّى لظاهرة الأدوية المزوّرة — تخدم المريض وتبني العلامة', 'size': 17, 'bold': True, 'color': DARK, 'rtl': True}],
         align=PP_ALIGN.RIGHT)
items = [
    'سلسلة «كيف تتحقق أن الدواء أصلي؟» — تعليم + بناء ثقة',
    'فيديوهات «الفرق بين صيدلية مرخّصة وصفحة وهمية» بأسلوب قصصي',
    'محتوى «جناح الثقة» — لقطات حقيقية من رحلة توثيق الصيدليات',
    'هاشتاغ حامل للحملة (بمراجعة ناطق وقانوني)',
]
y = Inches(3.2)
for it in items:
    add_rect(s, Inches(0.75), y + Inches(0.1), Inches(0.24), Inches(0.24), fill=ACCENT, shape=MSO_SHAPE.OVAL)
    add_text(s, Inches(1.2), y, Inches(11.3), Inches(0.6),
             [{'text': it, 'size': 16, 'color': MUTED, 'rtl': True}], align=PP_ALIGN.RIGHT)
    y = y + Inches(0.85)
add_footer(s, 10)


# ================== شريحة 11: الأونلاين — تيك توك ==================
s = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(s)
add_title(s, 'الأونلاين: تيك توك / ريلز — الأوسع انتشاراً', kicker='~90% من مستخدمي الإنترنت')
ideas = [
    '«دواءك موجود؟» — مقاطع قصيرة تبحث عن دواء نادر وتُظهر توفّره الفعلي.',
    '«رحلة الطلب في 10 ثوانٍ» — من تصوير الوصفة إلى استلام الطرد.',
    'سلسلة «ضحّينا على الصفحات المزوّرة» — بفكاهة حذرة تشرح الفخاخ.',
    'قصص حقيقية من المستخدمين (UGC) بدعم تعاوني.',
    'معلومة صيدلانية قصيرة — بمراجعة صيدلانية صارمة (منع الجرعات/الاستبدال).',
]
add_rect(s, Inches(0.6), Inches(2.0), Inches(12.1), Inches(0.8), fill=PRIMARY, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.2)
add_text(s, Inches(0.9), Inches(2.2), Inches(11.6), Inches(0.5),
         [{'text': 'المحتوى قصير، حقيقي، وموثوق — لا ادّعاءات صحية خارج الرقابة', 'size': 16, 'bold': True, 'color': WHITE, 'rtl': True}],
         align=PP_ALIGN.RIGHT)
y = Inches(3.2)
for it in ideas:
    add_rect(s, Inches(0.75), y + Inches(0.1), Inches(0.24), Inches(0.24), fill=PRIMARY, shape=MSO_SHAPE.OVAL)
    add_text(s, Inches(1.2), y - Inches(0.02), Inches(11.3), Inches(0.9),
             [{'text': it, 'size': 15, 'color': MUTED, 'rtl': True, 'line_spacing': 1.2}], align=PP_ALIGN.RIGHT)
    y = y + Inches(0.95)
add_footer(s, 11)


# ================== شريحة 12: الأونلاين — فيسبوك/ماسنجر/واتساب ==================
s = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(s)
add_title(s, 'الأونلاين: فيسبوك + ماسنجر + واتساب', kicker='الثقة والمحادثة')
left = [
    'الاستفادة من تحذير النقابة: «الدواء يُعرض في الصيدلية، لا في صفحة».',
    'محتوى طمأنة بجهة مرخّصة معزّز بدليل التوثيق.',
    'مجموعات الأحياء/الأمهات — حملات إرشاد ووعي صحي.',
]
right = [
    'بوت ماسنجر/واتساب: تأكيد الطلب، تتبّع الحالة، دعم مباشر.',
    'جروب واتساب/تيليجرام رسمي — خدمة وطلبات وتعويض لمستخدمي iOS.',
    'بيئة محادثة مألوفة تقلّل الاحتكاك وترفع التحوّل.',
]
for (x, title, py, items, color) in [(Inches(0.6), 'المحتوى والثقة', Inches(2.15), left, PRIMARY),
                                      (Inches(7.0), 'المحادثة والأتمتة', Inches(2.15), right, ACCENT)]:
    add_rect(s, x, Inches(1.95), Inches(5.8), Inches(4.7), fill=WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.03)
    add_rect(s, x, Inches(1.95), Inches(5.8), Inches(0.65), fill=color)
    add_text(s, x + Inches(0.3), Inches(2.1), Inches(5.2), Inches(0.5),
             [{'text': title, 'size': 18, 'bold': True, 'color': WHITE}], align=PP_ALIGN.RIGHT)
    yy = Inches(2.9)
    for it in items:
        add_rect(s, x + Inches(0.35), yy + Inches(0.1), Inches(0.22), Inches(0.22), fill=color, shape=MSO_SHAPE.OVAL)
        add_text(s, x + Inches(0.75), yy - Inches(0.02), Inches(4.7), Inches(0.9),
                 [{'text': it, 'size': 14, 'color': MUTED, 'rtl': True, 'line_spacing': 1.2}], align=PP_ALIGN.RIGHT)
        yy = yy + Inches(1.2)
add_footer(s, 12)


# ================== شريحة 13: الأونلاين — إنستغرام/ASO/الإحالة ==================
s = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(s)
add_title(s, 'الأونلاين: إنستغرام + ASO + الإحالة', kicker='Growth')
cols = [
    ('إنستغرام', PRIMARY, ['ستوريات «ماذا في طلبك؟» — محتوى حي يومي', 'هايلايتس «كيف تعمل صلة» + «لماذا نحن آمنون»', 'تعاون مؤثّرون محليون مع إفصاح ومراجعة']),
    ('ASO + SEO', ACCENT, ['كلمات: «صيدلية أونلاين»، «توصيل أدوية بغداد»', 'أيقونة + لقطات عربية دعائية', 'تقييمات حقيقية + الردّ على كل تقييم']),
    ('إحالة + ولاء', TEAL_DARK, ['كود إحالة عائلي (الدواء شأن عائلي)', 'برنامج ولاء بنقاط لكل طلب', 'خصم محدود لأوّل طلب / شحن مؤقّت']),
]
for i, (t, color, items) in enumerate(cols):
    x = Inches(0.6) + i * Inches(4.2)
    add_rect(s, x, Inches(2.0), Inches(3.9), Inches(4.6), fill=WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.04)
    add_rect(s, x, Inches(2.0), Inches(3.9), Inches(0.8), fill=color)
    add_text(s, x + Inches(0.2), Inches(2.18), Inches(3.5), Inches(0.5),
             [{'text': t, 'size': 19, 'bold': True, 'color': WHITE}], align=PP_ALIGN.CENTER)
    yy = Inches(3.0)
    for it in items:
        add_rect(s, x + Inches(0.3), yy + Inches(0.1), Inches(0.2), Inches(0.2), fill=color, shape=MSO_SHAPE.OVAL)
        add_text(s, x + Inches(0.7), yy - Inches(0.05), Inches(2.9), Inches(1.2),
                 [{'text': it, 'size': 13, 'color': MUTED, 'rtl': True, 'line_spacing': 1.15}], align=PP_ALIGN.RIGHT)
        yy = yy + Inches(1.15)
add_footer(s, 13)


# ================== شريحة 14: الأوفلاين — نظرة عامة ==================
s = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(s)
add_title(s, 'الأوفلاين — حضورٌ ملموس في الشارع والصعيد', kicker='Offline')
items = [
    ('الشراكات الصحية', 'عيادات، مستشفيات، أطباء، مختبرات'),
    ('نقاط الحضور المادي', 'ملصقات، أكياس تسليم، بطاقة QR'),
    ('السائقون كوسيلة إعلانية', 'أزياء موحّدة + لافتات متحرّكة'),
    ('التوعية المجتمعية', 'حملات بأضرار الأدوية المزوّرة'),
    ('المناسبات والأجندة', 'مواسم الذروة والأعياد'),
]
for i, (t, d) in enumerate(items):
    row = i // 3
    col = i % 3
    x = Inches(0.6) + col * Inches(4.2)
    y = Inches(2.15) + row * Inches(2.3)
    add_rect(s, x, y, Inches(3.9), Inches(2.0), fill=WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.06)
    add_rect(s, x, y, Inches(0.16), Inches(2.0), fill=TEAL_DARK)
    add_text(s, x + Inches(0.4), y + Inches(0.3), Inches(3.3), Inches(0.6),
             [{'text': t, 'size': 17, 'bold': True, 'color': DARK}], align=PP_ALIGN.RIGHT)
    add_text(s, x + Inches(0.4), y + Inches(1.0), Inches(3.3), Inches(0.9),
             [{'text': d, 'size': 13, 'color': MUTED, 'rtl': True, 'line_spacing': 1.2}], align=PP_ALIGN.RIGHT)
add_footer(s, 14)


# ================== شريحة 15: الأوفلاين — الشراكات الصحية ==================
s = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(s)
add_title(s, 'الأوفلاين: الشراكات الصحية (أعلى مردود)', kicker='Partnerships')
add_rect(s, Inches(0.6), Inches(2.0), Inches(12.1), Inches(1.0), fill=PRIMARY, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.1)
add_text(s, Inches(0.9), Inches(2.2), Inches(11.6), Inches(0.6),
         [{'segments': [{'text': 'ابتكار: ', 'size': 17, 'bold': True, 'color': ACCENT, 'rtl': True},
                        {'text': 'بوضع رمز QR في استقبال العيادة والمستشفى: «صوّر وصفتك، وأدويةُ توصلك إلى بيتك.»', 'size': 17, 'bold': True, 'color': WHITE, 'rtl': True}]}],
         align=PP_ALIGN.RIGHT)
pts = [
    'العيادات والمستشفيات والأطباء — QR عند الاستقبال وصالة الانتظار.',
    'المختبرات وأقسام التصوير — «استلم نتيجتك ووصفتك من مكانك».',
    'جمعيات خيرية — حملة «صلة خير» توصل الدواء للمحتاجين (قصة إنسانية).',
    'المراعاة الحاكمة: سرّية بيانات المريض واتفاقيات الوصول.',
]
y = Inches(3.4)
for p in pts:
    add_rect(s, Inches(0.75), y + Inches(0.08), Inches(0.24), Inches(0.24), fill=PRIMARY, shape=MSO_SHAPE.OVAL)
    add_text(s, Inches(1.2), y, Inches(11.3), Inches(0.7),
             [{'text': p, 'size': 16, 'color': MUTED, 'rtl': True, 'line_spacing': 1.2}], align=PP_ALIGN.RIGHT)
    y = y + Inches(0.85)
add_footer(s, 15)


# ================== شريحة 16: الأوفلاين — السائق كوسيلة إعلانية ==================
s = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(s)
add_title(s, 'الأوفلاين: السائقون كوسيلة إعلانية متحرّكة', kicker='Guerrilla', accent_color=ACCENT)
add_rect(s, Inches(0.6), Inches(2.0), Inches(12.1), Inches(1.1), fill=ACCENT, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.1)
add_text(s, Inches(0.9), Inches(2.2), Inches(11.6), Inches(0.8),
         [{'text': 'السائق يتنقّل في كل نطاق يومياً — إعلانٌ مجانيٌّ متواصل يصل إلى كل حيّ', 'size': 18, 'bold': True, 'color': DARK, 'rtl': True}],
         align=PP_ALIGN.RIGHT)
pts = [
    'أزياء موحّدة + لافتات على السيارات/الدراجات تحمل «صِلة دوائي — دواؤك يصلك».',
    'أكياس تسليم موحّدة تحمل الرمز وتمرّ عبر الأحياء يداً بيد.',
    'كود دعوة للسائق + برنامج «سائق يجلب سائقاً».',
    'بطاقة ترحيب داخل كل طلب مع رمز QR وإمكانية مشاركة.',
]
y = Inches(3.5)
for p in pts:
    add_rect(s, Inches(0.75), y + Inches(0.08), Inches(0.24), Inches(0.24), fill=ACCENT, shape=MSO_SHAPE.OVAL)
    add_text(s, Inches(1.2), y, Inches(11.3), Inches(0.7),
             [{'text': p, 'size': 16, 'color': MUTED, 'rtl': True, 'line_spacing': 1.2}], align=PP_ALIGN.RIGHT)
    y = y + Inches(0.85)
add_footer(s, 16)


# ================== شريحة 17: الأوفلاين — توعية + مناسبات ==================
s = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(s)
add_title(s, 'الأوفلاين: التوعية المجتمعية والمناسبات', kicker='Brand purpose')
left = [
    'حملات بأضرار الأدوية المزوّرة في المدارس والجامعات والمراكز.',
    'بشراكة مع نقابة الصيادلة أو مؤسسات صحية — تُبنى الثقة وعلامة «المسؤول».',
    'أيام فحص/تثقيف صحي برعاية «صلة» في سياق صحي موثوق.',
]
right = [
    'مواسم ذروة الأمراض (تغيّر الطقس/الإنفلونزا): «تجهّز — دواؤك يصلك بسرعة».',
    'مناسبات دينية وأعياد: حملات رحيمة «لا تترك دواءك ينتظر».',
    'إطلاق إعلامي بتغطية صحافية محلية (PR) + دعوة مؤثّرين لتجربة طلب حقيقي.',
]
for (x, title, items, color) in [(Inches(0.6), 'التوعية والمجتمع', left, TEAL_DARK),
                                 (Inches(7.0), 'التوقيت والأجندة', right, PRIMARY)]:
    add_rect(s, x, Inches(1.95), Inches(5.8), Inches(4.8), fill=WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.03)
    add_rect(s, x, Inches(1.95), Inches(5.8), Inches(0.65), fill=color)
    add_text(s, x + Inches(0.3), Inches(2.1), Inches(5.2), Inches(0.5),
             [{'text': title, 'size': 18, 'bold': True, 'color': WHITE}], align=PP_ALIGN.RIGHT)
    yy = Inches(2.95)
    for it in items:
        add_rect(s, x + Inches(0.35), yy + Inches(0.1), Inches(0.22), Inches(0.22), fill=color, shape=MSO_SHAPE.OVAL)
        add_text(s, x + Inches(0.75), yy - Inches(0.02), Inches(4.6), Inches(1.0),
                 [{'text': it, 'size': 14, 'color': MUTED, 'rtl': True, 'line_spacing': 1.2}], align=PP_ALIGN.RIGHT)
        yy = yy + Inches(1.25)
add_footer(s, 17)


# ================== شريحة 18: الميزانية ==================
s = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(s)
add_title(s, 'الميزانية — توزيع الإنفاق', kicker='Budget')
data = [('محتوى + إبداع', 20), ('إعلانات أداء (ميتا/تيك توك/إنستغرام)', 35), ('تحفيز (خصم/إحالة/ولاء)', 15),
        ('فريق B2B + عرض', 15), ('توظيف سائقين', 10), ('أدوات/قياس/تسويق', 5)]
y = Inches(2.0)
for i, (lab, pct) in enumerate(data):
    bg = WHITE if i % 2 == 0 else LIGHT
    add_rect(s, Inches(0.6), y, Inches(6.3), Inches(0.65), fill=bg)
    add_rect(s, Inches(7.0), y, Inches(3.6), Inches(0.65), fill=bg)
    add_text(s, Inches(0.9), y + Inches(0.08), Inches(5.8), Inches(0.5),
             [{'text': lab, 'size': 15, 'color': DARK}], align=PP_ALIGN.RIGHT)
    add_text(s, Inches(7.0), y + Inches(0.08), Inches(3.6), Inches(0.5),
             [{'text': f'{pct}%', 'size': 15, 'bold': True, 'color': PRIMARY}], align=PP_ALIGN.CENTER)
    y = y + Inches(0.65)
add_rect(s, Inches(0.6), Inches(6.0), Inches(12.1), Inches(0.85), fill=PRIMARY, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.2)
add_text(s, Inches(0.9), Inches(6.18), Inches(11.6), Inches(0.6),
         [{'segments': [{'text': 'قاعدة: ', 'size': 15, 'bold': True, 'color': ACCENT, 'rtl': True},
                        {'text': 'لا تتجاوز الإنفاق المدفوع قبل أن يتحقّق CAC ≤ LTV و نسبة التوفّر ≥ الهدف.', 'size': 15, 'color': WHITE, 'rtl': True}]}],
         align=PP_ALIGN.RIGHT)
add_footer(s, 18)


# ================== شريحة 19: الخارطة 12 شهراً ==================
s = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(s)
add_title(s, 'خارطة الطريق — 12 شهراً', kicker='Roadmap')
stages = [
    ('الإثبات', '0–3 أشهر', 'سيولة نطاق واحد، قياس، محتوى ثقة، إحالة', PRIMARY),
    ('النمو', '3–9 أشهر', 'توسّع جغرافي، تصعيد تيك توك/إنستغرام، ASO، iOS، مضاعفة العرض', ACCENT),
    ('التوسّع', '9–18 شهراً', 'مدن جديدة، ولاء، شراكات عيادات، ربحية وحدة', TEAL_DARK),
]
x = Inches(0.6)
for t, dur, desc, color in stages:
    add_rect(s, x, Inches(2.2), Inches(3.9), Inches(4.2), fill=WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.05)
    add_rect(s, x, Inches(2.2), Inches(3.9), Inches(1.0), fill=color)
    add_text(s, x + Inches(0.2), Inches(2.4), Inches(3.5), Inches(0.6),
             [{'text': t, 'size': 21, 'bold': True, 'color': WHITE}], align=PP_ALIGN.CENTER)
    add_text(s, x + Inches(0.2), Inches(3.25), Inches(3.5), Inches(0.7),
             [{'text': dur, 'size': 13, 'bold': True, 'color': GREY}], align=PP_ALIGN.CENTER)
    add_text(s, x + Inches(0.3), Inches(3.9), Inches(3.3), Inches(2.3),
             [{'text': desc, 'size': 15, 'color': MUTED, 'rtl': True, 'line_spacing': 1.3}], align=PP_ALIGN.RIGHT)
    x = x + Inches(4.2)
add_footer(s, 19)


# ================== شريحة 20: المؤشّرات ==================
s = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(s)
add_title(s, 'مؤشّرات الأداء (KPIs)', kicker='Measurement')
add_rect(s, Inches(0.6), Inches(2.0), Inches(12.1), Inches(1.1), fill=DARK, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.12)
add_text(s, Inches(0.9), Inches(2.2), Inches(11.6), Inches(0.7),
         [{'segments': [{'text': 'المؤشّر الأساسيّ: ', 'size': 17, 'bold': True, 'color': ACCENT, 'rtl': True},
                        {'text': 'عدد الطلبات المكتملة المدفوعة / أسبوع (North-Star)', 'size': 17, 'bold': True, 'color': WHITE, 'rtl': True}]}],
         align=PP_ALIGN.RIGHT)
cats = [
    ('الاستحواذ', 'تنزيل، تثبيت، CAC'),
    ('رحلة الطلب', 'بحث/رفع، عروض، تأكيد، تسليم'),
    ('جودة السوق', 'نسبة عروض، نسبة ردّ الصيدلية، تسليم ناجح'),
    ('الاحتفاظ', 'تكرار الطلب 30 يوماً'),
    ('الاقتصاد', 'AOV، تحويل السلة، LTV/CAC'),
    ('العرض', 'صيدليات نشطة، سائقون، زمن التوصيل'),
]
for i, (t, d) in enumerate(cats):
    row = i // 3
    col = i % 3
    x = Inches(0.6) + col * Inches(4.2)
    y = Inches(3.5) + row * Inches(1.7)
    add_rect(s, x, y, Inches(3.9), Inches(1.5), fill=WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.08)
    add_text(s, x + Inches(0.3), y + Inches(0.2), Inches(3.3), Inches(0.5),
             [{'text': t, 'size': 16, 'bold': True, 'color': PRIMARY}], align=PP_ALIGN.RIGHT)
    add_text(s, x + Inches(0.3), y + Inches(0.75), Inches(3.3), Inches(0.6),
             [{'text': d, 'size': 12, 'color': MUTED, 'rtl': True}], align=PP_ALIGN.RIGHT)
add_footer(s, 20)


# ================== شريحة 21: المخاطر والامتثال ==================
s = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(s)
add_title(s, 'المخاطر والامتثال — القطاع الصحي', kicker='Governance', accent_color=ACCENT)
risks = [
    ('قانوني/تنظيمي (أدوية)', 'استشارة قانونية وصيدلانية؛ توثيق؛ منع ادّعاءات'),
    ('الادّعاءات الصحية/النصائح', 'مراجعة قبل النشر؛ منع الجرعات والاستبدال'),
    ('الخصوصية والوصفات', 'تشفير؛ تقليل البيانات؛ سياسة؛ موافقة'),
    ('جودة التسليم / تخلّف COD', 'رمز استلام؛ توثيق؛ مكافأة سائق'),
    ('التبنّي الرقمي المنخفض', 'تعليم وثقة وإتاحة الطلب عبر واتساب'),
    ('الاعتماد على الإعلانات', 'إحالة + محتوى + ASO + شفافية'),
]
for i, (t, d) in enumerate(risks):
    row = i // 3
    col = i % 3
    x = Inches(0.6) + col * Inches(4.2)
    y = Inches(2.1) + row * Inches(2.3)
    add_rect(s, x, y, Inches(3.9), Inches(2.1), fill=WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.06)
    add_rect(s, x, y, Inches(0.16), Inches(2.1), fill=ACCENT)
    add_text(s, x + Inches(0.4), y + Inches(0.3), Inches(3.3), Inches(0.6),
             [{'text': t, 'size': 15, 'bold': True, 'color': DARK}], align=PP_ALIGN.RIGHT)
    add_text(s, x + Inches(0.4), y + Inches(1.0), Inches(3.3), Inches(0.9),
             [{'text': d, 'size': 12, 'color': MUTED, 'rtl': True, 'line_spacing': 1.2}], align=PP_ALIGN.RIGHT)
add_footer(s, 21)


# ================== شريحة 22: البدء ==================
s = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(s, DARK)
add_rect(s, 0, 0, SW, Inches(0.35), fill=ACCENT)
add_text(s, Inches(0.6), Inches(0.9), Inches(12.1), Inches(0.8),
         [{'text': 'نبدأ من هنا — أول 30 يوماً', 'size': 34, 'bold': True, 'color': WHITE}], align=PP_ALIGN.RIGHT)
steps = [
    ('١', 'القياس', 'تتبّع التحويلات + UTM + مؤشّر أصلي'),
    ('٢', 'السيولة', '25 صيدلية + 15 سائقاً في نطاق واحد'),
    ('٣', 'الثقة', 'صفحة هبوط + محتوى «صيدلية مرخّصة = دواء آمن»'),
    ('٤', 'النمو', 'برنامج إحالة + حملة بصرية محدودة'),
]
y = Inches(2.0)
for num, t, d in steps:
    add_rect(s, Inches(0.7), y, Inches(0.7), Inches(0.7), fill=ACCENT, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.4)
    add_text(s, Inches(0.7), y + Inches(0.05), Inches(0.7), Inches(0.6),
             [{'text': num, 'size': 20, 'bold': True, 'color': DARK}], align=PP_ALIGN.CENTER)
    add_text(s, Inches(1.6), y + Inches(0.02), Inches(3.2), Inches(0.5),
             [{'text': t, 'size': 22, 'bold': True, 'color': WHITE}], align=PP_ALIGN.RIGHT)
    add_text(s, Inches(5.0), y + Inches(0.05), Inches(7.6), Inches(0.6),
             [{'text': d, 'size': 16, 'color': RGBColor(0xC8, 0xE0, 0xDF), 'rtl': True}], align=PP_ALIGN.RIGHT)
    y = y + Inches(1.05)
add_footer(s, 22)


# ================== شريحة 23: الخلاصة ==================
s = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(s, DARK)
add_rect(s, 0, 0, SW, Inches(0.35), fill=ACCENT)
add_text(s, Inches(0.6), Inches(1.6), Inches(12.1), Inches(1.2),
         [{'text': 'المعركة معركة ثقة، لا معركة أسعار.', 'size': 36, 'bold': True, 'color': WHITE, 'rtl': True, 'line_spacing': 1.1}],
         align=PP_ALIGN.RIGHT)
add_rect(s, Inches(0.6), Inches(3.0), Inches(3.5), Inches(0.10), fill=ACCENT)
add_text(s, Inches(0.6), Inches(3.4), Inches(12.1), Inches(2.4),
         [{'text': 'قِس أولاً، ثم اكسب السيولة في نطاقٍ واحد، ثم قدّمْ نفسك «القناة القانونية المعتمدة» التي تنقذ العراقي من فوضى الصفحات غير الموثوقة — وابنِ محرّك نموّ مستداماً (إحالة + واتساب + محتوى + ASO + شراكات).', 'size': 20, 'color': RGBColor(0xC8, 0xE0, 0xDF), 'rtl': True, 'line_spacing': 1.35}],
         align=PP_ALIGN.RIGHT)
add_text(s, Inches(0.6), Inches(5.6), Inches(12.1), Inches(0.9),
         [{'segments': [{'text': 'صِـلَة دوائي  •  ', 'size': 18, 'bold': True, 'color': ACCENT, 'rtl': True},
                        {'text': 'دواؤك… يصلك من صيدلية مرخّصة قريبة', 'size': 18, 'color': WHITE, 'rtl': True}]}],
         align=PP_ALIGN.RIGHT)
add_footer(s, 23)

out = '/home/user/Skills/docs/Sillaiq_Presentation.pptx'
prs.save(out)
print("SAVED:", out)
print("SLIDES:", len(prs.slides._sldIdLst))
