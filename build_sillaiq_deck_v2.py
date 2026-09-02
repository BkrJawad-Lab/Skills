# -*- coding: utf-8 -*-
"""
يبني برزنتيشن PowerPoint تسويقي شامل «الإصدار الثاني — 31 شريحة»
لتطبيق «صِلَة دوائي» (العراق) من الاستراتيجية الشاملة.

تحسينات v2 مقارنة بالنسخة الأولى (23 شريحة):
- غطاء الخلاصة التنفيذية (المعركة معركة ثقة) + ثلاث ركائز للبداية
- خارطة أول ٣٠ يوماً، سلوك المستهلك، تحليل منافسين مفصّل
- بيان التموضع الكامل + عبارة لكل طرف + 4 personas للمرضى
- تقسيم أكثر دقة للأونلاين (إنستغرام منفصل، نمو غير مكلّف)
- آليات النمو التشغيلية + 7Ps + سيناريوا الميزانية + بوابات الإطلاق
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# ---------- الهوية البصرية ----------
PRIMARY   = RGBColor(0x0E, 0x7C, 0x7B)   # تركوازي صحي
DARK      = RGBColor(0x0B, 0x25, 0x45)   # كحلي عميق
ACCENT    = RGBColor(0xE0, 0xA4, 0x58)   # ذهبي دافئ
LIGHT     = RGBColor(0xF4, 0xF7, 0xFA)   # خلفية فاتحة
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
GREY      = RGBColor(0x6B, 0x74, 0x84)
MUTED     = RGBColor(0x50, 0x5A, 0x6A)
TEAL_DARK = RGBColor(0x0A, 0x5C, 0x5B)
ICE       = RGBColor(0xC8, 0xE0, 0xDF)
SOFT      = RGBColor(0x9F, 0xB7, 0xC7)

FONT = "Segoe UI"

prs = Presentation()
prs.slide_width  = Inches(13.333)
prs.slide_height = Inches(7.5)
SW = prs.slide_width
SH = prs.slide_height

TOTAL = 31


def set_rtl(paragraph):
    pPr = paragraph._p.get_or_add_pPr()
    pPr.set('rtl', '1')
    pPr.set('bidi', '1')


def add_text(slide, x, y, w, h, runs, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.TOP,
             rtl=True, line_spacing=1.0):
    """runs: list of paragraphs; each = dict {text|segments,size,bold,color,align,space_before,line_spacing}"""
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
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
    add_rect(slide, 0, 0, SW, Inches(0.06), fill=accent_color)
    if kicker:
        add_text(slide, Inches(0.6), Inches(0.30), Inches(11.6), Inches(0.5),
                 [{'text': kicker, 'size': 13, 'bold': True, 'color': accent_color, 'rtl': True}],
                 align=PP_ALIGN.RIGHT)
    add_text(slide, Inches(0.6), Inches(0.58), Inches(12.1), Inches(0.95),
             [{'text': title, 'size': 29, 'bold': True, 'color': DARK}],
             align=PP_ALIGN.RIGHT)
    add_rect(slide, Inches(0.6), Inches(1.66), Inches(1.0), Inches(0.055), fill=accent_color)


def add_footer(slide, page):
    add_text(slide, Inches(0.5), Inches(7.08), Inches(12.3), Inches(0.3),
             [{'text': 'صِلَة دوائي — الاستراتيجية التسويقية الشاملة  |  v2', 'size': 10, 'color': GREY}],
             align=PP_ALIGN.RIGHT)
    add_text(slide, Inches(0.5), Inches(7.08), Inches(12.3), Inches(0.3),
             [{'text': f'{page} / {TOTAL}', 'size': 10, 'color': GREY}], align=PP_ALIGN.LEFT, rtl=False)


def slide_bg(slide, color=LIGHT):
    add_rect(slide, 0, 0, SW, SH, fill=color)


def bullets(slide, x, y, w, items, color=MUTED, dot=PRIMARY, size=15, gap=0.82,
            h=0.95, line_spacing=1.2, bold_last=False):
    """قائمة نقطية موحّدة"""
    yy = y
    for i, it in enumerate(items):
        add_rect(slide, x + Inches(0.05), yy + Inches(0.09), Inches(0.23), Inches(0.23),
                 fill=dot, shape=MSO_SHAPE.OVAL)
        add_text(slide, x + Inches(0.5), yy - Inches(0.02), w - Inches(0.5), h,
                 [{'text': it, 'size': size, 'color': color, 'rtl': True,
                   'line_spacing': line_spacing,
                   'bold': (bold_last and i == len(items) - 1)}],
                 align=PP_ALIGN.RIGHT)
        yy += Inches(gap)


def card(slide, x, y, w, h, header, hcolor, items, header_h=0.68, size=14,
         body_top=0.95, gap=0.95, item_h=0.9):
    """بطاقة بعنوان ملوّن وعناصر نقطية"""
    add_rect(slide, x, y, w, h, fill=WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.035)
    add_rect(slide, x, y, w, header_h, fill=hcolor)
    add_text(slide, x + Inches(0.3), y + Inches(0.12), w - Inches(0.6), Inches(0.5),
             [{'text': header, 'size': 17, 'bold': True, 'color': WHITE}], align=PP_ALIGN.RIGHT)
    yy = y + Inches(body_top)
    for it in items:
        add_rect(slide, x + Inches(0.35), yy + Inches(0.08), Inches(0.2), Inches(0.2),
                 fill=hcolor, shape=MSO_SHAPE.OVAL)
        add_text(slide, x + Inches(0.75), yy - Inches(0.02), w - Inches(1.05), item_h,
                 [{'text': it, 'size': size, 'color': MUTED, 'rtl': True, 'line_spacing': 1.18}],
                 align=PP_ALIGN.RIGHT)
        yy += Inches(gap)


def banner(slide, y, h, text, fill=PRIMARY, tcolor=WHITE, size=16, lead=None, lead_color=None):
    """شريط بيان عريض"""
    add_rect(slide, Inches(0.6), y, Inches(12.1), h, fill=fill, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.12)
    if lead:
        add_text(slide, Inches(0.9), y + Inches(h / 2 - 0.28), Inches(11.6), Inches(0.6),
                 [{'segments': [{'text': lead, 'size': size, 'bold': True,
                                 'color': lead_color or ACCENT, 'rtl': True},
                                {'text': text, 'size': size, 'bold': True, 'color': tcolor, 'rtl': True}]}],
                 align=PP_ALIGN.RIGHT)
    else:
        add_text(slide, Inches(0.9), y + Inches(h / 2 - 0.28), Inches(11.6), Inches(0.6),
                 [{'text': text, 'size': size, 'bold': True, 'color': tcolor, 'rtl': True}],
                 align=PP_ALIGN.RIGHT)


def table(slide, y0, header, rows, col_x, col_w, row_h=0.56, hdr_h=0.56,
          hdr_size=14, body_size=13, first_right=True, aligns=None, alt=LIGHT):
    """جدول مبني من مستطيلات — col_x: قائمة مواضع X، col_w: عرض الأعمدة"""
    yy = y0
    for cx, cw in zip(col_x, col_w):
        add_rect(slide, cx, yy, cw, hdr_h, fill=DARK)
    for j, htxt in enumerate(header):
        al = (aligns or [PP_ALIGN.CENTER] * len(header))[j]
        add_text(slide, col_x[j] + Inches(0.25), yy + Inches(0.06), col_w[j] - Inches(0.45), Inches(0.45),
                 [{'text': htxt, 'size': hdr_size, 'bold': True, 'color': WHITE}], align=al)
    yy = yy + hdr_h
    for i, row in enumerate(rows):
        bg = WHITE if i % 2 == 0 else alt
        for cx, cw in zip(col_x, col_w):
            add_rect(slide, cx, yy, cw, row_h, fill=bg)
        for j, cell in enumerate(row):
            al = (aligns or [PP_ALIGN.CENTER] * len(row))[j]
            bold = (j == 0 and first_right)
            colr = DARK if j == 0 else (PRIMARY if j == 1 else MUTED)
            add_text(slide, col_x[j] + Inches(0.25), yy + Inches(0.04), col_w[j] - Inches(0.45), row_h,
                     [{'text': cell, 'size': body_size, 'bold': bold, 'color': colr,
                       'rtl': True, 'line_spacing': 1.05}], align=al)
        yy = yy + row_h
    return yy


# ================== 1 — الغلاف ==================
s = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(s, DARK)
add_rect(s, 0, 0, SW, Inches(0.35), fill=ACCENT)
add_rect(s, SW - Inches(3.2), Inches(1.0), Inches(3.0), Inches(3.0), fill=TEAL_DARK, shape=MSO_SHAPE.OVAL)
add_rect(s, SW - Inches(2.0), Inches(0.0), Inches(1.6), Inches(1.6), fill=PRIMARY, shape=MSO_SHAPE.OVAL)
add_rect(s, -Inches(1.0), Inches(5.2), Inches(2.6), Inches(2.6), fill=TEAL_DARK, shape=MSO_SHAPE.OVAL)
add_text(s, Inches(1.0), Inches(1.3), Inches(9.5), Inches(0.6),
         [{'text': 'صِـلَة دوائي  |  Silla Al-Dawai', 'size': 20, 'bold': True, 'color': ACCENT, 'rtl': True}],
         align=PP_ALIGN.RIGHT)
add_text(s, Inches(1.0), Inches(2.2), Inches(11.0), Inches(2.2),
         [{'text': 'الاستراتيجية التسويقية الشاملة', 'size': 48, 'bold': True, 'color': WHITE, 'rtl': True, 'line_spacing': 1.05},
          {'text': 'منصّة الرعاية الصيدلانية الرقمية في العراق — أونلاين وأوفلاين', 'size': 23, 'color': ICE, 'rtl': True, 'space_before': 10}],
         align=PP_ALIGN.RIGHT)
add_rect(s, Inches(1.0), Inches(4.9), Inches(2.6), Inches(0.10), fill=ACCENT)
add_text(s, Inches(1.0), Inches(5.2), Inches(9.5), Inches(0.6),
         [{'segments': [{'text': 'الثقة  •  التوفّر الفعلي  •  الدفع عند الاستلام  •  النمو', 'size': 16, 'bold': True, 'color': WHITE}]}],
         align=PP_ALIGN.RIGHT)
add_text(s, Inches(1.0), Inches(6.35), Inches(11.0), Inches(0.5),
         [{'text': 'دواؤك… يصلك من صيدلية مرخّصة قريبة منك', 'size': 15, 'color': SOFT}],
         align=PP_ALIGN.RIGHT)
add_footer(s, 1)

# ================== 2 — خطة العرض ==================
s = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(s)
add_title(s, 'خطة العرض', kicker='Agenda')
items = [
    ('١', 'الخلاصة التنفيذية — من أين نبدأ؟'),
    ('٢', 'السوق العراقي — أرقام تحسم القرار'),
    ('٣', 'المشكلة الكبرى: فوضى الأدوية غير الموثوقة'),
    ('٤', 'المنافسون وموقع التمايز'),
    ('٥', 'التموضع وبيان العلامة'),
    ('٦', 'الجماهير المستهدفة: مريض، صيدلية، سائق'),
    ('٧', 'نموذج النموّ — الـ Flywheel'),
    ('٨', 'الإبداع أونلاين وأوفلاين'),
    ('٩', 'آليات النمو، المؤشّرات، ومزيج التسويق'),
    ('١٠', 'الخارطة، الميزانية، المخاطر — والبدء'),
]
y = Inches(1.98)
for num, txt in items:
    add_rect(s, Inches(0.7), y, Inches(0.55), Inches(0.5), fill=PRIMARY, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.4)
    add_text(s, Inches(0.7), y + Inches(0.02), Inches(0.55), Inches(0.45),
             [{'text': num, 'size': 15, 'bold': True, 'color': WHITE}], align=PP_ALIGN.CENTER)
    add_text(s, Inches(1.55), y + Inches(0.04), Inches(10.6), Inches(0.45),
             [{'text': txt, 'size': 18, 'color': DARK}], align=PP_ALIGN.RIGHT)
    y = y + Inches(0.52)
add_footer(s, 2)

# ================== 3 — الخلاصة التنفيذية ==================
s = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(s)
add_title(s, 'الخلاصة التنفيذية: المعركة معركة ثقة', kicker='Executive Summary')
add_rect(s, Inches(0.6), Inches(1.95), Inches(12.1), Inches(2.5), fill=WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.04)
add_rect(s, Inches(0.6), Inches(1.95), Inches(12.1), Inches(0.14), fill=PRIMARY)
add_text(s, Inches(0.95), Inches(2.3), Inches(11.5), Inches(2.0),
         [{'segments': [{'text': 'السوق العراقي غارقٌ بصفحاتٍ وتطبيقات تبيع الأدوية عبر فيسبوك وإنستغرام وتيليغرام ', 'size': 17, 'color': MUTED, 'rtl': True},
                        {'text': 'بلا ترخيص ولا رقابة', 'size': 17, 'bold': True, 'color': DARK, 'rtl': True},
                        {'text': ' — وقد حذّرت نقابة صيادلة العراق صراحةً من هذه الظاهرة لأنّها تمثّل خطراً على حياة المرضى (أدوية مغشوشة، منتهية الصلاحية، حسابات تنتحل اسم أطباء).', 'size': 17, 'color': MUTED, 'rtl': True}],
           'line_spacing': 1.4}],
         align=PP_ALIGN.RIGHT)
banner(s, Inches(4.75), Inches(0.95), '«صِلَة دوائي = القناة القانونية المعتمدة الوحيدة»',
       fill=DARK, size=19, lead='الموقع الذي لا يملكه أيّ منافس: ', lead_color=ACCENT)
bullets(s, Inches(0.7), Inches(5.95), Inches(12.0),
        ['كل قرار في هذه الاستراتيجية يجب أن يصدح بهذه الحقيقة — قبل أي قرار سعر أو حملة.',
         'ثلاث ركائز للبداية (تفصيل في الشريحة التالية): قِسْ أولاً  •  كثّف جغرافياً  •  اصنع الفجوة التنظيمية صناعةً لا عائقاً.'],
        size=14.5, gap=0.62, h=0.55)
add_footer(s, 3)

# ================== 4 — ثلاث ركائز للبداية ==================
s = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(s)
add_title(s, 'ثلاث ركائز للبداية — بالترتيب', kicker='Starting Principles')
pillars = [
    ('١ — قِسْ قبل أن تنفق', 'النِسبةُ الأصلية للنجاح هي الطلب المكتمل المدفوع، لا عدد التنزيلات.',
     'فعّل AppsFlyer/Firebase + GA4 + UTM على كل رابط ورمز QR قبل أول دينار.', PRIMARY),
    ('٢ — كثافة جغرافية', 'مشكلة السوق الثنائي (البيضة والدجاجة) تُحَلّ بالتكثّف في نطاق واحد (بغداد).',
     'املأه صيدليات وسائقين حتى يتحقّق «توفّر عالٍ» — ولا حملات طلب واسعة قبل رصيد عرضٍ حقيقي.', ACCENT),
    ('٣ — اصنع الفجوة التنظيمية', 'تحوّل التحذير القانوني من عائقٍ منافسٍ إلى ميزةٍ وُجدت لأجلنا.',
     'حملة «دواؤك يصلك من صيدلية مرخّصة» تتصدّى للصفحات غير الموثوقة: تخدم المريض وتبني العلامة.', TEAL_DARK),
]
for i, (t, d1, d2, c) in enumerate(pillars):
    x = Inches(0.6) + i * Inches(4.2)
    add_rect(s, x, Inches(2.0), Inches(3.9), Inches(4.6), fill=WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.05)
    add_rect(s, x, Inches(2.0), Inches(3.9), Inches(0.95), fill=c)
    add_text(s, x + Inches(0.25), Inches(2.18), Inches(3.4), Inches(0.6),
             [{'text': t, 'size': 18, 'bold': True, 'color': WHITE}], align=PP_ALIGN.RIGHT)
    add_text(s, x + Inches(0.3), Inches(3.2), Inches(3.3), Inches(1.3),
             [{'text': d1, 'size': 14, 'color': DARK, 'rtl': True, 'line_spacing': 1.25}], align=PP_ALIGN.RIGHT)
    add_text(s, x + Inches(0.3), Inches(4.75), Inches(3.3), Inches(1.7),
             [{'text': d2, 'size': 13, 'color': MUTED, 'rtl': True, 'line_spacing': 1.25}], align=PP_ALIGN.RIGHT)
add_footer(s, 4)

# ================== 5 — خارطة أول 30 يوماً ==================
s = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(s)
add_title(s, 'خارطة أول ٣٠ يوماً', kicker='First 30 Days')
rows = [
    ('١', 'القياس + العرض', 'تتبّع كامل من التنزيل إلى الطلب؛ تجهيز ١٠–٢٥ صيدلية مرخّصة مؤكّدة في النطاق؛ ضبط «التوفّر الفعلي» وهدف زمن التلبية.'),
    ('٢', 'تجربة التسليم', 'توظيف ١٠–١٥ سائقاً مؤكّداً؛ ضبط زمن التوصيل؛ إصدار دليل «صفر تأخير» للسائقين.'),
    ('٣', 'أول محتوى ثقة + إحالة', 'صفحة هبوط عربية؛ محتوى «صيدلية مرخّصة = دواء آمن»؛ تفعيل برنامج الإحالة.'),
    ('٤', 'أول حملة بصرية', 'حملة محلية محدودة (ميتا/إنستغرام/TikTok) في نفس النطاق، برسالة الثقة والدفع عند الاستلام — مع مراقبة CAC مقابل القيمة.'),
]
y0 = Inches(2.0)
col_x = [Inches(0.6), Inches(2.0), Inches(4.7)]
col_w = [Inches(1.4), Inches(2.7), Inches(8.0)]
yy = table(s, y0, ('الأسبوع', 'الأولوية', 'الإجراء'), rows, col_x, col_w,
           row_h=0.95, hdr_h=0.56, body_size=13.5,
           aligns=[PP_ALIGN.CENTER, PP_ALIGN.RIGHT, PP_ALIGN.RIGHT])
add_rect(s, Inches(0.6), Inches(6.35), Inches(12.1), Inches(0.55), fill=DARK, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.3)
add_text(s, Inches(0.9), Inches(6.43), Inches(11.6), Inches(0.4),
         [{'segments': [{'text': 'القاعدة: ', 'size': 14, 'bold': True, 'color': ACCENT, 'rtl': True},
                        {'text': 'لا زبوناً يدفع له إعلان ليجد «لا توفّر» — يرتدّ ولا يعود.', 'size': 14, 'color': WHITE, 'rtl': True}]}],
         align=PP_ALIGN.RIGHT)
add_footer(s, 5)

# ================== 6 — أرقام السوق ==================
s = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(s)
add_title(s, 'السوق العراقي — أرقام تحسم القرار', kicker='Iraq Market  •  أوائل 2025')
rows = [
    ('عدد السكان', '~46.5 مليون', '72% في المدن — كتلة سوق حضرية'),
    ('مستخدمو الإنترنت', '~38.0 مليون', '81.7% — بنية رقمية واسعة'),
    ('هويّات التواصل الاجتماعي', '~34.3 مليون', '73.8% — جمهور ضخم للتعليم الذاتي'),
    ('تيك توك', 'يبلغ 90% من مستخدمي الإنترنت', 'القناة الأوسع انتشاراً في العراق'),
    ('فيسبوك', '~20.1 مليون', '75% من البالغين — الأشد ثقة للجمهور الأوسع'),
    ('إنستغرام', '~19.0 مليون', '69% من البالغين — بصرية وحياة'),
    ('سناب شات', '~18.5 مليون', 'جمهور شاب جداً'),
    ('العمر الوسيط', '20.8 سنة', 'سوق شابّ — والدواء يمسّ العائلات كلّها'),
]
col_x = [Inches(0.6), Inches(4.9), Inches(9.2)]
col_w = [Inches(4.3), Inches(4.3), Inches(3.5)]
table(s, Inches(1.95), ('المؤشّر', 'القيمة', 'الدلالة التسويقية'), rows, col_x, col_w,
      row_h=0.55, hdr_h=0.54, body_size=13,
      aligns=[PP_ALIGN.RIGHT, PP_ALIGN.CENTER, PP_ALIGN.RIGHT])
add_footer(s, 6)

# ================== 7 — الفوضى غير الموثوقة ==================
s = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(s)
add_title(s, 'المشكلة الكبرى: فوضى الأدوية غير الموثوقة', kicker='Hidden Opportunity', accent_color=ACCENT)
add_text(s, Inches(0.6), Inches(1.9), Inches(12.1), Inches(0.5),
         [{'text': 'نقابة صيادلة العراق تُحذّر المواطنين رسمياً من شراء الأدوية ومستحضرات العناية والمكمّلات عبر صفحات التواصل', 'size': 17, 'bold': True, 'color': DARK, 'rtl': True}],
         align=PP_ALIGN.RIGHT)
add_rect(s, Inches(0.6), Inches(2.5), Inches(12.1), Inches(0.06), fill=ACCENT)
bullets(s, Inches(0.7), Inches(2.8), Inches(11.9),
        ['البيع الإلكتروني للأدوية المسجّلة خارج صيدلية مرخّصة مخالفٌ للقانون — والدواء يحتاج نصيحة صيدلي وتوضيح محاذير الاستعمال.',
         'تحذير النقابة صريح: «أي إعلان أو ترويج عبر مواقع التواصل مخالفةٌ يحاسب عليها المعلن» — والصيدلي المخالف يُحاسَب نقابياً.',
         'تحقيقات صحفية تكشف حسابات وهمية على فيسبوك تنتحل صفة أطباء معروفين: علاجات وهمية، أدوية منتهية الصلاحية، وخلطات عشبية — وعجز عن ملاحقتها.',
         'المصدر الشرعي الوحيد للدواء بحكم التعريف: الصيدلية المرخّصة الخاضعة للرقابة.'],
        size=15, gap=0.72, h=0.7)
banner(s, Inches(5.9), Inches(0.95), '«صِلَة دوائي = القناة القانونية المعتمدة» — حجّةٌ لا يملكها أحد، ودعايةٌ مجانية إن أُحسنت صياغتها كحملة توعية وطنية.',
       fill=PRIMARY, size=15.5)
add_footer(s, 7)

# ================== 8 — سلوك المستهلك ==================
s = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(s)
add_title(s, 'سلوك المستهلك العراقي', kicker='Consumer Behavior')
stats = [
    ('60–70%', 'من عمليات الشراء الإلكترونية ما تزال نقداً عند الاستلام — «حماية لحقّك».',
     'صِلَة تتبنّاه منذ اليوم: ميزةٌ لا عيب.', PRIMARY),
    ('ثقة قبل سعر', 'غياب التحقّق يرفع الاحتيال؛ فالكسب هو للطرف «الموثّق» لا «الأرخص».',
     'التوثيق هو سعرُ الثقة في هذا السوق.', ACCENT),
    ('محادثة قبل صفحة', 'الأعمال والمشتريات تُدار عبر واتساب/تيليغرام أكثر من الصفحات.',
     'قناة الطلب/التأكيد عبر المحادثة المألوفة.', TEAL_DARK),
]
for i, (big, d, act, c) in enumerate(stats):
    x = Inches(0.6) + i * Inches(4.2)
    add_rect(s, x, Inches(2.0), Inches(3.9), Inches(3.5), fill=WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.05)
    add_rect(s, x, Inches(2.0), Inches(3.9), Inches(0.14), fill=c)
    add_text(s, x + Inches(0.3), Inches(2.35), Inches(3.3), Inches(0.7),
             [{'text': big, 'size': 24, 'bold': True, 'color': c}], align=PP_ALIGN.RIGHT)
    add_text(s, x + Inches(0.3), Inches(3.25), Inches(3.3), Inches(1.2),
             [{'text': d, 'size': 13.5, 'color': MUTED, 'rtl': True, 'line_spacing': 1.25}], align=PP_ALIGN.RIGHT)
    add_text(s, x + Inches(0.3), Inches(4.6), Inches(3.3), Inches(0.8),
             [{'text': act, 'size': 13, 'bold': True, 'color': DARK, 'rtl': True, 'line_spacing': 1.2}], align=PP_ALIGN.RIGHT)
add_rect(s, Inches(0.6), Inches(5.85), Inches(12.1), Inches(0.9), fill=DARK, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.15)
add_text(s, Inches(0.9), Inches(6.05), Inches(11.6), Inches(0.55),
         [{'segments': [{'text': 'مدن الطلب الأعلى (البدء والتوسّع):  ', 'size': 15, 'bold': True, 'color': ACCENT, 'rtl': True},
                        {'text': 'بغداد  •  البصرة  •  أربيل  •  الموصل', 'size': 16, 'bold': True, 'color': WHITE, 'rtl': True}]}],
         align=PP_ALIGN.RIGHT)
add_footer(s, 8)

# ================== 9 — تحليل المنافسين ==================
s = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(s)
add_title(s, 'تحليل المنافسين — وموقع التمايز', kicker='Competitive Landscape')
rows = [
    ('فلق (Falaq) — الموصل', 'سوق B2B يربط الصيدليات بالمذاخر والشركات', 'صِلَة تخدم المريض النهائي B2C بسوق عروض وتوصيل للباب'),
    ('فاركس (Pharx) — الموصل', 'يربط المريض بشبكة صيدليات للبحث عن التوفّر', 'صِلَة تضيف: مقارنة عروض، دفع عند الاستلام، سائقين، منظومة متكاملة'),
    ('Iraq Gate Pharmacy', 'متجر أونلاين يعتمد غالباً على صيدلية/متجر واحد', 'صِلَة = سوق عروض متعدّدة الصيدليات'),
    ('صيدليات الدواء (Al-Dawaa)', 'سلاسل راسخة خليجية مع توصيل وولاء', 'مؤشّر الاتجاه الإقليمي — ليس المنافس العراقي المباشر'),
    ('صفحات/تطبيقات غير موثوقة', 'تبيع الأدوية بلا ترخيص عبر فيسبوك/تيليغرام', 'التميّز الأكبر: القناة القانونية المعتمدة — حجّةٌ لا يملكها أحد'),
]
col_x = [Inches(0.6), Inches(4.3), Inches(8.3)]
col_w = [Inches(3.7), Inches(4.0), Inches(4.4)]
yy = table(s, Inches(1.95), ('المنافس', 'طبيعته', 'تمايز صِلَة في مواجهته'), rows, col_x, col_w,
           row_h=0.72, hdr_h=0.56, body_size=12.5,
           aligns=[PP_ALIGN.RIGHT, PP_ALIGN.RIGHT, PP_ALIGN.RIGHT])
add_text(s, Inches(0.6), Inches(6.45), Inches(12.1), Inches(0.5),
         [{'segments': [{'text': 'ملاحظة منهجية: ', 'size': 12, 'bold': True, 'color': ACCENT, 'rtl': True},
                        {'text': 'قائمة مؤشّرات لا حقائق نهائية — يُتحقّق من نطاق كل منافس ببياناتٍ محدّثة.', 'size': 12, 'color': GREY, 'rtl': True}]}],
         align=PP_ALIGN.RIGHT)
add_footer(s, 9)

# ================== 10 — بيان التموضع ==================
s = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(s)
add_title(s, 'بيان التموضع (للمريض)', kicker='Positioning Statement')
add_rect(s, Inches(0.6), Inches(2.0), Inches(12.1), Inches(3.4), fill=WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.04)
add_rect(s, Inches(0.6), Inches(2.0), Inches(0.18), Inches(3.4), fill=PRIMARY)
add_text(s, Inches(1.1), Inches(2.35), Inches(11.2), Inches(2.8),
         [{'segments': [
             {'text': '«صِلَة دوائي — منصّة الرعاية الصيدلانية في العراق. دواؤك يصلك من ', 'size': 19, 'color': MUTED, 'rtl': True},
             {'text': 'صيدليات مرخّصة قريبة', 'size': 19, 'bold': True, 'color': PRIMARY, 'rtl': True},
             {'text': '، تُقارن عروضها فعلياً، وتدفع ', 'size': 19, 'color': MUTED, 'rtl': True},
             {'text': 'نقداً عند الاستلام', 'size': 19, 'bold': True, 'color': PRIMARY, 'rtl': True},
             {'text': '. ثقتك من ورائها سلطةٌ لا تُنتحل: صيدليّة مرخّصة، ودواءٌ فعلي التوفّر — ', 'size': 19, 'color': MUTED, 'rtl': True},
             {'text': 'لا وعودٌ على صفحات مجهولة.»', 'size': 19, 'bold': True, 'color': DARK, 'rtl': True},
         ]}],
         align=PP_ALIGN.RIGHT, line_spacing=1.45)
banner(s, Inches(5.85), Inches(1.0), 'صِلَة ليست تطبيق توصيل، بل شبكة صحية مُدارة ومرخّصة.',
       fill=DARK, size=19, lead='«لأنّ صحتك لا تحتمل المجازفة» — ', lead_color=ACCENT)
add_footer(s, 10)

# ================== 11 — عبارة لكل طرف + الركائز ==================
s = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(s)
add_title(s, 'عبارة التموضع لكل طرف — والركائز الثلاث', kicker='One Brand  •  Three Voices')
voices = [
    ('للمريض', '«قارن العروض، وادفع عند الاستلام، واطمئنّ: صيدليّة مرخّصة لا صفحة مجهولة.»', PRIMARY),
    ('للصيدليّة', '«قناة رقمية إضافية… وأنت صاحب السعر والقرار.»', ACCENT),
    ('للسائق', '«دخلٌ مرن، ودورٌ يوصّل دواءً لا مجرّد طرود.»', TEAL_DARK),
]
y = Inches(1.95)
for t, q, c in voices:
    add_rect(s, Inches(0.6), y, Inches(12.1), Inches(0.95), fill=WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.12)
    add_rect(s, Inches(11.6), y, Inches(1.1), Inches(0.95), fill=c)
    add_text(s, Inches(11.6), y + Inches(0.25), Inches(1.1), Inches(0.5),
             [{'text': t, 'size': 15, 'bold': True, 'color': WHITE}], align=PP_ALIGN.CENTER)
    add_text(s, Inches(0.9), y + Inches(0.18), Inches(10.4), Inches(0.6),
             [{'text': q, 'size': 16, 'bold': True, 'color': DARK, 'rtl': True}], align=PP_ALIGN.RIGHT)
    y += Inches(1.12)
add_text(s, Inches(0.6), Inches(5.45), Inches(12.1), Inches(0.45),
         [{'text': 'الركائز الثلاث لعلامة «صِلَة»', 'size': 17, 'bold': True, 'color': PRIMARY, 'rtl': True}],
         align=PP_ALIGN.RIGHT)
pill_data = [
    ('الثقة المرخّصة', 'الجميع موثّقون'),
    ('التوفّر الفعلي', 'لا وعود — عرضٌ حقيقي'),
    ('الدفع عند الاستلام', 'معاملتك بيدك'),
]
for i, (t, d) in enumerate(pill_data):
    x = Inches(0.6) + i * Inches(4.2)
    add_rect(s, x, Inches(5.95), Inches(3.9), Inches(0.85), fill=DARK, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.18)
    add_text(s, x + Inches(0.3), Inches(6.08), Inches(3.3), Inches(0.6),
             [{'segments': [{'text': t + '  ', 'size': 15, 'bold': True, 'color': ACCENT, 'rtl': True},
                            {'text': '— ' + d, 'size': 12.5, 'color': WHITE, 'rtl': True}]}],
             align=PP_ALIGN.RIGHT)
add_footer(s, 11)

# ================== 12 — المرضى (4 personas) ==================
s = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(s)
add_title(s, 'الجماهير: المرضى (B2C) — 4 Personas', kicker='Audience')
rows = [
    ('مريض مزمن / مسنّ / مقيّد الحركة', 'يعتمد أدويةً باستمرار أو لا يقدر على الذهاب', 'استمرارية الدواء + التوفّر', '«لا تنقطع أدويتك. نوصلها لك.»'),
    ('أمّ / ربّ أسرة', 'تدير صحة العائلة (أطفال، كبار)', 'راحة، ثقة، وسرعة', '«دواؤك يصلك وأنت في بيتك.»'),
    ('الموظف / العامل', 'لا يملك وقتاً', 'توفير الوقت والسعر والتوفّر', '«اطلب من مكتبك، يصل باب بيتك.»'),
    ('الشاب الرقمي', 'مرتاح للمنصّات ويقارن', 'الشفافية والسعر', '«هل دواؤك متوفّر لديهم فعلاً؟ شوف بنفسك.»'),
]
col_x = [Inches(0.6), Inches(3.75), Inches(7.3), Inches(10.4)]
col_w = [Inches(3.15), Inches(3.55), Inches(3.1), Inches(2.33)]
table(s, Inches(1.95), ('الشريحة', 'من هم؟', 'المحفّز الأكبر', 'رسالتنا'), rows, col_x, col_w,
      row_h=0.85, hdr_h=0.56, body_size=12.5,
      aligns=[PP_ALIGN.RIGHT, PP_ALIGN.RIGHT, PP_ALIGN.RIGHT, PP_ALIGN.RIGHT])
add_rect(s, Inches(0.6), Inches(6.3), Inches(12.1), Inches(0.6), fill=DARK, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.25)
add_text(s, Inches(0.9), Inches(6.4), Inches(11.6), Inches(0.4),
         [{'segments': [{'text': 'حواجز القرار: ', 'size': 13.5, 'bold': True, 'color': ACCENT, 'rtl': True},
                        {'text': 'الثقة بالترخيص  •  التوفّر  •  وضوح السعر  •  خصوصية الوصفة  •  الدفع عند الاستلام', 'size': 13.5, 'color': WHITE, 'rtl': True}]}],
         align=PP_ALIGN.RIGHT)
add_footer(s, 12)

# ================== 13 — الصيدليات + السائقون ==================
s = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(s)
add_title(s, 'الجماهير: الصيدليات (B2B) والسائقون (Gig)', kicker='Audience')
card(s, Inches(0.6), Inches(2.0), Inches(6.0), Inches(4.6), 'الصيدليات', PRIMARY,
     ['صيدلية مستقلة صغيرة — تطلب أوامر إضافية بلا مصروف تسويق، وتهتمّ باستقلاليّتها.',
      'محفّزها: «أوامر إضافية بلا جهد تسويقي، وأنت من يحدّد الأسعار».',
      'صيدلية متوسّطة / سلسلة — تريد توسّعاً رقمياً وسجلّ تسويات واضح.',
      'محفّزها: «لوحة متكاملة وتسوية شفّافة ومرجعية».',
      'البوابة: B2B ميداني — فريق + Demo + Onboarding، ابدأ بالصيدليات ذات المرجعية.'],
     size=14, gap=0.8, item_h=0.85)
card(s, Inches(6.9), Inches(2.0), Inches(5.8), Inches(4.6), 'السائقون', TEAL_DARK,
     ['الباحثون عن دخل مرن: طلاّب، مستقلّون، أصحاب سيارات ودراجات.',
      'محفّزهم: المرونة، الأرباح الواضحة، والأوامر القريبة.',
      'القيمة المعنوية: «دورٌ يوصّل دواءً يغيّر يوم إنسان» — سببٌ للعمل لا مجرّد أجرة.',
      'آلية النمو: إحالة سائق + شفافية الأرباح + ترقية «سائق يجلب سائقاً».',
      'التدريب: مهنيّة، سرّيّة، وزمن «صفر تأخير» — التسليم = وجهُ العلامة.'],
     size=14, gap=0.8, item_h=0.85)
add_footer(s, 13)

# ================== 14 — الـ Flywheel ==================
s = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(s)
add_title(s, 'نموذج النموّ — الـ Flywheel', kicker='Growth Model')
banner(s, Inches(1.9), Inches(0.85), '«الكثافة قبل الاتّساع» (Density-First) — ابدأ بنطاقٍ واحد مكثّف حتى يتحقّق توفّرٌ عالٍ، ثم توسّع.',
       fill=PRIMARY, size=15)
boxes = [
    ('توفّرٌ عالٍ', Inches(8.9), Inches(3.15)), ('تجربة ممتازة', Inches(8.9), Inches(4.35)),
    ('كلمة شفاهة / إحالة', Inches(8.9), Inches(5.55)), ('مزيدٌ من الزبائن', Inches(4.7), Inches(3.15)),
    ('أوامر أكثر', Inches(4.7), Inches(4.35)), ('عرضٌ أقوى: صيدليات وسائقون', Inches(4.7), Inches(5.55)),
]
for t, x, y in boxes:
    add_rect(s, x, y, Inches(3.9), Inches(0.95), fill=DARK, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.35)
    add_text(s, x, y + Inches(0.18), Inches(3.9), Inches(0.6),
             [{'text': t, 'size': 15, 'bold': True, 'color': WHITE}], align=PP_ALIGN.CENTER)
add_rect(s, Inches(0.6), Inches(3.05), Inches(3.7), Inches(3.55), fill=WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.05)
add_text(s, Inches(0.85), Inches(3.3), Inches(3.2), Inches(0.5),
         [{'text': 'لماذا يعمل؟', 'size': 17, 'bold': True, 'color': PRIMARY, 'rtl': True}], align=PP_ALIGN.RIGHT)
add_text(s, Inches(0.85), Inches(3.95), Inches(3.2), Inches(2.5),
         [{'text': 'كلّ صيدلية وسائقٍ جديد في نفس النطاق يزيد التوفّر، يقلّل تكلفة الجذب، ويصعّد الثقة — قيمةٌ تتزايد مع كل دورة.', 'size': 13.5, 'color': MUTED, 'rtl': True, 'line_spacing': 1.3}],
         align=PP_ALIGN.RIGHT)
add_rect(s, Inches(0.6), Inches(6.15), Inches(12.1), Inches(0.7), fill=ACCENT, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.3)
add_text(s, Inches(0.9), Inches(6.28), Inches(11.6), Inches(0.45),
         [{'segments': [{'text': 'التبكير (Seeding): ', 'size': 14, 'bold': True, 'color': DARK, 'rtl': True},
                        {'text': 'أطلق في مناطق طلبٍ شديد — بجوار المستشفيات والعيادات، أحياء سكنية كثيفة، مناطق تجارية — وأعطها أولويّة التغطية.', 'size': 14, 'color': DARK, 'rtl': True}]}],
         align=PP_ALIGN.RIGHT)
add_footer(s, 14)

# ================== 15 — الأونلاين نظرة عامة ==================
s = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(s)
add_title(s, 'الأونلاين — منظومة القنوات الإبداعية', kicker='Online Creative System')
channels = [
    ('تيك توك / ريلز', 'القناة الأوسع (90%) — فيديوهات قصيرة حقيقية', PRIMARY),
    ('فيسبوك + ماسنجر', 'الثقة والجمهور الأوسع — محتوى تحذير + بوت', PRIMARY),
    ('إنستغرام', 'بصرية وحياة — ستوريات يومية ومؤثّرون محليون', ACCENT),
    ('واتساب / تيليغرام', 'بيع وخدمة وتأكيد أوامر + بديل iOS', TEAL_DARK),
    ('ASO + SEO', 'نموّ غير مكلّف من المتجر والبحث العربي', PRIMARY),
    ('إحالة + ولاء', 'الدواء شأنٌ عائلي — أعلى تحويل في القطاع الصحي', ACCENT),
]
for i, (t, d, c) in enumerate(channels):
    row = i // 3
    col = i % 3
    x = Inches(0.6) + col * Inches(4.2)
    y = Inches(2.15) + row * Inches(2.35)
    add_rect(s, x, y, Inches(3.9), Inches(2.05), fill=WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.06)
    add_rect(s, x, y, Inches(0.16), Inches(2.05), fill=c)
    add_text(s, x + Inches(0.4), y + Inches(0.28), Inches(3.3), Inches(0.6),
             [{'text': t, 'size': 18, 'bold': True, 'color': DARK}], align=PP_ALIGN.RIGHT)
    add_text(s, x + Inches(0.4), y + Inches(1.0), Inches(3.3), Inches(0.95),
             [{'text': d, 'size': 13, 'color': MUTED, 'rtl': True, 'line_spacing': 1.25}], align=PP_ALIGN.RIGHT)
add_footer(s, 15)

# ================== 16 — الحملة الكبرى ==================
s = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(s)
add_title(s, 'الحملة الكبرى: «صيدليتُك لا صفحةٌ مجهولة»', kicker='Flagship Campaign', accent_color=ACCENT)
banner(s, Inches(1.95), Inches(0.95), 'حملة توعوية وطنية تتصدّى لمخاطر صفحات الأدوية المزوّرة — مستندة إلى تحذير نقابة الصيادلة؛ تخدم المريض وتبني العلامة في آنٍ واحد.',
       fill=ACCENT, tcolor=DARK, size=15.5)
bullets(s, Inches(0.7), Inches(3.3), Inches(11.9),
        ['سلسلة محتوى «كيف تتحقق أن الدواء أصلي؟» — الملصق السعري، المصدر، العيوب، تاريخ الصلاحية: تعليم + بناء ثقة.',
         'فيديوهات «الفرق بين صيدلية مرخّصة وصفحة وهمية» بأسلوب قصصي (شخصٌ يعاني من دواء منتهٍ الصلاحية).',
         'محتوى «جناح الثقة»: لقطات حقيقية من رحلة التوثيق داخل صِلَة (مراجعة الترخيص والوثائق).',
         'هاشتاغٌ حامل للحملة — بمراجعة ناطق وقانوني للصياغة قبل النشر.'],
        size=15.5, gap=0.85, h=0.8, dot=ACCENT)
add_footer(s, 16)

# ================== 17 — تيك توك ==================
s = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(s)
add_title(s, 'الأونلاين: تيك توك / ريلز — الأوسع انتشاراً', kicker='~90% من مستخدمي الإنترنت')
banner(s, Inches(1.95), Inches(0.85), 'محتوى قصير، حقيقي، موثوق — بلا ادّعاءات صحية خارج الرقابة.', fill=PRIMARY, size=15.5)
bullets(s, Inches(0.7), Inches(3.15), Inches(11.9),
        ['«دواؤك موجود؟» — مقاطع قصيرة تبحث عن دواءٍ نادر/شائع وتُظهر توفّره الفعلي في الصيدليات.',
         '«رحلة الأمر في ١٠ ثوانٍ» — من تصوير الوصفة إلى استلام الطرد (By-Example).',
         'سلسلة «ضحّينا على الصفحات المزوّرة» — بفكاهة محلية حذرة تشرح الفخاخ.',
         'قصص حقيقية من المستخدمين (UGC) — مرضى وأهاليهم، بدعم تعاوني.',
         'تحدٍّ/مسابقة بوسمٍ خاص — بشروط مراجعة الرعاة وعدم التضليل.'],
        size=15, gap=0.8, h=0.75)
add_footer(s, 17)

# ================== 18 — فيسبوك + ماسنجر + واتساب ==================
s = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(s)
add_title(s, 'الأونلاين: فيسبوك + ماسنجر + واتساب', kicker='Trust & Conversation')
card(s, Inches(0.6), Inches(2.0), Inches(6.0), Inches(4.6), 'المحتوى والثقة', PRIMARY,
     ['استثمار تحذير النقابة: «الدواء يُعرض في الصيدلية، لا في صفحة» + طمأنة بجهةٍ مرخّصة.',
      'دليل توثيق مرئي يرافق كل ادّعاء ثقة.',
      'مجموعات المجتمع المحلي/الأمّهات — حملات إرشاد ووعي صحي.',
      'إعلانات مخصّصة للأحياء المستهدفة في نطاق الإطلاق.'],
     size=14, gap=0.85, item_h=0.95)
card(s, Inches(6.9), Inches(2.0), Inches(5.8), Inches(4.6), 'المحادثة والأتمتة', ACCENT,
     ['بوت ماسنجر/واتساب: تأكيد الأمر، تتبّع الحالة، ودعم مباشر — بيئة محادثة مألوفة تقلّل الاحتكاك.',
      'جروب واتساب/تيليغرام رسمي: خدمة وأوامر وتعويضٍ لمستخدمي iOS قبل الإطلاق.',
      'فصل الخدمة عن التسويق — قناة خدمة مستقلة بقواعد إلغاء اشتراك وخصوصية.'],
     size=14, gap=0.85, item_h=0.95)
add_footer(s, 18)

# ================== 19 — إنستغرام ==================
s = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(s)
add_title(s, 'الأونلاين: إنستغرام — البصرية والحياة', kicker='Visual & Lifestyle', accent_color=ACCENT)
cards = [
    ('ستوريات يومية', '«ماذا في أمرك؟» — محتوى حيّ يومي: تجهيز الطرد، لحظة التسليم، وجه السائق، شكر الزبون.'),
    ('هايلايتس دائمة', '«كيف تعمل صِلَة» + «لماذا نحن آمنون» — إجابة كل سؤال ثقة قبل أن يُطرح.'),
    ('مؤثّرون محليون', 'صحة، أمومة، حياة، أطباء — بإفصاح واضح عن الرعاية ومراجعة كل ادّعاء قبل النشر.'),
]
for i, (t, d) in enumerate(cards):
    x = Inches(0.6) + i * Inches(4.2)
    add_rect(s, x, Inches(2.15), Inches(3.9), Inches(3.6), fill=WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.05)
    add_rect(s, x, Inches(2.15), Inches(3.9), Inches(0.9), fill=ACCENT)
    add_text(s, x + Inches(0.25), Inches(2.33), Inches(3.4), Inches(0.55),
             [{'text': t, 'size': 18, 'bold': True, 'color': DARK}], align=PP_ALIGN.RIGHT)
    add_text(s, x + Inches(0.3), Inches(3.3), Inches(3.3), Inches(2.3),
             [{'text': d, 'size': 14, 'color': MUTED, 'rtl': True, 'line_spacing': 1.35}], align=PP_ALIGN.RIGHT)
banner(s, Inches(6.15), Inches(0.75), 'قاعدة المؤثّرين: إفصاح واضح + مراجعة ادّعاءات + لا نصيحة طبية إطلاقاً.',
       fill=DARK, size=14.5)
add_footer(s, 19)

# ================== 20 — نمو غير مكلّف ==================
s = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(s)
add_title(s, 'الأونلاين: نموّ غير مكلّف — ASO + SEO + إحالة + ولاء', kicker='Compounding Growth')
card(s, Inches(0.6), Inches(2.0), Inches(6.0), Inches(4.6), 'ASO + SEO', PRIMARY,
     ['كلمات عربية عالية النية: «صيدلية أونلاين»، «توصيل أدوية بغداد»، «اطلب دواء»، «صيدلية قريبة» + أسماء أدوية شائعة.',
      'أيقونة + لقطات شاشة عربية دعائية: «عروض الصيدليات» و«الدفع عند الاستلام».',
      'تقييمات ومراجعات حقيقية + ردّ على كل تقييم.',
      'مدوّنة/أسئلة شائعة Long-tail: «أين أشتري دواء الفلاني؟»، «كيف أرفع وصفتي؟»، «مخاطر شراء الدواء أونلاين».',
      'ربط صفحة Google Play + مشروع iOS.'],
     size=13, gap=0.8, item_h=0.85)
card(s, Inches(6.9), Inches(2.0), Inches(5.8), Inches(4.6), 'إحالة + ولاء + عرض حارس', ACCENT,
     ['كود إحالة عائلي — «الدواء شأنٌ عائلي»: مكافأة الموصي والموصى إليه معاً.',
      'برنامج ولاء: نقاط لكل أمر وتوصية — يُغذّي التكرار لا الشراء الأول فقط.',
      'عرض حارس: خصم محدود لأول أمر + شحن مجاني لفترة قصيرة — الإغراء المفرط يجلب زبائن لا يعودون.',
      'قياس الأثر بلا إفراط في الخصم.'],
     size=13, gap=0.8, item_h=0.85)
add_footer(s, 20)

# ================== 21 — الأوفلاين نظرة عامة ==================
s = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(s)
add_title(s, 'الأوفلاين — حضورٌ ملموس في الحيّ والشارع', kicker='Offline Presence', accent_color=TEAL_DARK)
items = [
    ('الشراكات الصحية', 'عيادات، مستشفيات، أطباء، مختبرات — أعلى مردود', PRIMARY),
    ('نقاط الحضور المادي', 'ملصقات، أكياس تسليم موحّدة، رموز QR', ACCENT),
    ('السائقون كوسيلة متحرّكة', 'أزياء موحّدة + لافتات — إعلان مجاني متواصل', TEAL_DARK),
    ('التوعية المجتمعية', 'حملات بأضرار الأدوية المزوّرة + أيام فحص', PRIMARY),
    ('المناسبات والأجندة', 'مواسم الذروة والأعياد + حفل إطلاق إعلامي', ACCENT),
]
for i, (t, d, c) in enumerate(items):
    row = i // 3
    col = i % 3
    x = Inches(0.6) + col * Inches(4.2)
    y = Inches(2.15) + row * Inches(2.35)
    add_rect(s, x, y, Inches(3.9), Inches(2.05), fill=WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.06)
    add_rect(s, x, y, Inches(0.16), Inches(2.05), fill=c)
    add_text(s, x + Inches(0.4), y + Inches(0.28), Inches(3.3), Inches(0.6),
             [{'text': t, 'size': 17, 'bold': True, 'color': DARK}], align=PP_ALIGN.RIGHT)
    add_text(s, x + Inches(0.4), y + Inches(1.0), Inches(3.3), Inches(0.95),
             [{'text': d, 'size': 13, 'color': MUTED, 'rtl': True, 'line_spacing': 1.25}], align=PP_ALIGN.RIGHT)
add_footer(s, 21)

# ================== 22 — الشراكات الصحية ==================
s = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(s)
add_title(s, 'الأوفلاين: الشراكات الصحية — أعلى مردود', kicker='Health Partnerships')
banner(s, Inches(1.95), Inches(0.95), '«صوّر وصفتك، وأدويةُ توصلك إلى بيتك» — رمز QR في استقبال العيادة وصالة الانتظار.',
       fill=PRIMARY, size=16, lead='الفكرة المحورية: ', lead_color=ACCENT)
bullets(s, Inches(0.7), Inches(3.3), Inches(11.9),
        ['العيادات والمستشفيات والأطباء — رمز QR عند الاستقبال (باتفاق مع العيادة، مع مراعاة سرّيّة بيانات المريض).',
         'أقسام التصوير والمختبرات — «استلم نتيجتك ووصفتك من مكانك» — يستغلّ زيارةٍ هي في ذروة الحاجة.',
         'جمعيات خيرية — حملة «صلة خير» توصّل الدواء للمحتاجين: قصة إنسانية تُغذي «القلب» من العلامة.',
         'الامتثال الحاكم: اتفاقيات وصول + خصوصية + لا نصيحة طبية من العلامة.'],
        size=15, gap=0.8, h=0.8)
add_footer(s, 22)

# ================== 23 — حضور مادي + سائق ==================
s = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(s)
add_title(s, 'الأوفلاين: حضورٌ مادي + السائق كوسيلة إعلانية متحرّكة', kicker='Guerrilla', accent_color=ACCENT)
banner(s, Inches(1.95), Inches(0.9), 'السائق يتجوّل في كل نطاق يومياً — إعلانٌ مجانيٌّ متواصل يصل إلى كل حيّ.',
       fill=ACCENT, tcolor=DARK, size=16.5)
bullets(s, Inches(0.7), Inches(3.2), Inches(11.9),
        ['أزياء موحّدة + لافتات على السيارات/الدراجات تحمل «صِلَة دوائي — دواؤك يصلك».',
         'أكياس/شكائر ورقية موحّدة يُسلَّم بها الطرد + بطاقة ترحيب ورمز QR — تنتقل عبر الأحياء يداً بيد.',
         'ملصقات/بوسترات في الصيدليات الشريكة ومراكز الخدمة والمناطق عالية الطلب.',
         'كود دعوة السائق + برنامج «سائق يجلب سائقاً» — توسّع أسطولٍ يبني نفسه بنفسه.'],
        size=15, gap=0.8, h=0.8, dot=ACCENT)
add_footer(s, 23)

# ================== 24 — توعية + مناسبات ==================
s = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(s)
add_title(s, 'الأوفلاين: التوعية المجتمعية والمناسبات', kicker='Brand Purpose & Timing')
card(s, Inches(0.6), Inches(2.0), Inches(6.0), Inches(4.6), 'التوعية والمجتمع', TEAL_DARK,
     ['حملات بأضرار الأدوية المزوّرة في المدارس والجامعات والمراكز المجتمعية.',
      'بشراكة مع نقابة الصيادلة أو مؤسسات صحية — تُبنى الثقة وعلامة «المسؤول».',
      'أيام فحص/تثقيف صحي برعاية «صِلَة» — اسمٌ حاضرٌ في سياقٍ صحي موثوق.'],
     size=14, gap=0.9, item_h=1.0)
card(s, Inches(6.9), Inches(2.0), Inches(5.8), Inches(4.6), 'التوقيت والأجندة والإطلاق', PRIMARY,
     ['مواسم ذروة الأمراض (تغيّر الطقس، الإنفلونزا): «تجهّز — دواؤك يصلك بسرعة».',
      'مناسبات دينية وأعياد ورمضان: حملات رحيمية «لا تترك دواؤك ينتظر» + عروض ولاء (بمراجعة ثقافية/ناطق).',
      'حفل إطلاق بحضور الصيادلة والشركاء + تغطية إعلامية محلية (PR).',
      'دعوة الصحافة/المؤثّرين لتجربة أمرٍ حقيقي — سردٌ يقود محتوى أصيلاً.'],
     size=14, gap=0.9, item_h=1.0)
add_footer(s, 24)

# ================== 25 — آليات النمو ==================
s = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(s)
add_title(s, 'آليات النمو والتشغيل', kicker='Growth Mechanics')
rows = [
    ('إحالة + ولاء', 'خفض CAC وتكرار', 'الدواء شأن عائلي — قياس الأثر بلا إفراط في الخصم'),
    ('واتساب / تيليغرام', 'تقليل الاحتكاك + خدمة iOS', 'إلغاء اشتراك وخصوصية — فصل الخدمة عن التسويق'),
    ('B2B ميداني (Outbound)', 'ملء العرض', 'فريق + Demo + CSM — ابدأ بالصيدليات ذات المرجعية'),
    ('توظيف سائقين', 'إشباع الطلب', 'برامج إحالة + شفافية الأرباح'),
    ('شراكات عيادات/أطباء', 'إحالة عالية الثقة', 'خصوصية المريض — اتفاقيات وصلاحيات'),
    ('PR / المؤثّرون', 'بناء الثقة', 'إفصاح + مراجعة ادّعاءات'),
    ('ASO / SEO', 'نموّ غير مكلّف', 'تحسين المتجر والمحتوى العربي'),
]
col_x = [Inches(0.6), Inches(4.4), Inches(7.6)]
col_w = [Inches(3.8), Inches(3.2), Inches(5.1)]
table(s, Inches(1.95), ('الآلية', 'الأثر', 'ملاحظة تشغيلية'), rows, col_x, col_w,
      row_h=0.58, hdr_h=0.54, body_size=12.5,
      aligns=[PP_ALIGN.RIGHT, PP_ALIGN.RIGHT, PP_ALIGN.RIGHT])
add_footer(s, 25)

# ================== 26 — KPIs ==================
s = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(s)
add_title(s, 'مؤشّرات الأداء (KPIs)', kicker='Measurement')
banner(s, Inches(1.9), Inches(1.0), 'عدد الأوامر المكتملة المدفوعة / أسبوع — مؤشّر النجاح الأصلي، لا عدد التنزيلات.',
       fill=DARK, size=16, lead='المؤشّر الأساسي (North-Star): ', lead_color=ACCENT)
cats = [
    ('الاستحواذ', 'تنزيل، تثبيت، CAC'),
    ('رحلة الأمر', 'بحث/رفع، عروض، تأكيد، تسليم، دفع'),
    ('جودة السوق', 'نسبة أوامر حصلت على عرض، ردّ الصيدلية، تسليم ناجح'),
    ('الاحتفاظ', 'تكرار الأمر خلال ٣٠ يوماً، أوامر/عميل'),
    ('الاقتصاد', 'AOV، تحويل السلة، LTV/CAC، تراجع COD'),
    ('العرض', 'صيدليات نشطة، سائقون، زمن/تكلفة التوصيل'),
    ('التجربة', 'تقييم المتجر، NPS، معدل الشكاوى'),
]
for i, (t, d) in enumerate(cats):
    row = i // 3
    col = i % 3
    x = Inches(0.6) + col * Inches(4.2)
    y = Inches(3.3) + row * Inches(1.55)
    add_rect(s, x, y, Inches(3.9), Inches(1.4), fill=WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.08)
    add_text(s, x + Inches(0.3), y + Inches(0.16), Inches(3.3), Inches(0.5),
             [{'text': t, 'size': 15.5, 'bold': True, 'color': PRIMARY}], align=PP_ALIGN.RIGHT)
    add_text(s, x + Inches(0.3), y + Inches(0.68), Inches(3.3), Inches(0.65),
             [{'text': d, 'size': 11.5, 'color': MUTED, 'rtl': True, 'line_spacing': 1.15}], align=PP_ALIGN.RIGHT)
add_footer(s, 26)

# ================== 27 — 7Ps ==================
s = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(s)
add_title(s, 'مزيج التسويق (7Ps) — خلاصة', kicker='Marketing Mix')
rows = [
    ('المنتج', 'التوفّر الفعلي على الواجهة، رفع الوصفة بلا احتكاك، إثبات الثقة، إرشاد صيدلاني بحدود تنظيمية'),
    ('السعر', 'بلا رسوم خفية؛ عمولة واضحة للصيدلية؛ «عرض حارس» مقيّد'),
    ('التوزيع', 'نطاق مكثّف أولاً؛ تطبيق أندرويد/آيفون + واتساب + قنوات غير متجَر'),
    ('الترويج', 'ثقة + تعليم + إحالة + محتوى + شراكات — بلا حملات مدفوعة واسعة في البداية'),
    ('الأشخاص', 'التسليم = الوجه؛ تدريب السائقين على المهنيّة والسرّيّة وتقييمهم'),
    ('العمليات', 'سلاسة من الوصفة إلى الباب؛ التحقق النقدي؛ إدارة الغياب/تراجع الزبون'),
    ('الإثبات المادي', 'شارة التوثيق، الأزياء الموحّدة، رموز QR، هُوية عراقية محبوبة'),
]
col_x = [Inches(0.6), Inches(3.4)]
col_w = [Inches(2.8), Inches(9.3)]
table(s, Inches(1.95), ('العنصر', 'القرار'), rows, col_x, col_w,
      row_h=0.58, hdr_h=0.54, body_size=13,
      aligns=[PP_ALIGN.RIGHT, PP_ALIGN.RIGHT])
add_footer(s, 27)

# ================== 28 — خارطة 12 شهراً ==================
s = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(s)
add_title(s, 'خارطة الطريق — 12 شهراً', kicker='Roadmap')
stages = [
    ('المرحلة ١ — الإثبات', '0–3 أشهر',
     ['سيولة نطاقٍ واحد (بغداد)', 'قياس كامل + مؤشّر أصلي', 'محتوى ثقة + إحالة'], PRIMARY),
    ('المرحلة ٢ — النمو', '3–9 أشهر',
     ['توسّع جغرافي محوّلي', 'تصعيد TikTok/Instagram + ASO', 'إطلاق iOS + مضاعفة العرض'], ACCENT),
    ('المرحلة ٣ — التوسّع', '9–18 شهراً',
     ['مدن جديدة: البصرة، أربيل، الموصل', 'برنامج ولاء + شراكات عيادات/مستشفيات', 'ربحية الوحدة (Unit Economics)'], TEAL_DARK),
]
x = Inches(0.6)
for t, dur, items_, color in stages:
    add_rect(s, x, Inches(2.1), Inches(3.9), Inches(4.4), fill=WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.05)
    add_rect(s, x, Inches(2.1), Inches(3.9), Inches(1.05), fill=color)
    add_text(s, x + Inches(0.25), Inches(2.26), Inches(3.4), Inches(0.55),
             [{'text': t, 'size': 19, 'bold': True, 'color': WHITE}], align=PP_ALIGN.RIGHT)
    add_text(s, x + Inches(0.25), Inches(2.78), Inches(3.4), Inches(0.4),
             [{'text': dur, 'size': 13, 'bold': True, 'color': ICE}], align=PP_ALIGN.RIGHT)
    yy = Inches(3.5)
    for it in items_:
        add_rect(s, x + Inches(0.35), yy + Inches(0.08), Inches(0.2), Inches(0.2), fill=color, shape=MSO_SHAPE.OVAL)
        add_text(s, x + Inches(0.75), yy - Inches(0.02), Inches(2.95), Inches(0.75),
                 [{'text': it, 'size': 13.5, 'color': MUTED, 'rtl': True, 'line_spacing': 1.15}], align=PP_ALIGN.RIGHT)
        yy += Inches(0.95)
    x = x + Inches(4.2)
add_footer(s, 28)

# ================== 29 — الميزانية ==================
s = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(s)
add_title(s, 'الميزانية — سيناريوهات تقريبية', kicker='Budget  •  فرضيات تخطيطية')
# يمين: التوزيع مع أشرطة
data = [('محتوى + بصرية', 20), ('إعلانات أداء (ميتا/تيك توك/إنستا)', 35),
        ('تحفيز (أول أمر/إحالة/ولاء)', 15), ('فريق B2B ميداني + عرض', 15),
        ('توظيف سائقين + تحفيز', 10), ('أدوات/قياس/تسويق رقمي', 5)]
y = Inches(2.0)
add_text(s, Inches(7.3), y, Inches(5.4), Inches(0.4),
         [{'text': 'توزيع الإنفاق', 'size': 16, 'bold': True, 'color': DARK, 'rtl': True}], align=PP_ALIGN.RIGHT)
y += Inches(0.55)
for i, (lab, pct) in enumerate(data):
    add_text(s, Inches(7.3), y, Inches(4.3), Inches(0.35),
             [{'text': lab, 'size': 12.5, 'color': MUTED, 'rtl': True}], align=PP_ALIGN.RIGHT)
    # شريط: أقصى عرض 3.2in عند 35%
    barw = 3.2 * pct / 35.0
    add_rect(s, Inches(9.9) - Inches(barw), y + Inches(0.34), Inches(barw), Inches(0.16),
             fill=PRIMARY if i != 1 else ACCENT, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.5)
    add_text(s, Inches(7.3), y + Inches(0.28), Inches(2.4), Inches(0.35),
             [{'text': f'{pct}%', 'size': 12.5, 'bold': True, 'color': PRIMARY}], align=PP_ALIGN.LEFT, rtl=False)
    y += Inches(0.62)
# يسار: السيناريوان
add_rect(s, Inches(0.6), Inches(2.0), Inches(6.2), Inches(2.1), fill=WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.05)
add_rect(s, Inches(0.6), Inches(2.0), Inches(6.2), Inches(0.7), fill=PRIMARY)
add_text(s, Inches(0.9), Inches(2.12), Inches(5.6), Inches(0.5),
         [{'text': 'سيناريو بوتستراب (ميزانية محدودة)', 'size': 15.5, 'bold': True, 'color': WHITE}], align=PP_ALIGN.RIGHT)
add_text(s, Inches(0.9), Inches(2.9), Inches(5.6), Inches(1.1),
         [{'text': 'إحالة + واتساب + محتوى + ASO + شراكات — بهدف الإثبات فقط؛ مدفوعٌ محدود جداً.', 'size': 13, 'color': MUTED, 'rtl': True, 'line_spacing': 1.3}],
         align=PP_ALIGN.RIGHT)
add_rect(s, Inches(0.6), Inches(4.3), Inches(6.2), Inches(2.1), fill=WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.05)
add_rect(s, Inches(0.6), Inches(4.3), Inches(6.2), Inches(0.7), fill=TEAL_DARK)
add_text(s, Inches(0.9), Inches(4.42), Inches(5.6), Inches(0.5),
         [{'text': 'سيناريو النمو', 'size': 15.5, 'bold': True, 'color': WHITE}], align=PP_ALIGN.RIGHT)
add_text(s, Inches(0.9), Inches(5.2), Inches(5.6), Inches(1.1),
         [{'text': 'تصعيد TikTok/Instagram + ظهور iOS + شراكات عيادات + توسّع محوّلي بعد تحقيق بوابات المرحلة.', 'size': 13, 'color': MUTED, 'rtl': True, 'line_spacing': 1.3}],
         align=PP_ALIGN.RIGHT)
banner(s, Inches(6.65), Inches(0.85), 'لا تتجاوز الإنفاق المدفوع قبل أن يتحقّق CAC ≤ LTV ونِسبة التوفّر ≥ الهدف — الإنفاق يتبع الأداء لا العكس.',
       fill=DARK, size=14, lead='القاعدة الذهبية: ', lead_color=ACCENT)
add_footer(s, 29)

# ================== 30 — المخاطر والامتثال ==================
s = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(s)
add_title(s, 'إدارة المخاطر والامتثال', kicker='Governance', accent_color=ACCENT)
risks = [
    ('ازدحام جغرافي (سيولة ضعيفة)', 'الكثافة أولاً؛ رصد Fill Rate باستمرار'),
    ('قانوني / تنظيمي (أدوية)', 'استشارة قانونية وصيدلانية؛ توثيق؛ بلا ادّعاءات'),
    ('الادّعاءات الصحية / النصائح', 'مراجعة قبل النشر؛ منع الجرعات والاستبدالات'),
    ('الخصوصية والوصفات', 'تشفير، تقليل البيانات، سياسة، موافقة، مسؤول حماية'),
    ('جودة التسليم / تخلّف COD', 'رمز استلام، توثيق، مكافأة سائق، إدارة الغياب'),
    ('التبنّي الرقمي المنخفض', 'تعليم + ثقة + إتاحة الأمر عبر واتساب'),
]
for i, (t, d) in enumerate(risks):
    row = i // 3
    col = i % 3
    x = Inches(0.6) + col * Inches(4.2)
    y = Inches(1.95) + row * Inches(1.95)
    add_rect(s, x, y, Inches(3.9), Inches(1.8), fill=WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.06)
    add_rect(s, x, y, Inches(0.14), Inches(1.8), fill=ACCENT)
    add_text(s, x + Inches(0.35), y + Inches(0.2), Inches(3.3), Inches(0.6),
             [{'text': t, 'size': 14.5, 'bold': True, 'color': DARK, 'rtl': True, 'line_spacing': 1.1}], align=PP_ALIGN.RIGHT)
    add_text(s, x + Inches(0.35), y + Inches(0.95), Inches(3.3), Inches(0.75),
             [{'text': d, 'size': 12.5, 'color': MUTED, 'rtl': True, 'line_spacing': 1.2}], align=PP_ALIGN.RIGHT)
add_rect(s, Inches(0.6), Inches(6.0), Inches(12.1), Inches(0.85), fill=PRIMARY, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.15)
add_text(s, Inches(0.9), Inches(6.13), Inches(11.6), Inches(0.6),
         [{'segments': [{'text': 'بوابات الإطلاق: ', 'size': 14.5, 'bold': True, 'color': ACCENT, 'rtl': True},
                        {'text': 'تتبّع كامل  •  توفّرٌ فعلي في النطاق  •  مراجعة صيدلانية/قانونية/ناطق  •  CAC ≤ LTV قبل أي حملة واسعة', 'size': 14, 'color': WHITE, 'rtl': True}]}],
         align=PP_ALIGN.RIGHT)
add_footer(s, 30)

# ================== 31 — الخلاصة ==================
s = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(s, DARK)
add_rect(s, 0, 0, SW, Inches(0.35), fill=ACCENT)
add_rect(s, SW - Inches(3.0), Inches(4.6), Inches(2.6), Inches(2.6), fill=TEAL_DARK, shape=MSO_SHAPE.OVAL)
add_text(s, Inches(0.6), Inches(1.2), Inches(12.1), Inches(1.2),
         [{'text': 'المعركة معركة ثقة، لا معركة أسعار.', 'size': 36, 'bold': True, 'color': WHITE, 'rtl': True, 'line_spacing': 1.1}],
         align=PP_ALIGN.RIGHT)
add_rect(s, Inches(0.6), Inches(2.7), Inches(3.5), Inches(0.10), fill=ACCENT)
add_text(s, Inches(0.6), Inches(3.15), Inches(11.9), Inches(2.0),
         [{'text': 'قِس أولاً، ثم اكسب السيولة في نطاقٍ واحد، ثم قدّم نفسك «القناة القانونية المعتمدة» التي تُنقذ العراقي من فوضى الصفحات غير الموثوقة — وابنِ محرّك نموّ مستداماً: إحالة + واتساب + محتوى + ASO + شراكات.', 'size': 19, 'color': ICE, 'rtl': True, 'line_spacing': 1.45}],
         align=PP_ALIGN.RIGHT)
add_text(s, Inches(0.6), Inches(5.35), Inches(12.1), Inches(0.9),
         [{'segments': [{'text': 'نبدأ من أول ٣٠ يوماً:  ', 'size': 17, 'bold': True, 'color': ACCENT, 'rtl': True},
                        {'text': 'قياس  →  سيولة  →  ثقة  →  نموّ', 'size': 17, 'color': WHITE, 'rtl': True}]}],
         align=PP_ALIGN.RIGHT)
add_text(s, Inches(0.6), Inches(6.4), Inches(12.1), Inches(0.9),
         [{'segments': [{'text': 'صِـلَة دوائي  •  ', 'size': 18, 'bold': True, 'color': ACCENT, 'rtl': True},
                        {'text': 'دواؤك… يصلك من صيدلية مرخّصة قريبة', 'size': 18, 'color': WHITE, 'rtl': True}]}],
         align=PP_ALIGN.RIGHT)
add_footer(s, 31)

out = '/home/user/Skills/docs/Sillaiq_Presentation_v2.pptx'
prs.save(out)
print('SAVED:', out)
print('SLIDES:', len(prs.slides._sldIdLst))
assert len(prs.slides._sldIdLst) == TOTAL, 'slide count mismatch'
