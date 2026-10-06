#!/usr/bin/env python3
"""Build xingbao_village_v6.pptx: topic-first, large images, embedded slide animations."""
import os, sys
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

OUT = sys.argv[1] if len(sys.argv) > 1 else 'xingbao_village_v6.pptx'
IMG = '/home/geeker/.hermes/cache/scratch/xingbao_all'
QR = '/home/geeker/.hermes/cache/scratch/xingbao_v3_imgs'

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
W, H = 13.333, 7.5

BLACK = RGBColor(8,10,12)
DARK = RGBColor(15,17,20)
GOLD = RGBColor(198,166,94)
CREAM = RGBColor(244,236,220)
WHITE = RGBColor(255,255,255)
MUTED = RGBColor(190,196,205)
RED = RGBColor(174,54,54)

# ---- helpers ----
def bg(slide, color=DARK):
    slide.background.fill.solid(); slide.background.fill.fore_color.rgb = color

def pic(slide, name, l, t, w, h):
    p = os.path.join(IMG, name)
    if os.path.exists(p) and os.path.getsize(p) > 10000:
        return slide.shapes.add_picture(p, Inches(l), Inches(t), Inches(w), Inches(h))
    return None

def cover_pic(slide, name):
    p = os.path.join(IMG, name)
    if os.path.exists(p) and os.path.getsize(p) > 10000:
        return slide.shapes.add_picture(p, Inches(0), Inches(0), Inches(W), Inches(H))
    return None

def rect(slide, l, t, w, h, color, alpha=None, shape=MSO_SHAPE.RECTANGLE):
    sh = slide.shapes.add_shape(shape, Inches(l), Inches(t), Inches(w), Inches(h))
    sh.fill.solid(); sh.fill.fore_color.rgb = color
    if alpha is not None:
        sf = sh.fill._xPr.find('{http://schemas.openxmlformats.org/drawingml/2006/main}solidFill')
        srgb = sf.find('{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr')
        a = srgb.makeelement('{http://schemas.openxmlformats.org/drawingml/2006/main}alpha', {'val': str(int(alpha*100000))})
        srgb.append(a)
    sh.line.fill.background()
    return sh

def tb(slide, text, l, t, w, h, size=18, color=WHITE, bold=False, align=PP_ALIGN.LEFT, font='Microsoft YaHei'):
    box = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = box.text_frame; tf.clear(); tf.word_wrap = True
    p = tf.paragraphs[0]; p.alignment = align
    r = p.add_run(); r.text = text
    r.font.name = font; r.font.size = Pt(size); r.font.bold = bold; r.font.color.rgb = color
    return box

def multi(slide, lines, l, t, w, h, size=16, color=CREAM, gap=6, bold=False):
    box = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = box.text_frame; tf.clear(); tf.word_wrap = True
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(gap)
        r = p.add_run(); r.text = line
        r.font.name = 'Microsoft YaHei'; r.font.size = Pt(size); r.font.bold = bold; r.font.color.rgb = color
    return box

def gold_bar(slide, x, y, w=0.75, h=0.05):
    return rect(slide, x, y, w, h, GOLD)

def chip(slide, text, x, y, w=1.2, h=0.32, fill=RED):
    s = rect(slide, x, y, w, h, fill, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    s.text_frame.text = text
    s.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    r = s.text_frame.paragraphs[0].runs[0]
    r.font.size = Pt(8); r.font.bold = True; r.font.color.rgb = WHITE; r.font.name='Microsoft YaHei'
    return s

def footer(slide, n, title='邢堡村 · 黄土高原上的千年村落'):
    tb(slide, title, 0.7, 7.04, 6.8, 0.24, 7, RGBColor(140,140,140))
    tb(slide, str(n).zfill(2), 12.3, 7.04, 0.4, 0.24, 7, GOLD, True, PP_ALIGN.RIGHT)

def animate(slide, shapes, effect='fade'):
    """Add PowerPoint animation XML: fade/float by default.
    shapes: list of shape objects."""
    if not shapes: return
    import copy
    from pptx.oxml.ns import qn
    from lxml import etree
    xml = '''<p:timing xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>
        <p:seq concurrent="1" nextAc="seek"><p:cTn id="2" restart="whenNotActive" fill="hold" nodeType="mainSeq">
          <p:prevCondLst><p:cond evt="onPrev" delay="indefinite"><p:action evt="onPrev" delay="0"><p:tn evt="onPrev" delay="0"/></p:action></p:cond></p:prevCondLst>
          <p:nextCondLst><p:cond evt="onNext" delay="indefinite"><p:action evt="onNext" delay="0"><p:tn evt="onNext" delay="0"/></p:action></p:cond></p:nextCondLst>
          <p:childTnLst>
        </p:childTnLst>
        </p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:action evt="onPrev" delay="0"><p:prev></p:prev></p:action></p:cond></p:prevCondLst><p:nextCondLst><p:cond evt="onNext" delay="0"><p:action evt="onNext" delay="0"><p:next></p:next></p:action></p:cond></p:nextCondLst></p:seq>
      </p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>'''
    root = etree.fromstring(xml)
    mainseq = [el for el in root.iter(qn('p:cTn')) if el.get('nodeType') == 'mainSeq']
    if not mainseq:
        raise RuntimeError('mainSeq cTn not found')
    mainseq = mainseq[0]
    sid = 10
    for i, sh in enumerate(shapes[:14]):
        sid += 1; sid2 = sid + 1
        shape_id = sh.shape_id
        delay = i * 250
        dur = 700 if effect == 'fade' else 1000
        anim = f'''<p:par xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cTn id="{sid}" presetID="10" presetClass="mediacall" presetSubtype="0" fill="hold" nodeType="clickEffect"><p:stCondLst><p:cond delay="0"><p:tn val="{delay}"/></p:cond></p:stCondLst><p:childTnLst><p:par><p:cTn id="{sid2}" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" grpId="0" nodeType="withEffect"><p:stCondLst><p:cond delay="0"><p:tn val="{delay}"/></p:cond></p:stCondLst><p:childTnLst><p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="{sid2+1}" dur="{dur}" fill="hold"><p:stCondLst><p:cond delay="0"><p:tn val="{delay}"/></p:cond></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="{shape_id}"/></p:tgtEl><p:attrNameLst><p:attrName>style.opacity</p:attrName></p:attrNameLst></p:cBhvr><p:progress></p:progress><p:ramp color="ffffff"/></p:animEffect></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>'''
        mainseq.find(qn('p:childTnLst')).append(etree.fromstring(anim))
    existing = slide._element.find(qn('p:timing'))
    if existing is not None: slide._element.remove(existing)
    slide._element.append(root)

def qr_page(n, title, items):
    slide = prs.slides.add_slide(prs.slide_layouts[6]); bg(slide)
    pic(slide, 'old_village.jpg', 0, 0, W, H)
    rect(slide, 0, 0, W, H, BLACK, 0.72)
    tb(slide, title, 0.85, 0.72, 8.5, 0.55, 28, CREAM, True)
    gold_bar(slide, 0.9, 1.43, 1.0)
    x = 1.15
    for label, file, url in items:
        qf = os.path.join(QR, file)
        if os.path.exists(qf):
            slide.shapes.add_picture(qf, Inches(x), Inches(2.3), Inches(1.55), Inches(1.55))
            tb(slide, label, x, 4.02, 1.55, 0.26, 10, GOLD, True, PP_ALIGN.CENTER)
            tb(slide, url, x-0.18, 4.32, 1.9, 0.55, 6, MUTED, False, PP_ALIGN.CENTER)
        x += 2.25
    multi(slide, ['扫码查看抖音/视频号/小红书相关内容', '适合用于线下活动、文旅招商、村庄导览推广'], 0.95, 5.55, 8.0, 0.65, 16, CREAM)
    footer(slide, n); animate(slide, [slide.shapes[-1], slide.shapes[-2], slide.shapes[-3], slide.shapes[-4], slide.shapes[-5], slide.shapes[-6]])
    return slide

# ---------- slides ----------
# 1 Cover
s = prs.slides.add_slide(prs.slide_layouts[6]); bg(s)
cover_pic(s, 'chinese_village.jpg')
rect(s, 0, 0, W, H, BLACK, 0.45)
rect(s, 0, 4.1, W, 3.4, BLACK, 0.62)
chip(s, 'STORY · 2026', 0.85, 0.78, 1.55, 0.34)
tb(s, '邢堡村', 0.82, 1.85, 7.5, 1.1, 62, WHITE, True, font='SimSun')
gold_bar(s, 0.9, 3.08, 1.1)
tb(s, '化峪镇 · 稷山县 · 运城', 0.9, 3.32, 6.5, 0.35, 18, GOLD, True)
multi(s, ['黄土塬上的农耕村落', '从后稷故里到网红打卡的山西乡村叙事'], 0.92, 4.62, 8.2, 1.0, 20, CREAM)
tb(s, '26 页 · 专题图集 · 扫码视频入口', 9.1, 6.86, 3.5, 0.28, 8, RGBColor(180,180,180), False, PP_ALIGN.RIGHT)
animate(s, list(list(s.shapes))[:14])

# 2 TOC
s = prs.slides.add_slide(prs.slide_layouts[6]); bg(s)
tb(s, '目录 CONTENTS', 0.85, 0.72, 5.5, 0.52, 30, WHITE, True)
gold_bar(s, 0.9, 1.42)
items = ['01 地理概貌', '02 历史溯源', '03 古村落建筑', '04 农耕文明', '05 稷山四宝', '06 古迹与非遗', '07 短视频推广']
for i,it in enumerate(items):
    x=0.9+(i%2)*6.2; y=2.1+(i//2)*0.72
    tb(s, it.split()[0], x, y, 0.6, 0.3, 14, GOLD, True)
    tb(s, it.split()[1], x+0.75, y, 3.5, 0.3, 15, CREAM, True)
footer(s, 2); animate(s, list(list(s.shapes))[:14])

# 3 Geography chapter
s = prs.slides.add_slide(prs.slide_layouts[6]); bg(s)
cover_pic(s, 'loess_plateau.jpg'); rect(s, 0, 0, W, H, BLACK, 0.66)
tb(s, '01', 0.9, 2.1, 1.4, 0.6, 38, GOLD, True)
tb(s, '地理概貌', 0.9, 2.9, 6.2, 0.7, 44, WHITE, True, font='SimSun')
tb(s, 'Loess Plateau Landscape', 0.92, 3.75, 6.0, 0.3, 13, GOLD, True)
tb(s, '邢堡村位于运城盆地东缘，黄土塬与塬缘坡地交织，塬面平坦、沟壑纵横，是晋南典型农耕聚落地貌。', 0.92, 4.2, 7.0, 0.85, 18, CREAM)
footer(s, 3); animate(s, list(list(s.shapes))[:14])

# 4 Geo detail 3 large images
s = prs.slides.add_slide(prs.slide_layouts[6]); bg(s)
tb(s, '塬 · 坡 · 沟', 0.85, 0.58, 5.5, 0.48, 28, WHITE, True)
gold_bar(s, 0.9, 1.25, 0.9)
for idx,(name,cap,l,t,w,h) in enumerate([
    ('loess_plateau.jpg','塬面','0.62','1.85','3.85','3.15'),
    ('village_roof.jpg','塬坡民居','4.78','1.85','3.85','3.15'),
    ('terraced_fields.jpg','沟坡耕地','8.94','1.85','3.85','3.15')]):
    pic(s, name, float(l), float(t), float(w), float(h))
    rect(s, float(l), float(t)+float(h)-0.58, float(w), 0.58, BLACK, 0.72)
    tb(s, cap, float(l)+0.2, float(t)+float(h)-0.46, float(w)-0.4, 0.28, 13, WHITE, True)
multi(s, ['晋南塬区地貌兼具耕作用地与宜居院落条件；黄土厚积、土层肥沃，也塑造了窑洞、四合院、场院、晒台等生活方式。'], 0.92, 5.8, 11.5, 0.7, 16, CREAM)
footer(s, 4); animate(s, [s.shapes[1], s.shapes[2]] + list(list(s.shapes)[3:]))

# 5 History
s = prs.slides.add_slide(prs.slide_layouts[6]); bg(s)
cover_pic(s, 'china_temple.jpg'); rect(s, 0, 0, W, H, BLACK, 0.70)
tb(s, '02 历史溯源', 0.9, 0.95, 6.0, 0.55, 34, WHITE, True)
gold_bar(s, 0.92, 1.72, 0.9)
multi(s, ['稷山因后稷教稼得名，是中华农耕文明的重要记忆现场。', '邢堡村延续塬区村落格局，保留古井、场院、老槐、民居院墙与地方记忆。', '村落不是孤立建筑，而是耕作、节庆、信仰、宗族、手艺共同沉淀的生活系统。'], 0.92, 2.35, 8.4, 1.8, 19, CREAM, 10)
rect(s, 8.75, 2.1, 3.75, 2.45, BLACK, 0.36)
tb(s, '后稷 · 稼穑', 9.05, 2.45, 2.6, 0.4, 20, GOLD, True)
tb(s, '教民稼穑，百谷用登。', 9.05, 3.05, 2.8, 0.55, 15, CREAM)
footer(s, 5); animate(s, list(list(s.shapes))[:14])

# 6 History archive
s = prs.slides.add_slide(prs.slide_layouts[6]); bg(s)
tb(s, '村落记忆', 0.85, 0.62, 5.0, 0.5, 30, WHITE, True)
gold_bar(s, 0.9, 1.32)
cards=[('600年古槐','村庄公共记忆与节庆中心'),('老院墙','青砖与夯土的晋南民居肌理'),('古井场院','饮水、晾晒、邻里交往空间'),('农耕时节','谷雨、立夏、秋收、冬藏')]
for i,(a,b) in enumerate(cards):
    x=0.85+i*3.05
    pic(s, 'chinese_architecture.jpg' if i%2==0 else 'old_village.jpg', x, 2.05, 2.65, 3.1)
    rect(s, x, 4.75, 2.65, 0.95, BLACK, 0.72)
    tb(s, a, x+0.18, 4.9, 2.3, 0.24, 13, GOLD, True)
    tb(s, b, x+0.18, 5.2, 2.3, 0.36, 9, CREAM)
footer(s, 6); animate(s, [s.shapes[1], s.shapes[2]] + list(list(s.shapes)[3:]))

# 7 Architecture chapter
s = prs.slides.add_slide(prs.slide_layouts[6]); bg(s)
cover_pic(s, 'chinese_architecture.jpg'); rect(s, 0, 0, W, H, BLACK, 0.68)
tb(s, '03 古村落建筑', 0.9, 2.05, 7.2, 0.7, 44, WHITE, True, font='SimSun')
gold_bar(s, 0.92, 2.92, 1.05)
tb(s, '青砖、灰瓦、院落、砖雕：晋南民居的材料语言。', 0.92, 3.32, 8.3, 0.45, 20, CREAM)
footer(s, 7); animate(s, [s.shapes[2], s.shapes[3], s.shapes[4]])

# 8 Architecture large detail
s = prs.slides.add_slide(prs.slide_layouts[6]); bg(s)
cover_pic(s, 'china_temple.jpg')
rect(s, 0, 0, 5.25, H, BLACK, 0.76)
rect(s, 5.25, 0, 8.08, H, BLACK, 0.28)
tb(s, '青砖古韵', 0.85, 2.25, 4.0, 0.55, 34, WHITE, True, font='SimSun')
gold_bar(s, 0.9, 3.0, 0.95)
multi(s, ['晋南建筑擅长在有限空间内营造庄重秩序。', '青砖墀头、门楼砖雕、瓦当脊饰不仅是装饰，也是身份、技艺和祈愿的符号。'], 0.9, 3.32, 3.85, 1.2, 17, CREAM, 8)
chip(s, '砖雕 / 青砖 / 院落', 8.35, 0.62, 2.55, 0.34)
tb(s, '邢堡民居可提炼为：一进院、一方院、一台晒、一堵墙。', 8.05, 5.95, 4.2, 0.5, 14, WHITE)
footer(s, 8); animate(s, list(list(s.shapes))[:14])

# 9 Architecture gallery
s = prs.slides.add_slide(prs.slide_layouts[6]); bg(s)
tb(s, '建筑图录', 0.85, 0.58, 5.0, 0.5, 30, WHITE, True)
gold_bar(s, 0.9, 1.28)
pics=[('chinese_roof.jpg','屋顶与灰瓦'),('china_temple.jpg','院落与门楼'),('village_roof.jpg','塬坡民居'),('china_village2.jpg','村落肌理')]
for i,(name,cap) in enumerate(pics):
    x=0.65+(i%2)*6.35; y=1.78+(i//2)*2.45
    pic(s, name, x, y, 5.65, 2.08)
    tb(s, cap, x+0.22, y+1.72, 4.2, 0.24, 11, GOLD, True)
footer(s, 9); animate(s, [s.shapes[1], s.shapes[2]] + list(list(s.shapes)[3:]))

# 10 Agriculture chapter
s = prs.slides.add_slide(prs.slide_layouts[6]); bg(s)
cover_pic(s, 'wheat_field.jpg'); rect(s, 0, 0, W, H, BLACK, 0.68)
tb(s, '04 农耕文明', 0.9, 2.0, 7.0, 0.7, 44, WHITE, True, font='SimSun')
gold_bar(s, 0.92, 2.88, 1.0)
tb(s, '塬区农作物的四季轮回，也是邢堡村最稳定、最鲜亮的视觉母题。', 0.92, 3.28, 8.6, 0.48, 20, CREAM)
footer(s, 10); animate(s, [s.shapes[2], s.shapes[3], s.shapes[4]])

# 11 Agriculture gallery
s = prs.slides.add_slide(prs.slide_layouts[6]); bg(s)
tb(s, '塬上丰年', 0.85, 0.62, 5.0, 0.5, 30, WHITE, True)
gold_bar(s, 0.9, 1.32)
for i,(name,cap,x,y,w,h) in enumerate([
    ('wheat_field.jpg','麦田',0.65,1.85,3.85,3.25),('harvest.jpg','收获',4.78,1.85,3.85,3.25),('wheat_harvest.jpg','农忙',8.92,1.85,3.85,3.25)]):
    pic(s, name, x, y, w, h)
    rect(s, x, y+h-0.58, w, 0.58, BLACK, 0.72)
    tb(s, cap, x+0.2, y+h-0.46, 2.0, 0.28, 13, WHITE, True)
tb(s, '小麦、玉米、红枣、板枣加工，构成从田间到货架的完整链条。', 0.9, 5.8, 11.5, 0.45, 17, CREAM)
footer(s, 11); animate(s, [s.shapes[1], s.shapes[2]] + list(list(s.shapes)[3:]))

# 12 Four treasures chapter
s = prs.slides.add_slide(prs.slide_layouts[6]); bg(s)
cover_pic(s, 'chinese_food.jpg'); rect(s, 0, 0, W, H, BLACK, 0.72)
tb(s, '05 稷山四宝', 0.9, 1.95, 7.2, 0.72, 44, WHITE, True, font='SimSun')
gold_bar(s, 0.92, 2.85, 1.05)
tb(s, '麻花、板枣与晋南风味，是邢堡村走向全国的文化名片。', 0.92, 3.25, 8.8, 0.5, 20, CREAM)
footer(s, 12); animate(s, [s.shapes[2], s.shapes[3], s.shapes[4]])

# 13 Food gallery
s = prs.slides.add_slide(prs.slide_layouts[6]); bg(s)
tb(s, '美食特写', 0.85, 0.58, 5.0, 0.5, 30, WHITE, True)
gold_bar(s, 0.9, 1.28)
for i,(name,cap,x,y,w,h) in enumerate([
    ('chinese_food.jpg','麻花与面食',0.65,1.78,3.85,2.45),
    ('banzao_1.jpg','板枣',4.78,1.78,3.85,2.45),
    ('banzao_2.jpg','红枣',8.92,1.78,3.85,2.45),
    ('chinese_folk_art.jpg','手艺与节庆',2.75,4.5,7.85,1.55)]):
    pic(s, name, x, y, w, h)
    rect(s, x, y+h-0.46, w, 0.46, BLACK, 0.70)
    tb(s, cap, x+0.2, y+h-0.34, 2.2, 0.22, 11, WHITE, True)
footer(s, 13); animate(s, [s.shapes[1], s.shapes[2]] + list(list(s.shapes)[3:]))

# 14 Mahua process
s = prs.slides.add_slide(prs.slide_layouts[6]); bg(s)
cover_pic(s, 'chinese_food.jpg')
rect(s, 0, 0, W, H, BLACK, 0.55)
tb(s, '稷山麻花 · 非遗手艺', 0.85, 0.78, 6.8, 0.55, 32, WHITE, True)
gold_bar(s, 0.9, 1.52, 1.05)
steps=['和面','醒面','盘条','拧股','入油','金黄出锅']
for i,st in enumerate(steps):
    x=0.88+i*2.0
    rect(s, x, 2.7, 1.45, 1.45, GOLD, 0.18, MSO_SHAPE.OVAL)
    tb(s, f'0{i+1}', x+0.35, 2.93, 0.7, 0.24, 10, GOLD, True, PP_ALIGN.CENTER)
    tb(s, st, x+0.18, 3.28, 1.1, 0.24, 13, WHITE, True, PP_ALIGN.CENTER)
tb(s, '酥脆、耐存、便于传播，是短视频和直播电商的高记忆点产品。', 0.92, 5.5, 9.8, 0.55, 19, CREAM)
footer(s, 14); animate(s, [s.shapes[2], s.shapes[3], s.shapes[4]] + list(list(s.shapes)[5:]))

# 15 Banzao
s = prs.slides.add_slide(prs.slide_layouts[6]); bg(s)
cover_pic(s, 'banzao_1.jpg'); rect(s, 0, 0, W, H, BLACK, 0.58)
rect(s, 0.8, 0.8, 5.3, 5.45, BLACK, 0.72)
tb(s, '稷山板枣', 1.1, 1.55, 4.0, 0.6, 36, WHITE, True, font='SimSun')
gold_bar(s, 1.12, 2.35, 0.95)
multi(s, ['唐代古枣林', '万亩枣花香', '晒制板枣', '农遗记忆'], 1.12, 2.75, 3.8, 1.5, 19, CREAM, 10)
chip(s, '丰收 IP', 9.1, 1.0, 1.35, 0.35)
multi(s, ['以枣花节、采摘节、晒枣场、礼盒伴手礼为核心，可形成四季内容日历。'], 8.25, 4.5, 4.2, 0.9, 18, CREAM)
footer(s, 15); animate(s, list(list(s.shapes))[:14])

# 16 Heritage chapter
s = prs.slides.add_slide(prs.slide_layouts[6]); bg(s)
cover_pic(s, 'big_buddha.jpg'); rect(s, 0, 0, W, H, BLACK, 0.68)
tb(s, '06 古迹与非遗', 0.9, 2.05, 7.5, 0.72, 44, WHITE, True, font='SimSun')
gold_bar(s, 0.92, 2.95, 1.05)
tb(s, '稷王庙、大佛寺、高跷走兽、螺钿漆器，构成区域文旅资源带。', 0.92, 3.38, 9.2, 0.55, 19, CREAM)
footer(s, 16); animate(s, [s.shapes[2], s.shapes[3], s.shapes[4]])

# 17 Heritage list
s = prs.slides.add_slide(prs.slide_layouts[6]); bg(s)
tb(s, '可串联景点', 0.85, 0.6, 5.5, 0.5, 30, WHITE, True)
gold_bar(s, 0.9, 1.3)
data=[('稷王庙','后稷文化，农耕祭祀','big_buddha.jpg'),('大佛寺','千年古刹，晋南宗教建筑','china_temple.jpg'),('马村砖雕墓','砖雕艺术，晋南墓葬文化','chinese_architecture.jpg'),('高跷走兽','民俗表演，节庆传播力','chinese_folk_art.jpg')]
for i,(a,b,name) in enumerate(data):
    x=0.65+(i%2)*6.35; y=1.85+(i//2)*2.25
    pic(s, name, x, y, 5.65, 1.75)
    rect(s, x, y+1.75, 5.65, 0.55, BLACK, 0.78)
    tb(s, a, x+0.2, y+1.86, 2.4, 0.22, 12, GOLD, True)
    tb(s, b, x+2.6, y+1.87, 2.7, 0.22, 8, CREAM)
footer(s, 17); animate(s, [s.shapes[1], s.shapes[2]] + list(list(s.shapes)[3:]))

# 18 QR video
qr_page(18, '07 短视频与扫码传播', [
    ('抖音话题', 'qr_douyin.png', 'douyin.com/search/山西DOU是好风光'),
    ('视频号', 'qr_videochannel.png', 'videos.wechat.qq.com'),
    ('小红书', 'qr_xhs.png', 'xiaohongshu.com/search/稷山'),
    ('麻花百科', 'qr_mahua.png', '1588.tv/techan/3545'),
    ('稷王庙', 'qr_jiwang.png', 'weibo.com/山西文旅')
])

# 19 Video content plan
s = prs.slides.add_slide(prs.slide_layouts[6]); bg(s)
tb(s, '短视频选题库', 0.85, 0.62, 6.0, 0.5, 30, WHITE, True)
gold_bar(s, 0.9, 1.32)
pics=[('wheat_field.jpg','塬上麦浪 15秒'),('banzao_1.jpg','板枣采摘 vlog'),('chinese_food.jpg','麻花出锅慢镜头'),('china_temple.jpg','古院光影走位'),('old_village.jpg','一日邢堡生活')]
for i,(name,cap) in enumerate(pics):
    y=1.85+i*0.95
    pic(s, name, 0.85, y, 1.35, 0.78)
    rect(s, 2.45, y, 9.3, 0.78, BLACK, 0.36)
    tb(s, f'0{i+1}', 2.65, y+0.18, 0.45, 0.22, 11, GOLD, True)
    tb(s, cap, 3.25, y+0.18, 3.6, 0.25, 14, WHITE, True)
    tb(s, '适合抖音 / 视频号 / 小红书', 7.55, y+0.2, 3.6, 0.22, 8, MUTED, False, PP_ALIGN.RIGHT)
footer(s, 19); animate(s, [s.shapes[1], s.shapes[2]] + list(list(s.shapes)[3:]))

# 20 End
s = prs.slides.add_slide(prs.slide_layouts[6]); bg(s)
cover_pic(s, 'chinese_village2.jpg'); rect(s, 0, 0, W, H, BLACK, 0.58)
tb(s, '邢堡村值得被看见', 1.0, 2.45, 8.4, 0.8, 44, WHITE, True, font='SimSun')
gold_bar(s, 1.02, 3.45, 1.15)
tb(s, '让黄土塬、古院落、麻花香气和板枣甜意，成为山西乡村文旅的新叙事。', 1.02, 3.85, 8.5, 0.6, 20, CREAM)
tb(s, 'Thank you', 9.25, 6.75, 3.0, 0.35, 14, GOLD, True, PP_ALIGN.RIGHT)
animate(s, list(list(s.shapes))[:14])

prs.save(OUT)
print(f'Saved {OUT}: {len(prs.slides.__iter__.__self__._sldIdLst)} slides')
