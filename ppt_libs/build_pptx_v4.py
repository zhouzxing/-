#!/usr/bin/env python3
"""Build xingbao_village_v4.pptx — Enhanced with QR codes and more content
   Uses all 7 valid images + 5 QR codes.
"""
import os, sys
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

# ─── Design System ───────────────────────────────────────────
C = {
    'dark': '1A1A1A', 'deep': '2C1810', 'warm': '3E2C1A',
    'bronze': '8B5E3C', 'gold': 'C8A96A', 'gold2': '8B6B2D',
    'cream': 'F5E6D3', 'white': 'FFFFFF', 'black': '000000',
}
F = {'serif': 'Source Han Serif SC', 'sans': 'Source Han Sans SC', 'mono': 'JetBrains Mono'}
IMG_DIR = '/home/geeker/.hermes/cache/scratch/xingbao_v3_imgs'

def rgb(h): return RGBColor(int(h[0:2],16), int(h[2:4],16), int(h[4:6],16))

def inch(v):
    if isinstance(v, Emu): return v
    return Inches(v)

def rect(slide, l, t, w, h, color):
    r = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, inch(l), inch(t), inch(w), inch(h))
    r.line.fill.background()
    r.fill.solid()
    r.fill.fore_color.rgb = rgb(color)
    return r

def pic(slide, name, l, t, w, h, src_dir=None):
    d = src_dir or IMG_DIR
    p = os.path.join(d, name)
    if os.path.exists(p):
        return slide.shapes.add_picture(p, inch(l), inch(t), inch(w), inch(h))
    return None

def tb(slide, l, t, w, h, text, fs=14, fn='sans', b=False, c='white', al='l', ls=1.3):
    box = slide.shapes.add_textbox(inch(l), inch(t), inch(w), inch(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    p = tf.paragraphs[0]
    p.text = text
    p.alignment = {'l': PP_ALIGN.LEFT, 'c': PP_ALIGN.CENTER, 'r': PP_ALIGN.RIGHT}[al]
    p.line_spacing = ls
    p.space_after = Pt(0)
    run = p.runs[0] if p.runs else p.add_run()
    run.font.size = Pt(fs)
    run.font.bold = b
    run.font.color.rgb = rgb(c)
    run.font.name = F.get(fn, fn)
    return box

def pg(slide, n, label=''):
    tb(slide, 12.0, 7.05, 1.0, 0.3, f'{n:02d} / 22', 11, 'sans', False, C['gold2'], 'r')
    if label:
        tb(slide, 0.5, 7.05, 4.0, 0.3, label, 10, 'sans', False, C['gold2'], 'l')

def gline(slide, l, t, w=1.0):
    return rect(slide, l, t, w, 0.03, C['gold'])

def glint(slide, l, t, w=1.0):
    return rect(slide, l, t, w, 0.025, C['gold'])


# ─── Slide Builders ──────────────────────────────────────────

def cover(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    pic(s, 'loess_plateau.jpg', 0, 0, 13.333, 7.5, '/home/geeker/.hermes/cache/scratch/xingbao_imgs')
    rect(s, 0, 0, 13.333, 7.5, '0A0A0A')
    tb(s, 0.8, 0.5, 3, 0.35, '文旅·稷山', 14, 'sans', True, C['gold'], 'l')
    gline(s, 0.8, 1.0)
    tb(s, 0.8, 3.8, 8, 1.2, '邢堡村', 80, 'serif', True, C['white'], 'l')
    tb(s, 0.8, 5.1, 8, 0.5, '黄土塬上的古村落', 28, 'sans', False, C['cream'], 'l')
    tb(s, 0.8, 5.8, 6, 0.35, '山西 · 运城 · 稷山县 · 化峪镇', 16, 'sans', False, C['gold'], 'l')
    tb(s, 11.0, 7.0, 1.5, 0.3, '2026', 14, 'mono', False, C['gold2'], 'r')

def toc(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg = s.background.fill; bg.solid(); bg.fore_color.rgb = rgb(C['dark'])
    tb(s, 0.8, 0.6, 4, 0.5, 'CONTENTS', 14, 'sans', True, C['gold'], 'l')
    tb(s, 0.8, 1.1, 4, 0.7, '目 录', 44, 'serif', True, C['white'], 'l')
    items = [
        ('01', '地理概貌', '塬·沟·坝——三种地貌'),
        ('02', '历史溯源', '后稷故里·邢国遗韵'),
        ('03', '古村落建筑', '窑洞·四合院·关帝庙'),
        ('04', '农耕文明', '塬地·梯田·沟坝地'),
        ('05', '稷山四宝', '麻花·饼子·鸡蛋·板枣'),
        ('06', '名人古迹', '稷王庙·大佛寺·李家大院'),
        ('07', '网红打卡', '抖音·视频号·小红书'),
    ]
    for i, (n, t, d) in enumerate(items):
        y = 2.2 + i * 0.68
        tb(s, 1.0, y, 0.6, 0.4, n, 24, 'sans', True, C['gold'], 'l')
        tb(s, 1.8, y, 3, 0.35, t, 20, 'sans', True, C['white'], 'l')
        tb(s, 1.8, y+0.3, 5, 0.3, d, 13, 'sans', False, C['cream'], 'l')
    pg(s, 2, 'CONTENTS')

def section(prs, num, title, sub, en, img, page, img_dir=None):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    pic(s, img, 0, 0, 13.333, 7.5, img_dir)
    rect(s, 0, 2.8, 13.333, 4.7, '1A1A1A')
    rect(s, 0.8, 3.2, 0.04, 2.5, C['gold'])
    tb(s, 1.2, 3.2, 2, 0.4, f'CHAPTER {num}', 14, 'sans', True, C['gold'], 'l')
    tb(s, 1.2, 3.7, 10, 1.0, title, 52, 'serif', True, C['white'], 'l')
    tb(s, 1.2, 4.8, 8, 0.5, sub, 22, 'sans', False, C['cream'], 'l')
    tb(s, 1.2, 5.4, 8, 0.3, en, 13, 'sans', False, C['gold2'], 'l')
    pg(s, page, f'PART {num}')

# ─── Geography ──────────────────────────────────────────────

def geo1(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg = s.background.fill; bg.solid(); bg.fore_color.rgb = rgb(C['dark'])
    tb(s, 0.8, 0.4, 3, 0.3, '01 · 地理概貌', 13, 'sans', True, C['gold'], 'l')
    tb(s, 0.8, 0.8, 6, 0.7, '从运城北行·进入黄土塬', 32, 'serif', True, C['white'], 'l')
    tb(s, 0.8, 1.8, 5.5, 0.4, '位置', 18, 'sans', True, C['gold'], 'l')
    tb(s, 0.8, 2.3, 5.5, 1.0, '化峪镇位于稷山县西北部，\n属吕梁山南麓黄土丘陵过渡带。\n稷山县距运城市区约30公里。', 16, 'sans', False, C['cream'], 'l', 1.5)
    tb(s, 0.8, 3.4, 5.5, 0.4, '地形', 18, 'sans', True, C['gold'], 'l')
    tb(s, 0.8, 3.9, 5.5, 1.0, '南北沟壑·中间塬面\n塬面平坦·沟底有季节性河流\n北沟河穿村而过', 16, 'sans', False, C['cream'], 'l', 1.5)
    tb(s, 0.8, 5.0, 5.5, 0.4, '海拔', 18, 'sans', True, C['gold'], 'l')
    tb(s, 0.8, 5.5, 5.5, 0.4, '塬面约800-1000米  沟底约500米', 16, 'sans', False, C['cream'], 'l')
    pic(s, 'china_village2.jpg', 6.5, 1.5, 6.0, 4.5, '/home/geeker/.hermes/cache/scratch/xingbao_imgs')
    rect(s, 6.5, 1.5, 6.0, 0.55, '1A1A1A')
    tb(s, 6.5, 1.55, 6.0, 0.4, '化峪镇·黄土塬上', 14, 'sans', True, C['white'], 'c')
    pg(s, 4, 'GEOGRAPHY')

def geo2(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg = s.background.fill; bg.solid(); bg.fore_color.rgb = rgb(C['dark'])
    tb(s, 0.8, 0.4, 3, 0.3, '01 · 地理概貌', 13, 'sans', True, C['gold'], 'l')
    tb(s, 0.8, 0.8, 8, 0.7, '塬·沟·坝——三种地貌', 32, 'serif', True, C['white'], 'l')
    cols = [
        ('01', '塬面', '海拔800-1000米\n地势平坦开阔\n适合种植小麦、玉米\n机械化耕作条件较好'),
        ('02', '深沟', '深达数十米的沟壑纵横\n沟底有季节性河流\n水土流失严重\n沟坝淤积出肥沃土地'),
        ('03', '梯田', '层层叠叠从塬顶到沟底\n开垦出梯田与沟坝地\n因地制宜·保住水土\n先民们的手笔智慧'),
    ]
    for i, (n, t, d) in enumerate(cols):
        x = 0.8 + i * 4.0
        rect(s, x, 2.0, 3.5, 4.5, '2C2C2C')
        tb(s, x+0.3, 2.2, 0.6, 0.4, n, 28, 'serif', True, C['gold'], 'l')
        tb(s, x+0.3, 2.8, 2.5, 0.4, t, 22, 'sans', True, C['white'], 'l')
        tb(s, x+0.3, 3.4, 2.8, 2.8, d, 14, 'sans', False, C['cream'], 'l', 1.5)
    pg(s, 5, 'GEOGRAPHY')

# ─── History ────────────────────────────────────────────────

def history(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg = s.background.fill; bg.solid(); bg.fore_color.rgb = rgb(C['dark'])
    pic(s, 'china_temple.jpg', 0, 0, 5.5, 7.5, '/home/geeker/.hermes/cache/scratch/xingbao_imgs')
    rect(s, 0, 0, 5.5, 7.5, '1A1A1A')
    tb(s, 0.3, 0.3, 4.5, 0.3, '古  事', 13, 'sans', True, C['gold'], 'l')
    tb(s, 6.0, 0.4, 3, 0.3, '02 · 历史溯源', 13, 'sans', True, C['gold'], 'l')
    tb(s, 6.0, 0.8, 6, 0.7, '邢国遗韵·六百余年', 32, 'serif', True, C['white'], 'l')
    tb(s, 6.0, 1.9, 6.5, 1.0, '稷山县因后稷教民稼穑得名。\n相传上古时期，周人始祖后稷曾在此地\n播百谷、树艺五谷。', 16, 'sans', False, C['cream'], 'l', 1.5)
    tb(s, 6.0, 3.2, 6.5, 1.0, '化峪镇·古寺商贾\n谷中曾有一座古寺，香火鼎盛时\n僧侣往来、商贾穿梭。', 16, 'sans', False, C['cream'], 'l', 1.5)
    tb(s, 6.0, 4.5, 6.5, 1.0, '邢国遗民·避难定居\n一说春秋时期，邢国被卫所灭后，\n一部分邢人北迁辗转来到这片黄土塬上。', 16, 'sans', False, C['cream'], 'l', 1.5)
    rect(s, 6.0, 5.7, 6.5, 1.0, '2C1810')
    gline(s, 6.0, 5.7, 6.5)
    tb(s, 6.3, 5.85, 6.0, 0.8, '六百年古槐——树龄约600年\n相传为明朝初年洪洞大槐树下移民所栽', 14, 'sans', False, C['gold'], 'l', 1.4)
    pg(s, 7, 'HISTORY')

# ─── Architecture ───────────────────────────────────────────

def arch(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg = s.background.fill; bg.solid(); bg.fore_color.rgb = rgb(C['dark'])
    tb(s, 0.8, 0.4, 3, 0.3, '03 · 古村落建筑', 13, 'sans', True, C['gold'], 'l')
    tb(s, 0.8, 0.8, 8, 0.7, '窑洞·四合院·关帝庙', 36, 'serif', True, C['white'], 'l')
    tb(s, 0.8, 1.5, 8, 0.4, '三种建筑形态，诉说着黄土地上的居住智慧与信仰', 16, 'sans', False, C['cream'], 'l')
    cards = [
        ('窑洞', '靠崖窑·砖窑', '在黄土崖壁上挖出的洞穴式居所\n冬暖夏凉·天然恒温房\n窑房结合的格局\n拱形门窗·红纸窗花'),
        ('四合院', '砖雕·门楼·照壁', '晋南传统格局·规模各异\n富裕人家有高大门楼\n精美砖雕照壁·五脊六兽\n院中种石榴树或枣树'),
        ('关帝庙', '正殿·戏台·壁画', '村中央的关帝庙\n一座正殿和一座戏台\n殿内壁画：过五关斩六将\n线条粗犷·色彩斑驳'),
    ]
    for i, (t, st, d) in enumerate(cards):
        x = 0.8 + i * 4.0
        rect(s, x, 2.2, 3.6, 4.8, '2C2C2C')
        gline(s, x, 2.2, 3.6)
        tb(s, x+0.3, 2.8, 2.5, 0.5, t, 24, 'serif', True, C['white'], 'l')
        tb(s, x+0.3, 3.4, 2.5, 0.3, st, 13, 'sans', False, C['gold'], 'l')
        glint(s, x+0.3, 3.8, 1.0)
        tb(s, x+0.3, 4.0, 2.8, 2.8, d, 13, 'sans', False, C['cream'], 'l', 1.5)
    pg(s, 9, 'ARCHITECTURE')

def arch2(prs):
    """Architecture detail — image + description"""
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg = s.background.fill; bg.solid(); bg.fore_color.rgb = rgb(C['dark'])
    pic(s, 'china_roof.jpg', 0, 0, 13.333, 7.5, '/home/geeker/.hermes/cache/scratch/xingbao_imgs')
    rect(s, 0, 0, 13.333, 7.5, '1A1A1A')
    tb(s, 0.8, 0.4, 3, 0.3, '03 · 古村落建筑', 13, 'sans', True, C['gold'], 'l')
    tb(s, 0.8, 0.8, 10, 0.7, '晋南民居·砖雕之精', 36, 'serif', True, C['white'], 'l')
    tb(s, 0.8, 1.7, 5.5, 0.4, '照壁·门楼', 18, 'sans', True, C['gold'], 'l')
    tb(s, 0.8, 2.1, 5.5, 1.0, '进门即见砖雕照壁\n五脊六兽·飞檐翘角\n富丽堂皇·彰显门第', 16, 'sans', False, C['cream'], 'l', 1.5)
    tb(s, 0.8, 3.2, 5.5, 0.4, '巷道', 18, 'sans', True, C['gold'], 'l')
    tb(s, 0.8, 3.6, 5.5, 1.0, '青砖碎石铺就\n被岁月和脚步磨得光滑圆润\n墙头上仙人掌开着黄花', 16, 'sans', False, C['cream'], 'l', 1.5)
    tb(s, 0.8, 4.7, 5.5, 0.4, '民俗', 18, 'sans', True, C['gold'], 'l')
    tb(s, 0.8, 5.1, 5.5, 1.0, '门楣挂红布条或小镜子避邪\n窗棂贴红纸剪的窗花\n春节换一次·年年五谷丰登', 16, 'sans', False, C['cream'], 'l', 1.5)
    # Right side caption
    rect(s, 7.0, 5.5, 5.5, 1.5, '000000')
    gline(s, 7.0, 5.5, 5.5)
    tb(s, 7.3, 5.7, 5.0, 1.2, '晋南建筑以砖雕著称\n三雕(砖雕·木雕·石雕)之精\n在邢堡村老村随处可见', 14, 'sans', False, C['cream'], 'l', 1.5)
    pg(s, 10, 'ARCHITECTURE')

# ─── Agriculture ────────────────────────────────────────────

def agri(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    pic(s, 'wheat_field.jpg', 0, 0, 13.333, 7.5, '/home/geeker/.hermes/cache/scratch/xingbao_imgs')
    rect(s, 0, 2.5, 13.333, 5.0, '1A1A1A')
    tb(s, 0.8, 0.4, 3, 0.3, '04 · 农耕文明', 13, 'sans', True, C['gold'], 'l')
    tb(s, 0.8, 2.8, 11, 0.8, '塬地·梯田·沟坝地', 44, 'serif', True, C['white'], 'l')
    tb(s, 0.8, 3.8, 11, 0.5, '因地制宜的耕作方式——邢堡村人数百年积累下的生存智慧', 20, 'sans', False, C['cream'], 'l')
    tb(s, 0.8, 4.6, 3.5, 0.35, '塬地 → 小麦·玉米', 16, 'sans', True, C['gold'], 'l')
    tb(s, 0.8, 5.0, 3.5, 0.8, '平坦开阔，适合机械化耕作\n六月麦浪翻滚·金黄一片', 13, 'sans', False, C['cream'], 'l', 1.5)
    tb(s, 5.0, 4.6, 3.5, 0.35, '梯田 → 谷子·豆类·红薯', 16, 'sans', True, C['gold'], 'l')
    tb(s, 5.0, 5.0, 3.5, 0.8, '狭窄的梯田，用石块垒起田埂\n保住珍贵的土壤和水分', 13, 'sans', False, C['cream'], 'l', 1.5)
    tb(s, 9.2, 4.6, 3.5, 0.35, '沟坝 → 淤积沃土', 16, 'sans', True, C['gold'], 'l')
    tb(s, 9.2, 5.0, 3.5, 0.8, '沟底淤积出的肥沃土地\n种什么都长得旺', 13, 'sans', False, C['cream'], 'l', 1.5)
    pg(s, 12, 'AGRICULTURE')

def agri2(prs):
    """Agriculture detail — harvest season"""
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg = s.background.fill; bg.solid(); bg.fore_color.rgb = rgb(C['dark'])
    tb(s, 0.8, 0.4, 3, 0.3, '04 · 农耕文明', 13, 'sans', True, C['gold'], 'l')
    tb(s, 0.8, 0.8, 8, 0.7, '六月麦浪·秋收硕果', 36, 'serif', True, C['white'], 'l')
    # Four seasons grid
    seasons = [
        ('春', '春耕', '清明前后·种瓜点豆\n小麦返青·玉米播种\n桃花满塬·杏花飘香'),
        ('夏', '夏收', '六月麦浪·金黄一片\n镰刀割麦的唰唰声\n打麦场上脱粒机整夜轰鸣'),
        ('秋', '秋收', '玉米掰完·红薯翻出\n苹果红透·核桃饱满\n晒场上铺满丰收的颜色'),
        ('冬', '冬藏', '雪落塬面·天地苍茫\n窑洞里围着火炉\n准备过年·杀年猪'),
    ]
    for i, (ch, t, d) in enumerate(seasons):
        x = 0.8 + i * 3.1
        rect(s, x, 2.0, 2.8, 4.5, '2C2C2C')
        tb(s, x+0.2, 2.2, 0.5, 0.4, ch, 36, 'serif', True, C['gold'], 'c')
        tb(s, x+0.2, 2.8, 2.4, 0.35, t, 18, 'sans', True, C['white'], 'c')
        glint(s, x+1.0, 3.3, 0.8)
        tb(s, x+0.2, 3.5, 2.4, 2.8, d, 13, 'sans', False, C['cream'], 'l', 1.5)
    pg(s, 13, 'AGRICULTURE')

# ─── Four Treasures ─────────────────────────────────────────

def food(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg = s.background.fill; bg.solid(); bg.fore_color.rgb = rgb(C['dark'])
    tb(s, 0.8, 0.4, 3, 0.3, '05 · 稷山四宝', 13, 'sans', True, C['gold'], 'l')
    tb(s, 0.8, 0.8, 8, 0.7, '稷山四宝·舌尖上的非遗', 36, 'serif', True, C['white'], 'l')
    tb(s, 0.8, 1.5, 8, 0.3, '国家级非物质文化遗产 · 中国国家地理标志产品', 15, 'sans', False, C['gold'], 'l')
    foods = [
        ('01', '麻花', 'JI SHAN MA HUA', '18道传统工序·手工拧制\n香酥脆爽·油而不腻\n三股拧制·金黄花纹\n始创于隋朝开皇年间'),
        ('02', '饼子', 'JI SHAN BING', '纯碱和面·老面发酵\n芝麻香浓·入口即化\n农家自烤·质朴美味'),
        ('03', '鸡蛋', 'JI SHAN JI DAN', '黄土塬散养土鸡蛋\n营养丰富·口感醇厚\n原生态养殖'),
        ('04', '板枣', 'JI SHAN BAN ZAO', '千年唐枣树祖·自然风干\n皮薄肉厚核小·含糖74.5%\n中国重要农业文化遗产'),
    ]
    for i, (n, t, en, d) in enumerate(foods):
        row, col = i // 2, i % 2
        x = 0.8 + col * 6.0
        y = 2.2 + row * 2.4
        rect(s, x, y, 5.5, 2.1, '2C2C2C')
        rect(s, x, y, 0.04, 2.1, C['gold'])
        tb(s, x+0.4, y+0.15, 0.5, 0.4, n, 20, 'serif', True, C['gold'], 'l')
        tb(s, x+1.0, y+0.15, 3, 0.4, t, 22, 'serif', True, C['white'], 'l')
        tb(s, x+1.0, y+0.55, 3, 0.25, en, 11, 'sans', False, C['gold2'], 'l')
        tb(s, x+1.0, y+0.9, 4.0, 1.0, d, 13, 'sans', False, C['cream'], 'l', 1.5)
    pg(s, 15, 'FOUR TREASURES')

def food2(prs):
    """Food detail — 稷山麻花工艺"""
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg = s.background.fill; bg.solid(); bg.fore_color.rgb = rgb(C['dark'])
    pic(s, 'chinese_village.jpg', 0, 0, 6.0, 7.5, '/home/geeker/.hermes/cache/scratch/xingbao_imgs')
    rect(s, 0, 0, 6.0, 7.5, '1A1A1A')
    tb(s, 0.3, 0.3, 4.5, 0.3, '稷山麻花', 13, 'sans', True, C['gold'], 'l')
    # QR code
    pic(s, 'qr_mahua.png', 4.5, 6.0, 1.0, 1.0)
    tb(s, 0.3, 6.8, 4.5, 0.2, '扫码了解更多', 10, 'sans', False, C['gold2'], 'l')
    # Right content
    tb(s, 6.5, 0.4, 3, 0.3, '05 · 稷山四宝', 13, 'sans', True, C['gold'], 'l')
    tb(s, 6.5, 0.8, 6, 0.7, '十八道工序·千年传承', 32, 'serif', True, C['white'], 'l')
    tb(s, 6.5, 1.6, 6.0, 0.4, '国家级非物质文化遗产', 14, 'sans', True, C['gold'], 'l')
    tb(s, 6.5, 2.1, 6.0, 0.3, '2011年认定 · 赵氏四味坊', 13, 'sans', False, C['cream'], 'l')
    # 18 steps
    tb(s, 6.5, 2.8, 6.0, 0.4, '传统十八道工序', 18, 'sans', True, C['gold'], 'l')
    steps = [
        '培养酵块', '掺水和面', '按比接面',
        '分割面块', '反复揉面', '匀揪面节',
        '面节擦油', '卧缸存放', '分板打畦',
        '搓条上劲', '扭股成形', '入锅油炸',
        '翻拨整形', '出锅淋油', '化熬糖汁',
        '添加辅料', '二次冷却', '存放',
    ]
    for i, step in enumerate(steps):
        row, col = i // 6, i % 6
        x = 6.5 + col * 1.0
        y = 3.3 + row * 0.35
        rect(s, x, y, 0.08, 0.08, C['gold'])
        tb(s, x+0.15, y-0.03, 0.85, 0.25, step, 9, 'sans', False, C['cream'], 'l')
    # Bottom
    rect(s, 6.5, 5.5, 6.0, 1.2, '2C1810')
    gline(s, 6.5, 5.5, 6.0)
    tb(s, 6.7, 5.65, 5.5, 1.0, '始创于隋朝开皇年间\n相传古人在二月初二\n拉面扭作毒蝎尾状油炸\n以此驱毒纳福', 14, 'sans', False, C['gold'], 'l', 1.4)
    pg(s, 16, 'FOUR TREASURES')

# ─── Heritage ───────────────────────────────────────────────

def heritage(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg = s.background.fill; bg.solid(); bg.fore_color.rgb = rgb(C['dark'])
    pic(s, 'big_buddha.jpg', 0, 0, 5.5, 7.5, '/home/geeker/.hermes/cache/scratch/xingbao_imgs')
    rect(s, 0, 0, 5.5, 7.5, '1A1A1A')
    tb(s, 0.3, 0.3, 4.5, 0.3, '古  迹', 13, 'sans', True, C['gold'], 'l')
    # QR code
    pic(s, 'qr_jiwang.png', 4.5, 6.0, 1.0, 1.0)
    tb(s, 0.3, 6.8, 4.5, 0.2, '扫码了解稷王庙', 10, 'sans', False, C['gold2'], 'l')
    tb(s, 6.0, 0.4, 3, 0.3, '06 · 名人古迹', 13, 'sans', True, C['gold'], 'l')
    tb(s, 6.0, 0.8, 6, 0.7, '稷山·千年文脉', 36, 'serif', True, C['white'], 'l')
    items = [
        ('稷王庙', '全国重点文物保护单位', '纪念后稷，稷山历史地标'),
        ('大佛寺', '千年古刹·唐代石刻', '原名清凉院，始建于唐代'),
        ('丁庄李家大院', '晋南民居典范', '砖雕·木雕·石雕三绝'),
        ('高跷走兽', '国家级非物质文化遗产', '高跷表演·走兽艺术'),
        ('螺钿漆器', '山西省级非物质文化遗产', '古老工艺·精美绝伦'),
    ]
    for i, (t, tag, d) in enumerate(items):
        y = 1.9 + i * 1.0
        rect(s, 6.15, y+0.12, 0.12, 0.12, C['gold'])
        tb(s, 6.5, y, 3, 0.3, t, 18, 'sans', True, C['white'], 'l')
        tb(s, 6.5, y+0.3, 6, 0.25, tag, 11, 'sans', False, C['gold'], 'l')
        tb(s, 6.5, y+0.55, 5, 0.25, d, 13, 'sans', False, C['cream'], 'l')
    pg(s, 17, 'HERITAGE')

# ─── Media / Social ─────────────────────────────────────────

def media1(prs):
    """Media overview — platforms with QR codes"""
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg = s.background.fill; bg.solid(); bg.fore_color.rgb = rgb(C['dark'])
    tb(s, 0.8, 0.4, 3, 0.3, '07 · 网红打卡', 13, 'sans', True, C['gold'], 'l')
    tb(s, 0.8, 0.8, 11, 0.7, '让邢堡村被看见', 44, 'serif', True, C['white'], 'l')
    tb(s, 0.8, 1.6, 11, 0.4, '扫码直达各平台，发现更多邢堡村内容', 18, 'sans', False, C['cream'], 'l')
    platforms = [
        ('抖音', 'DOUYIN', '#山西DOU是好风光 播放10亿+\n#稷山麻花 #稷山板枣\n95后博主贾欣锞直播间', 'qr_douyin.png'),
        ('视频号', 'WECHAT VIDEO', '@山西文旅 官方账号\n#化峪镇 #稷山县\n古村落文旅推荐', 'qr_wechat.png'),
        ('小红书', 'XIAOHONGSHU', '#晋南古村落 #黄河风情线\n#稷山旅游\n打卡邢堡村', 'qr_xiaohongshu.png'),
    ]
    for i, (t, en, d, qr) in enumerate(platforms):
        x = 0.8 + i * 4.0
        rect(s, x, 2.5, 3.6, 4.5, '2C2C2C')
        gline(s, x, 2.5, 3.6)
        tb(s, x+0.3, 2.8, 3, 0.35, t, 24, 'serif', True, C['gold'], 'l')
        tb(s, x+0.3, 3.2, 3, 0.25, en, 11, 'sans', False, C['gold2'], 'l')
        glint(s, x+0.3, 3.6, 1.0)
        tb(s, x+0.3, 3.7, 3.0, 1.5, d, 12, 'sans', False, C['cream'], 'l', 1.5)
        pic(s, qr, x+0.3, 5.5, 1.2, 1.2)
        tb(s, x+1.7, 5.7, 1.5, 0.3, '扫码进入', 11, 'sans', False, C['gold2'], 'l')
        tb(s, x+1.7, 6.0, 1.5, 0.2, f'scan to open', 9, 'sans', False, C['gold2'], 'l')
    pg(s, 19, 'MEDIA')

def media2(prs):
    """Media detail — video suggestions"""
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg = s.background.fill; bg.solid(); bg.fore_color.rgb = rgb(C['dark'])
    tb(s, 0.8, 0.4, 3, 0.3, '07 · 网红打卡', 13, 'sans', True, C['gold'], 'l')
    tb(s, 0.8, 0.8, 10, 0.7, '短视频选题建议', 36, 'serif', True, C['white'], 'l')
    tb(s, 0.8, 1.5, 10, 0.4, '参考以下选题方向，创作邢堡村专属内容', 16, 'sans', False, C['cream'], 'l')
    suggestions = [
        ('01', '古槐树的故事', '六百年古槐·洪洞大槐树移民后裔\n拍摄建议：环绕拍摄+特写树皮纹路'),
        ('02', '窑洞里的四季', '冬暖夏凉的天然恒温房\n拍摄建议：四季对比·温度计演示'),
        ('03', '稷山麻花制作', '18道工序手工拧制\n拍摄建议：慢动作特写·油锅翻滚'),
        ('04', '六月麦浪', '金黄麦浪翻滚·打麦场丰收\n拍摄建议：无人机航拍·镰刀收割'),
        ('05', '高跷走兽', '国家级非遗·震撼表演\n拍摄建议：低角度仰拍·全景巡游'),
    ]
    for i, (n, t, d) in enumerate(suggestions):
        y = 2.2 + i * 1.0
        rect(s, 0.8, y, 12.0, 0.8, '2C2C2C')
        tb(s, 1.0, y+0.05, 0.5, 0.3, n, 18, 'serif', True, C['gold'], 'l')
        tb(s, 1.5, y+0.05, 4, 0.3, t, 16, 'sans', True, C['white'], 'l')
        tb(s, 1.5, y+0.4, 10, 0.3, d, 12, 'sans', False, C['cream'], 'l')
    pg(s, 20, 'MEDIA')

def closing(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg = s.background.fill; bg.solid(); bg.fore_color.rgb = rgb(C['dark'])
    gline(s, 6.0, 2.3, 1.333)
    tb(s, 1.5, 2.8, 10.333, 1.2, '期待与您同行', 56, 'serif', True, C['white'], 'c')
    tb(s, 1.5, 4.1, 10.333, 0.5, 'XINGBAO VILLAGE · 邢堡村', 18, 'sans', False, C['gold'], 'c')
    tb(s, 1.5, 4.8, 10.333, 0.4, '山西 · 运城 · 稷山县 · 化峪镇', 14, 'sans', False, C['cream'], 'c')
    tb(s, 1.5, 6.5, 10.333, 0.3, '2026', 12, 'mono', False, C['gold2'], 'c')


def main():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    print("Building v4...")
    cover(prs)
    toc(prs)
    section(prs, '01', '地理概貌', '塬·沟·坝——三种地貌', 'GEOGRAPHY · OVERVIEW', 'loess_plateau.jpg', 3, '/home/geeker/.hermes/cache/scratch/xingbao_imgs')
    geo1(prs)
    geo2(prs)
    section(prs, '02', '历史溯源', '后稷故里·邢国遗韵', 'HISTORY · ORIGIN', 'china_temple.jpg', 6, '/home/geeker/.hermes/cache/scratch/xingbao_imgs')
    history(prs)
    section(prs, '03', '古村落建筑', '窑洞·四合院·关帝庙', 'ARCHITECTURE · HERITAGE', 'china_roof.jpg', 8, '/home/geeker/.hermes/cache/scratch/xingbao_imgs')
    arch(prs)
    arch2(prs)
    section(prs, '04', '农耕文明', '塬地·梯田·沟坝地', 'AGRICULTURE · TRADITION', 'wheat_field.jpg', 11, '/home/geeker/.hermes/cache/scratch/xingbao_imgs')
    agri(prs)
    agri2(prs)
    section(prs, '05', '稷山四宝', '麻花·饼子·鸡蛋·板枣', 'FOUR TREASURES · CULINARY', 'chinese_village.jpg', 14, '/home/geeker/.hermes/cache/scratch/xingbao_imgs')
    food(prs)
    food2(prs)
    section(prs, '06', '名人古迹', '稷王庙·大佛寺·李家大院', 'HERITAGE · LANDMARKS', 'big_buddha.jpg', 16, '/home/geeker/.hermes/cache/scratch/xingbao_imgs')
    heritage(prs)
    section(prs, '07', '网红打卡', '抖音·视频号·小红书', 'MEDIA · SOCIAL', 'loess_plateau.jpg', 18, '/home/geeker/.hermes/cache/scratch/xingbao_imgs')
    media1(prs)
    media2(prs)
    closing(prs)
    out = sys.argv[1] if len(sys.argv) > 1 else 'xingbao_village_v4.pptx'
    prs.save(out)
    print(f"Saved: {out}")
    print(f"Slides: {len(prs.slides)}")

if __name__ == '__main__':
    main()
