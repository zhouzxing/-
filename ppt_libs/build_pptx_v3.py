#!/usr/bin/env python3
"""Build xingbao_village_v3.pptx — Premium Travel Design
   All position/size values in INCHES, converted via Inches().
   Only python-pptx native API. No lxml XML.
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
IMG = '/home/geeker/.hermes/cache/scratch/xingbao_imgs'

def rgb(h): return RGBColor(int(h[0:2],16), int(h[2:4],16), int(h[4:6],16))

def inch(v):
    """If already Emu/Inches, return as-is; else wrap."""
    if isinstance(v, Emu): return v
    return Inches(v)

def rect(slide, l, t, w, h, color, thick=None):
    """Rectangle. All args in inches. thick: for gold lines, in points."""
    ht = Pt(thick) if isinstance(thick, float) and thick < 10 else inch(h)
    r = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, inch(l), inch(t), inch(w), ht)
    r.line.fill.background()
    r.fill.solid()
    r.fill.fore_color.rgb = rgb(color)
    return r

def pic(slide, name, l, t, w, h):
    p = os.path.join(IMG, name)
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
    tb(slide, 12.0, 7.05, 1.0, 0.3, f'{n:02d} / 18', 11, 'sans', False, C['gold2'], 'r')
    if label:
        tb(slide, 0.5, 7.05, 4.0, 0.3, label, 10, 'sans', False, C['gold2'], 'l')

def gline(slide, l, t, w=1.0):
    """Gold accent line."""
    return rect(slide, l, t, w, 0.03, C['gold'])

def glint(slide, l, t, w=1.0):
    """Gold accent line — thin."""
    return rect(slide, l, t, w, 0.025, C['gold'])


# ─── Slide Builders ──────────────────────────────────────────

def cover(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    pic(s, 'loess_plateau.jpg', 0, 0, 13.333, 7.5)
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
        ('01', '地理概貌', '黄土高原上的古村落'),
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

def section(prs, num, title, sub, en, img, page):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    pic(s, img, 0, 0, 13.333, 7.5)
    rect(s, 0, 2.8, 13.333, 4.7, '1A1A1A')
    rect(s, 0.8, 3.2, 0.04, 2.5, C['gold'])
    tb(s, 1.2, 3.2, 2, 0.4, f'CHAPTER {num}', 14, 'sans', True, C['gold'], 'l')
    tb(s, 1.2, 3.7, 10, 1.0, title, 52, 'serif', True, C['white'], 'l')
    tb(s, 1.2, 4.8, 8, 0.5, sub, 22, 'sans', False, C['cream'], 'l')
    tb(s, 1.2, 5.4, 8, 0.3, en, 13, 'sans', False, C['gold2'], 'l')
    pg(s, page, f'PART {num}')

def geo1(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg = s.background.fill; bg.solid(); bg.fore_color.rgb = rgb(C['dark'])
    tb(s, 0.8, 0.4, 3, 0.3, '01 · 地理概貌', 13, 'sans', True, C['gold'], 'l')
    tb(s, 0.8, 0.8, 6, 0.7, '从运城北行·进入黄土塬', 32, 'serif', True, C['white'], 'l')
    tb(s, 0.8, 1.8, 5.5, 0.4, '位置', 18, 'sans', True, C['gold'], 'l')
    tb(s, 0.8, 2.3, 5.5, 1.0, '化峪镇位于稷山县西北部，\n属吕梁山南麓黄土丘陵过渡带。', 16, 'sans', False, C['cream'], 'l', 1.5)
    tb(s, 0.8, 3.3, 5.5, 0.4, '地形', 18, 'sans', True, C['gold'], 'l')
    tb(s, 0.8, 3.8, 5.5, 1.0, '南北沟壑·中间塬面\n塬面平坦·沟底有季节性河流', 16, 'sans', False, C['cream'], 'l', 1.5)
    tb(s, 0.8, 4.8, 5.5, 0.4, '海拔', 18, 'sans', True, C['gold'], 'l')
    tb(s, 0.8, 5.3, 5.5, 0.4, '塬面约800-1000米  沟底约500米', 16, 'sans', False, C['cream'], 'l')
    pic(s, 'china_village2.jpg', 6.5, 1.5, 6.0, 4.5)
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

def history(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg = s.background.fill; bg.solid(); bg.fore_color.rgb = rgb(C['dark'])
    pic(s, 'china_temple.jpg', 0, 0, 5.5, 7.5)
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

def agri(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    pic(s, 'wheat_field.jpg', 0, 0, 13.333, 7.5)
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
    pg(s, 11, 'AGRICULTURE')

def food(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg = s.background.fill; bg.solid(); bg.fore_color.rgb = rgb(C['dark'])
    tb(s, 0.8, 0.4, 3, 0.3, '05 · 稷山四宝', 13, 'sans', True, C['gold'], 'l')
    tb(s, 0.8, 0.8, 8, 0.7, '稷山四宝·舌尖上的非遗', 36, 'serif', True, C['white'], 'l')
    tb(s, 0.8, 1.5, 8, 0.3, '国家级非物质文化遗产 · 中国国家地理标志产品', 15, 'sans', False, C['gold'], 'l')
    foods = [
        ('01', '麻花', 'JI SHAN MA HUA', '18道传统工序·手工拧制\n香酥脆爽·回味无穷'),
        ('02', '饼子', 'JI SHAN BING', '纯碱和面·老面发酵\n芝麻香浓·入口即化'),
        ('03', '鸡蛋', 'JI SHAN JI DAN', '黄土塬散养土鸡蛋\n营养丰富·口感醇厚'),
        ('04', '板枣', 'JI SHAN BAN ZAO', '百年古枣园·自然风干\n果肉厚实·甘甜可口'),
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
    pg(s, 13, 'FOUR TREASURES')

def heritage(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg = s.background.fill; bg.solid(); bg.fore_color.rgb = rgb(C['dark'])
    pic(s, 'big_buddha.jpg', 0, 0, 5.5, 7.5)
    rect(s, 0, 0, 5.5, 7.5, '1A1A1A')
    tb(s, 0.3, 0.3, 4.5, 0.3, '古  迹', 13, 'sans', True, C['gold'], 'l')
    tb(s, 6.0, 0.4, 3, 0.3, '06 · 名人古迹', 13, 'sans', True, C['gold'], 'l')
    tb(s, 6.0, 0.8, 6, 0.7, '稷山·千年文脉', 36, 'serif', True, C['white'], 'l')
    items = [
        ('稷王庙', '全国重点文物保护单位', '纪念后稷，稷山历史地标'),
        ('大佛寺', '千年古刹·唐代石刻', '山中的古寺，历史沧桑'),
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
    pg(s, 15, 'HERITAGE')

def media(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    pic(s, 'loess_plateau.jpg', 0, 0, 13.333, 7.5)
    rect(s, 0, 2.5, 13.333, 5.0, '1A1A1A')
    tb(s, 0.8, 0.4, 3, 0.3, '07 · 网红打卡', 13, 'sans', True, C['gold'], 'l')
    tb(s, 0.8, 2.8, 11, 0.8, '让邢堡村被看见', 44, 'serif', True, C['white'], 'l')
    platforms = [
        ('抖音', 'DOUYIN', '#山西DOU是好风光 10亿+播放\n#稷山麻花 #稷山板枣\n95后博主贾欣锞直播间'),
        ('视频号', 'WECHAT VIDEO', '@山西文旅 官方账号\n#化峪镇 #稷山县\n古村落文旅推荐'),
        ('小红书', 'XIAOHONGSHU', '#晋南古村落 #黄河风情线\n#稷山旅游\n打卡邢堡村'),
    ]
    for i, (t, en, d) in enumerate(platforms):
        x = 0.8 + i * 4.0
        rect(s, x, 4.3, 3.6, 2.5, '000000')
        tb(s, x+0.3, 4.5, 3, 0.3, t, 20, 'sans', True, C['gold'], 'l')
        tb(s, x+0.3, 4.85, 3, 0.25, en, 10, 'sans', False, C['gold2'], 'l')
        glint(s, x+0.3, 5.2, 1.0)
        tb(s, x+0.3, 5.35, 3.0, 1.3, d, 12, 'sans', False, C['cream'], 'l', 1.5)
    pg(s, 17, 'MEDIA')

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
    print("Building...")
    cover(prs)
    toc(prs)
    section(prs, '01', '地理概貌', '塬·沟·坝——三种地貌', 'GEOGRAPHY · OVERVIEW', 'loess_plateau.jpg', 3)
    geo1(prs)
    geo2(prs)
    section(prs, '02', '历史溯源', '后稷故里·邢国遗韵', 'HISTORY · ORIGIN', 'china_temple.jpg', 6)
    history(prs)
    section(prs, '03', '古村落建筑', '窑洞·四合院·关帝庙', 'ARCHITECTURE · HERITAGE', 'china_roof.jpg', 8)
    arch(prs)
    section(prs, '04', '农耕文明', '塬地·梯田·沟坝地', 'AGRICULTURE · TRADITION', 'wheat_field.jpg', 10)
    agri(prs)
    section(prs, '05', '稷山四宝', '麻花·饼子·鸡蛋·板枣', 'FOUR TREASURES · CULINARY', 'chinese_village.jpg', 12)
    food(prs)
    section(prs, '06', '名人古迹', '稷王庙·大佛寺·李家大院', 'HERITAGE · LANDMARKS', 'big_buddha.jpg', 14)
    heritage(prs)
    section(prs, '07', '网红打卡', '抖音·视频号·小红书', 'MEDIA · SOCIAL', 'loess_plateau.jpg', 16)
    media(prs)
    closing(prs)
    out = sys.argv[1] if len(sys.argv) > 1 else 'xingbao_village_v3.pptx'
    prs.save(out)
    print(f"Saved: {out}")
    print(f"Slides: {len(prs.slides)}")

if __name__ == '__main__':
    main()
