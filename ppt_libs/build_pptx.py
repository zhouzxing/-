#!/usr/bin/env python3
"""Build xingbao_village.pptx with embedded images"""
import json, sys
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

IMG_DIR = "/home/geeker/.hermes/cache/scratch/xingbao_imgs"
OUT = "/home/geeker/architecture/home/geeker/notes/xingbao_village_v2.pptx"

# Colors
PRIMARY = RGBColor(0x8B, 0x5E, 0x3C)
SECONDARY = RGBColor(0xC4, 0x95, 0x6C)
ACCENT = RGBColor(0xD4, 0xA8, 0x53)
LIGHT = RGBColor(0xF5, 0xE6, 0xD3)
DARK = RGBColor(0x2C, 0x18, 0x10)
BG_DARK = RGBColor(0x3E, 0x2C, 0x1A)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
MUTED = RGBColor(0x8B, 0x73, 0x55)

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

def set_bg(slide, color):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color

def add_textbox(slide, left, top, width, height, text, font_size=Pt(14),
                bold=False, color=WHITE, font_name="Arial", align="left"):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = font_size
    p.font.bold = bold
    p.font.color.rgb = color
    p.font.name = font_name
    if align == "center":
        from pptx.enum.text import PP_ALIGN
        p.alignment = PP_ALIGN.CENTER
    return txBox

def add_bullet_list(slide, left, top, width, height, items, font_size=Pt(13)):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.space_after = Pt(3)
        run = p.add_run()
        run.text = item
        run.font.size = font_size  # already Pt() wrapped
        run.font.color.rgb = WHITE
        run.font.name = "Arial"
        if item.startswith("●") or item.startswith("◉") or item.startswith("•"):
            run.font.size = Pt(14)
            run.font.color.rgb = ACCENT
    return txBox

# ============================================================
# SLIDE DEFINITIONS
# ============================================================
prs = Presentation()
prs.slide_width = SLIDE_W
prs.slide_height = SLIDE_H

# --- SLIDE 1: COVER ---
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, BG_DARK)
# Background image (loess)
slide.shapes.add_picture(
    f"{IMG_DIR}/loess_plateau.jpg",
    left=Inches(0), top=Inches(0), width=SLIDE_W, height=SLIDE_H
)
# Overlay text box
add_textbox(slide, Inches(1.5), Inches(2.0), Inches(10), Inches(1.5),
            "黄 土 塬 上 的 古 村 落", Pt(44), True, WHITE, align="center")
add_textbox(slide, Inches(1.5), Inches(3.8), Inches(10), Inches(0.8),
            "山西运城稷山县化峪镇邢堡村纪行", Pt(20), False, ACCENT, align="center")
add_textbox(slide, Inches(1.5), Inches(5.0), Inches(10), Inches(0.8),
            "黄土高原 · 后稷故里 · 华夏农耕文明发祥地", Pt(16), False, SECONDARY, align="center")

# --- SLIDE 2: TOC ---
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, DARK)
add_textbox(slide, Inches(1.0), Inches(0.5), Inches(11), Inches(0.8),
            "目  录", Pt(36), True, ACCENT)
toc_items = [
    "壹 · 地理概貌 —— 吕梁山南麓，黄土塬上",
    "贰 · 历史溯源 —— 后稷故里，华夏农耕文明发祥地",
    "参 · 古村落建筑 —— 窑洞 · 四合院 · 关帝庙 · 古井",
    "肆 · 农耕文明 —— 塬地 · 梯田 · 沟坝地 · 麦浪金黄",
    "伍 · 稷山四宝 —— 麻花 · 饼子 · 鸡蛋 · 板枣",
    "陆 · 名人古迹 —— 大佛寺 · 稷王庙 · 丁庄李家大院",
    "柒 · 网红打卡 —— 抖音 · 视频号 · 农文旅融合",
]
add_bullet_list(slide, Inches(1.5), Inches(1.8), Inches(10.5), Inches(5.0),
                toc_items, Pt(16))

# --- SLIDE 3: GEOGRAPHY HEADER ---
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, LIGHT)
add_textbox(slide, Inches(0.5), Inches(0.5), Inches(12), Inches(1.0),
            "壹 · 地理概貌", Pt(36), True, PRIMARY)
add_textbox(slide, Inches(0.5), Inches(1.6), Inches(12), Inches(0.8),
            "吕梁山南麓 · 黄土塬上 · 化峪河谷", Pt(22), False, SECONDARY)
# Image
slide.shapes.add_picture(
    f"{IMG_DIR}/loess_plateau.jpg",
    left=Inches(0.5), top=Inches(2.5), width=Inches(12.333), height=Inches(4.8)
)

# --- SLIDE 4: GEOGRAPHY DETAIL ---
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, LIGHT)
add_textbox(slide, Inches(0.5), Inches(0.3), Inches(12), Inches(0.7),
            "地理区位与地形地貌", Pt(28), True, PRIMARY)
# Left column
left_txt = (
    "位置：吕梁山前沿脚下，黄华峪口正南五华里\n"
    "距化峪镇2.5公里，距稷山县城约20公里\n"
    "村域面积约4550亩，耕地4024亩\n"
    "北临开东村，南接程杜村，交通四通八达\n"
    "村北侯西高速直达西安，村南管化路穿村而过"
)
add_textbox(slide, Inches(0.5), Inches(1.2), Inches(5.8), Inches(2.5),
            left_txt, Pt(14), False, DARK)
# Right column
right_txt = (
    "地形：典型黄土高原丘陵地貌\n"
    "整个村庄坐落在东西走向的黄土塬上\n"
    "北面深沟数米，沟底季节性小河\n"
    "南面层层叠叠的梯田，如同天阶\n"
    "塬地平坦、梯田狭窄、沟坝地肥沃"
)
add_textbox(slide, Inches(6.8), Inches(1.2), Inches(5.8), Inches(2.5),
            right_txt, Pt(14), False, DARK)
# Image bottom
slide.shapes.add_picture(
    f"{IMG_DIR}/china_village2.jpg",
    left=Inches(0.5), top=Inches(3.9), width=Inches(12.333), height=Inches(3.4)
)

# --- SLIDE 5: HISTORY HEADER ---
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, LIGHT)
add_textbox(slide, Inches(0.5), Inches(0.5), Inches(12), Inches(1.0),
            "贰 · 历史溯源", Pt(36), True, PRIMARY)
add_textbox(slide, Inches(0.5), Inches(1.6), Inches(12), Inches(0.8),
            "后稷故里 · 华夏农耕文明发祥地 · 千年古县", Pt(22), False, SECONDARY)
slide.shapes.add_picture(
    f"{IMG_DIR}/china_roof.jpg",
    left=Inches(0.5), top=Inches(2.6), width=Inches(12.333), height=Inches(4.7)
)

# --- SLIDE 6: HOUJI HISTORY ---
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, LIGHT)
add_textbox(slide, Inches(0.5), Inches(0.3), Inches(12), Inches(0.7),
            "后稷教民稼穑——华夏农耕之源", Pt(26), True, PRIMARY)
add_textbox(slide, Inches(0.5), Inches(1.0), Inches(12), Inches(0.6),
            "相传上古时期，周人始祖后稷在此'播百谷，树艺五谷'", Pt(16), False, SECONDARY)
bullets = [
    "● 稷山，中国首批'千年古县'之一",
    "  后稷（弃），周人始祖，被尊为'农神'",
    "  '播百谷，树艺五谷'——中华农耕文明从这里开始",
    "  国家板枣公园·守望千年·只为枣你",
    "  稷山四宝：麻花、饼子、鸡蛋、板枣",
    "  7处国家级文物保护单位：大佛寺、稷王庙、青龙寺、宋金墓等",
]
add_bullet_list(slide, Inches(0.5), Inches(1.8), Inches(12.5), Inches(3.0),
                bullets, Pt(14))
# Image
slide.shapes.add_picture(
    f"{IMG_DIR}/china_temple.jpg",
    left=Inches(0.5), top=Inches(4.5), width=Inches(12.333), height=Inches(2.8)
)

# --- SLIDE 7: VILLAGE NAME ---
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, LIGHT)
add_textbox(slide, Inches(0.5), Inches(0.3), Inches(12), Inches(0.7),
            "邢堡村名溯源与古槐记忆", Pt(26), True, PRIMARY)
add_textbox(slide, Inches(0.5), Inches(1.0), Inches(12), Inches(0.6),
            "六百余年古槐树，洪洞大移民的见证", Pt(16), False, SECONDARY)
bullets = [
    "● '邢堡'村名传说",
    "  一说：古时邢姓人家在此筑堡而居",
    "  另一说：春秋时期邢国遗民避难所，邢人北迁定居",
    "● 古槐树——村口的六百余年老树",
    "  明朝初年洪洞大移民时栽下",
    "  树干需三个成年人合抱，树皮皲裂如鳞，枝叶繁茂",
    "  '问我祖先在何处，山西洪洞大槐树'——移民记忆的活化石",
]
add_bullet_list(slide, Inches(0.5), Inches(1.8), Inches(12.5), Inches(5.0),
                bullets, Pt(14))

# --- SLIDE 8: ARCHITECTURE HEADER ---
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, LIGHT)
add_textbox(slide, Inches(0.5), Inches(0.5), Inches(12), Inches(1.0),
            "参 · 古村落建筑", Pt(36), True, PRIMARY)
add_textbox(slide, Inches(0.5), Inches(1.6), Inches(12), Inches(0.8),
            "窑洞 · 四合院 · 关帝庙 · 古井", Pt(22), False, SECONDARY)
slide.shapes.add_picture(
    f"{IMG_DIR}/china_roof.jpg",
    left=Inches(0.5), top=Inches(2.6), width=Inches(12.333), height=Inches(4.7)
)

# --- SLIDE 9: CAVE + COURTYARD ---
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, LIGHT)
add_textbox(slide, Inches(0.5), Inches(0.3), Inches(12), Inches(0.7),
            "晋南传统民居：窑洞与四合院", Pt(26), True, PRIMARY)
left_txt = (
    "靠崖窑洞\n"
    "  在黄土崖壁上挖出的洞穴式居所\n"
    "  冬暖夏凉——天然'恒温房'\n"
    "  门窗拱形，窗棂贴红纸窗花\n"
    "  '五谷丰登''六畜兴旺'图案年年换新\n"
    "  窑房结合：窑洞前接出一间砖房"
)
add_textbox(slide, Inches(0.5), Inches(1.2), Inches(5.8), Inches(4.0),
            left_txt, Pt(14), False, DARK)
right_txt = (
    "晋南四合院\n"
    "  富裕人家：高大门楼、砖雕照壁、五脊六兽\n"
    "  普通人家：三面土窑或砖窑围出方院\n"
    "  院中种石榴树或枣树——吉祥寓意\n"
    "  院门贴对联、门楣挂红布条或小镜子\n"
    "  巷道：青砖碎石路面，被岁月磨得光滑圆润"
)
add_textbox(slide, Inches(6.8), Inches(1.2), Inches(5.8), Inches(4.0),
            right_txt, Pt(14), False, DARK)
slide.shapes.add_picture(
    f"{IMG_DIR}/china_temple.jpg",
    left=Inches(0.5), top=Inches(4.9), width=Inches(12.333), height=Inches(2.4)
)

# --- SLIDE 10: CULTURE LANDMARKS ---
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, LIGHT)
add_textbox(slide, Inches(0.5), Inches(0.3), Inches(12), Inches(0.7),
            "古村落文化地标", Pt(26), True, PRIMARY)
add_textbox(slide, Inches(0.5), Inches(1.0), Inches(12), Inches(0.6),
            "关帝庙 · 古井 · 碑刻——村庄的记忆坐标", Pt(16), False, SECONDARY)
bullets = [
    "● 关帝庙",
    "  村中央，一间正殿 + 一座戏台",
    "  殿内壁画：关羽过五关斩六将，线条粗犷，民间画匠手笔",
    "  戏台对联：'借虚事指点实事，托古人提醒今人'",
    "● 古井",
    "  村东头，井口石沿被井绳磨出深深凹槽",
    "  '大清乾隆年间重修'石碑——见证井的悠久历史",
]
add_bullet_list(slide, Inches(0.5), Inches(1.8), Inches(12.5), Inches(5.0),
                bullets, Pt(14))

# --- SLIDE 11: AGRICULTURE HEADER ---
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, LIGHT)
add_textbox(slide, Inches(0.5), Inches(0.5), Inches(12), Inches(1.0),
            "肆 · 农耕文明", Pt(36), True, PRIMARY)
add_textbox(slide, Inches(0.5), Inches(1.6), Inches(12), Inches(0.8),
            "黄土塬上的智慧耕作 · 麦浪翻滚的金黄时节", Pt(22), False, SECONDARY)
slide.shapes.add_picture(
    f"{IMG_DIR}/wheat_field.jpg",
    left=Inches(0.5), top=Inches(2.6), width=Inches(12.333), height=Inches(4.7)
)

# --- SLIDE 12: TERRACES ---
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, LIGHT)
add_textbox(slide, Inches(0.5), Inches(0.3), Inches(12), Inches(0.7),
            "梯田·塬地·沟坝地——因地制宜的智慧", Pt(26), True, PRIMARY)
left_txt = (
    "塬地：平坦，适合种小麦和玉米\n"
    "梯田：狭窄，多种谷子、豆类和红薯\n"
    "沟坝地：沟底淤积出的肥沃土地，种什么都长得旺"
)
add_textbox(slide, Inches(0.5), Inches(1.2), Inches(5.8), Inches(2.0),
            left_txt, Pt(14), False, DARK)
right_txt = (
    "麦收时节\n"
    "  六月麦浪翻滚，黄土塬染上金黄色\n"
    "  家家户户男女老少齐上阵\n"
    "  镰刀割麦'唰唰'声此起彼伏\n"
    "  打麦场上麦垛堆得像小山\n"
    "  脱粒机整夜轰鸣，新麦清香弥漫"
)
add_textbox(slide, Inches(6.8), Inches(1.2), Inches(5.8), Inches(3.0),
            right_txt, Pt(14), False, DARK)
slide.shapes.add_picture(
    f"{IMG_DIR}/wheat_field.jpg",
    left=Inches(0.5), top=Inches(4.4), width=Inches(12.333), height=Inches(3.0)
)

# --- SLIDE 13: JISHAN FOUR TREASURES HEADER ---
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, LIGHT)
add_textbox(slide, Inches(0.5), Inches(0.5), Inches(12), Inches(1.0),
            "伍 · 稷山四宝——美食与特产", Pt(32), True, PRIMARY)
add_textbox(slide, Inches(0.5), Inches(1.6), Inches(12), Inches(0.8),
            "麻花 · 饼子 · 鸡蛋 · 板枣", Pt(24), False, SECONDARY)
slide.shapes.add_picture(
    f"{IMG_DIR}/chinese_village.jpg",
    left=Inches(0.5), top=Inches(2.6), width=Inches(12.333), height=Inches(4.7)
)

# --- SLIDE 14: FOUR TREASURES DETAIL ---
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, LIGHT)
add_textbox(slide, Inches(0.5), Inches(0.3), Inches(12), Inches(0.7),
            "稷山四宝", Pt(28), True, PRIMARY)
add_textbox(slide, Inches(0.5), Inches(1.0), Inches(12), Inches(0.6),
            "麻花 | 饼子 | 鸡蛋 | 板枣", Pt(18), False, SECONDARY)
bullets = [
    "🍜 稷山麻花——国家非物质文化遗产",
    "  '赵氏四味坊'历经六代人传承，18道纯手工技艺流程",
    "  油酥麻花：五谷香、爽心甜、到口酥、家常脆四种口味",
    "  年销量超1.5万吨，销售额达1.2亿元",
    "🥮 稷山饼子——打饼子专业户遍布全国16省市",
    "  数万人从事打饼子产业，每年创收超5亿元",
    "🥚 稷山鸡蛋——'中国鸡蛋十大品牌'",
    "  蛋鸡存栏1600万只，日产鸡蛋1200万枚",
    "🌰 稷山板枣——'枣中之王'",
    "  15.3万亩枣树，1.75万株千年以上古树",
    "  2009年获评'中国十大名枣之首'",
]
add_bullet_list(slide, Inches(0.5), Inches(1.6), Inches(12.5), Inches(5.5),
                bullets, Pt(13))

# --- SLIDE 15: MAHUA STORY ---
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, LIGHT)
add_textbox(slide, Inches(0.5), Inches(0.3), Inches(12), Inches(0.7),
            "非遗美食：稷山麻花的制作传奇", Pt(26), True, PRIMARY)
add_textbox(slide, Inches(0.5), Inches(1.0), Inches(12), Inches(0.6),
            "18道工序 · 六代传承 · 1600年历史", Pt(16), False, SECONDARY)
bullets = [
    "起源：南北朝时期，民间'咬蝎尾'习俗——油炸蝎尾状面食",
    "唐代：宰相将家乡麻花介绍给朝中同僚，成为宫廷佳品",
    "清代：'赵氏四味坊'百年老店创立",
    "今天：国家非遗 + 中华老字号，畅销全国31省市",
]
add_bullet_list(slide, Inches(0.5), Inches(1.8), Inches(12.5), Inches(2.0),
                bullets, Pt(14))
steps_txt = (
    "18道工序：\n"
    "培养酵块→掺水接面→分接面→掺水掺油→反复揉面→\n"
    "面节擦油→卧缸存放→搓条上劲→分板打畦→扭股成形→\n"
    "入锅油炸→添加辅料→出锅淋油→定型摆放→化熬糖汁→翻拨整形"
)
add_textbox(slide, Inches(0.5), Inches(3.9), Inches(12.5), Inches(2.0),
            steps_txt, Pt(12), False, MUTED)

# --- SLIDE 16: ANCIENT SITES HEADER ---
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, LIGHT)
add_textbox(slide, Inches(0.5), Inches(0.5), Inches(12), Inches(1.0),
            "陆 · 名人古迹 · 文化名片", Pt(36), True, PRIMARY)
add_textbox(slide, Inches(0.5), Inches(1.6), Inches(12), Inches(0.8),
            "7处国宝 · 3处全国重点文保单位 · 千年板枣古稀树群", Pt(20), False, SECONDARY)
slide.shapes.add_picture(
    f"{IMG_DIR}/big_buddha.jpg",
    left=Inches(0.5), top=Inches(2.6), width=Inches(12.333), height=Inches(4.7)
)

# --- SLIDE 17: ANCIENT SITES DETAIL ---
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, LIGHT)
add_textbox(slide, Inches(0.5), Inches(0.3), Inches(12), Inches(0.7),
            "稷山国宝级古迹", Pt(28), True, PRIMARY)
bullets = [
    "⛩️ 稷山大佛寺（AAA级景区）",
    "  始建于金代皇统二年（1142年），距今880余年",
    "  金代依崖巨佛：高20米，宽6.7米，气魄雄伟",
    "  山西省面积最大的单体仿元建筑",
    "⛩️ 稷王庙（全国重点文保单位）",
    "  始建于元至正五年（1345），占地10080平方米",
    "  集石雕、木刻、琉璃为一体，古建筑'三绝'",
    "⛩️ 丁庄李家大院",
    "  十二座四合院组成，东西长89.4米，南北宽34.5米",
    "⛩️ 其他国宝：青龙寺、宋金墓、马村墓室、法王庙等",
]
add_bullet_list(slide, Inches(0.5), Inches(1.0), Inches(12.5), Inches(5.0),
                bullets, Pt(13))
slide.shapes.add_picture(
    f"{IMG_DIR}/china_roof.jpg",
    left=Inches(0.5), top=Inches(5.3), width=Inches(12.333), height=Inches(2.0)
)

# --- SLIDE 18: INTANGIBLE HERITAGE ---
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, LIGHT)
add_textbox(slide, Inches(0.5), Inches(0.3), Inches(12), Inches(0.7),
            "非遗民俗：高跷走兽 · 高台花鼓 · 螺钿漆器", Pt(26), True, PRIMARY)
bullets = [
    "🎭 高跷走兽——稷山独特民俗",
    "  表演者脚踩高跷，头戴兽首，模仿动物形态游行",
    "🥁 稷山高台花鼓——国家级非物质文化遗产",
    "  鼓乐铿锵，舞步矫健，是晋南最盛大的民间表演",
    "🎨 稷山螺钿漆器髤饰技艺——国家级非物质文化遗产",
    "  以螺壳、贝壳等天然材料镶嵌漆器，图案精美",
    "🪕 金银细工制作技艺",
    "  传统金银加工技艺，是晋南民间工艺的代表",
    "🎯 邢堡村广场舞：《两只蝴蝶》表演队——17人专业队伍",
]
add_bullet_list(slide, Inches(0.5), Inches(1.0), Inches(12.5), Inches(5.0),
                bullets, Pt(13))

# --- SLIDE 19: SOCIAL MEDIA HEADER ---
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, LIGHT)
add_textbox(slide, Inches(0.5), Inches(0.5), Inches(12), Inches(1.0),
            "柒 · 网红打卡与短视频传播", Pt(32), True, PRIMARY)
add_textbox(slide, Inches(0.5), Inches(1.6), Inches(12), Inches(0.8),
            "抖音 · 视频号 · 山西DOU是好风光", Pt(22), False, SECONDARY)
slide.shapes.add_picture(
    f"{IMG_DIR}/loess_plateau.jpg",
    left=Inches(0.5), top=Inches(2.6), width=Inches(12.333), height=Inches(4.7)
)

# --- SLIDE 20: SOCIAL MEDIA DETAIL ---
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, LIGHT)
add_textbox(slide, Inches(0.5), Inches(0.3), Inches(12), Inches(0.7),
            "短视频传播矩阵", Pt(26), True, PRIMARY)
left_txt = (
    "🔥 '山西DOU是好风光'项目\n"
    "  字节跳动公益 + 抖音生活服务联合发起\n"
    "  话题播放量超10亿+，总投稿2万+\n"
    "  千万粉丝创作者@阿波发布游山西内容\n"
    "  带动@厦门芬兰 @福建大乔 @阿远旅行等百万粉创作者投稿\n"
    "📱 博主案例\n"
    "  '发癫吧，后浪！'——抖音博主探访山西古村\n"
    "  百万点赞，带动游客蜂拥而至\n"
    "  95后'贾欣锞'直播间——稷山农耕文明故事\n"
    "  年销售额近400万元，远销缅甸印尼"
)
add_textbox(slide, Inches(0.5), Inches(1.1), Inches(5.8), Inches(5.5),
            left_txt, Pt(12), False, DARK)
right_txt = (
    "🎥 参考视频链接\n"
    "  抖音：#山西DOU是好风光\n"
    "  https://www.douyin.com/search/%E5%B1%B1%E8%A5%BFDOU%E6%98%AF%E5%A5%BD%E9%A3%8E%E5%85%89\n"
    "\n"
    "  抖音：#稷山麻花\n"
    "  https://www.douyin.com/search/%E7%A8%BC%E5%B1%B1%E9%BA%BB%E8%8A%B1\n"
    "\n"
    "  抖音：#稷山板枣\n"
    "  https://www.douyin.com/search/%E7%A8%BC%E5%B1%B1%E6%9D%BF%E6%9E%A3\n"
    "\n"
    "  抖音：#晋宝 山西文旅\n"
    "  https://www.douyin.com/search/%E6%B5%85%E5%AE%9D%E5%B1%B1%E8%A5%BF\n"
    "\n"
    "  视频号：@山西文旅 官方账号\n"
    "\n"
    "📌 拍摄建议\n"
    "  麦收季航拍：黄土塬上的金黄麦浪\n"
    "  古窑洞内景：冬暖夏凉的恒温空间\n"
    "  麻花制作全过程：18道工序慢镜头\n"
    "  板枣采摘：千年古树上的红枣"
)
add_textbox(slide, Inches(6.8), Inches(1.1), Inches(5.8), Inches(5.5),
            right_txt, Pt(10), False, DARK)

# --- SLIDE 21: AGRICULTURAL TOURISM ---
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, LIGHT)
add_textbox(slide, Inches(0.5), Inches(0.3), Inches(12), Inches(0.7),
            "农文旅融合：守望千年·只为枣你", Pt(26), True, PRIMARY)
add_textbox(slide, Inches(0.5), Inches(1.0), Inches(12), Inches(0.6),
            "稷山国家板枣公园 · AAA级景区", Pt(16), False, SECONDARY)
bullets = [
    "● 国家板枣公园——全国第一个经济作物类枣林公园",
    "  1.75万株千年以上古板枣树 + 5万株五百年以上古树",
    "● '守望千年·只为枣你'农文旅融合发展示范园区",
    "  183家市场主体入驻，累计接待游客180余万人次",
    "  10大业态28个系列：吃住行游购娱全要素产业链",
    "● '稷颂'沉浸式文化主题演艺",
    "  展现粮为国本、教民稼穑、江山社稷的深远意义",
    "  业态亮点：星空民宿 | 亲子乐园 | 唐枣温泉 | 非遗街区 | 观光小火车",
    "● 获奖：国家AAA级景区 | 国家农村产业融合发展示范园",
    "● 黄河一号旅游公路为稷山带来更多发展机遇",
]
add_bullet_list(slide, Inches(0.5), Inches(1.6), Inches(12.5), Inches(5.0),
                bullets, Pt(13))

# --- SLIDE 22: CLOSING ---
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, DARK)
add_textbox(slide, Inches(1.0), Inches(2.5), Inches(11), Inches(1.5),
            "结  语", Pt(48), True, ACCENT, align="center")
add_textbox(slide, Inches(1.0), Inches(4.2), Inches(11), Inches(1.0),
            "黄土塬上的古村落，六百年老槐树下，后稷故里的风——",
            Pt(22), False, SECONDARY, align="center")

# --- SLIDE 23: SUMMARY ---
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, BG_DARK)
add_textbox(slide, Inches(0.5), Inches(0.3), Inches(12), Inches(0.7),
            "致敬 · 这片土地", Pt(28), True, ACCENT)
summary = [
    "🏛️ 地理与人文：吕梁山南麓，黄土塬上，后稷故里",
    "🏠 古村落建筑：窑洞 · 四合院 · 关帝庙 · 古井",
    "🌾 农耕文明：塬地 · 梯田 · 沟坝地 · 麦浪金黄",
    "🍜 特色美食：稷山四宝——麻花 · 饼子 · 鸡蛋 · 板枣",
    "⛩️ 名人古迹：大佛寺 · 稷王庙 · 丁庄李家大院 · 7处国宝",
    "🎭 非遗民俗：高跷走兽 · 高台花鼓 · 螺钿漆器 · 金银细工",
    "🎥 传播矩阵：抖音 · 视频号 · 山西DOU是好风光",
    "🌳 农文旅融合：守望千年·只为枣你 · 国家板枣公园",
]
add_bullet_list(slide, Inches(0.5), Inches(1.2), Inches(12.5), Inches(5.5),
                summary, Pt(16))

# --- SLIDE 24: ENDING ---
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, BG_DARK)
add_textbox(slide, Inches(1.0), Inches(2.5), Inches(11), Inches(1.5),
            "稷山四宝  麻花饼子鸡蛋枣", Pt(40), True, ACCENT, align="center")
add_textbox(slide, Inches(1.0), Inches(4.2), Inches(11), Inches(1.0),
            "—— 后稷故里 · 千年古县 · 守望千年 · 只为枣你",
            Pt(20), False, SECONDARY, align="center")

# Save
prs.save(OUT)
print(f"Saved: {OUT}")
print(f"Slides: {len(prs.slides)}")
