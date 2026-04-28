#!/usr/bin/env python3
"""生成「和 AI 作朋友——高中生的 AI 工具箱」工作坊簡報 .pptx

對齊 slides.html 內容：5 個 Step（NotebookLM → Napkin → Claude → ChatGPT
→ Gamma）+ AI 原理（含幻覺與輪盤連結）+ 老師示範 + 帶走觀念。
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
import os

# ── 配色（對齊 slides.html / slides-styles.css 主題） ──
BG_DARK = RGBColor(0x1A, 0x1A, 0x2E)
BG_SECTION = RGBColor(0x16, 0x21, 0x3E)
CARD_BG = RGBColor(0x22, 0x2B, 0x45)
ACCENT = RGBColor(0x00, 0xD2, 0xFF)         # primary cyan
ACCENT2 = RGBColor(0xFF, 0x6B, 0x6B)        # coral / red
ACCENT3 = RGBColor(0x4E, 0xCB, 0x71)        # green
ACCENT4 = RGBColor(0xFF, 0xAA, 0x33)        # orange
ACCENT5 = RGBColor(0xCC, 0x66, 0xFF)        # purple
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT_GRAY = RGBColor(0xBB, 0xBB, 0xCC)
SUBTITLE_GRAY = RGBColor(0x99, 0x99, 0xAA)

FONT_TITLE = "Helvetica Neue"
FONT_BODY = "Helvetica Neue"

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
SLIDE_W = prs.slide_width
SLIDE_H = prs.slide_height


# ═══════════════════════════════════════════
# 共用 helpers
# ═══════════════════════════════════════════

def new_slide(bg=BG_DARK):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = bg
    return slide


def add_rect(slide, left, top, width, height, color):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    return shape


def add_text(slide, left, top, width, height, text,
             size=18, color=WHITE, bold=False, align=PP_ALIGN.LEFT,
             font=FONT_BODY, line_spacing=1.4):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Pt(0)
    tf.margin_top = tf.margin_bottom = Pt(0)
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(size)
    p.font.color.rgb = color
    p.font.bold = bold
    p.font.name = font
    p.alignment = align
    p.line_spacing = line_spacing
    return box


def add_bullets(slide, left, top, width, height, items,
                size=22, color=LIGHT_GRAY, bullet_color=ACCENT, spacing=1.6):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Pt(0)
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        bullet = p.add_run()
        bullet.text = "▸  "
        bullet.font.size = Pt(size)
        bullet.font.color.rgb = bullet_color
        bullet.font.name = FONT_BODY
        bullet.font.bold = True
        body = p.add_run()
        body.text = item
        body.font.size = Pt(size)
        body.font.color.rgb = color
        body.font.name = FONT_BODY
        p.line_spacing = spacing
        p.space_after = Pt(size * 0.4)
    return box


def add_callout(slide, left, top, width, height, text,
                color=ACCENT, label=None, size=16):
    """色塊 callout：左側色條 + 半透明底色"""
    add_rect(slide, left, top, width, height, CARD_BG)
    add_rect(slide, left, top, Pt(5), height, color)
    if label:
        add_text(slide, left + Inches(0.2), top + Inches(0.1), width - Inches(0.4), Inches(0.4),
                 label, size=14, color=color, bold=True)
        add_text(slide, left + Inches(0.2), top + Inches(0.55), width - Inches(0.4),
                 height - Inches(0.6), text, size=size, color=WHITE)
    else:
        add_text(slide, left + Inches(0.2), top + Inches(0.2), width - Inches(0.4),
                 height - Inches(0.4), text, size=size, color=WHITE)


# ═══════════════════════════════════════════
# Slide patterns
# ═══════════════════════════════════════════

def title_slide():
    slide = new_slide(BG_SECTION)
    # Decorative line
    add_rect(slide, Inches(5.67), Inches(1.6), Inches(2), Pt(3), ACCENT)
    # Main title
    add_text(slide, Inches(0), Inches(1.9), SLIDE_W, Inches(1.4),
             "和 AI 作朋友", size=64, bold=True, color=WHITE,
             font=FONT_TITLE, align=PP_ALIGN.CENTER)
    add_text(slide, Inches(0), Inches(3.4), SLIDE_W, Inches(0.8),
             "高中生的 AI 工具箱", size=32, color=ACCENT,
             font=FONT_TITLE, align=PP_ALIGN.CENTER)
    add_text(slide, Inches(0), Inches(4.5), SLIDE_W, Inches(0.5),
             "顏永進  ｜  2026.04.30  ｜  雲林", size=20,
             color=SUBTITLE_GRAY, align=PP_ALIGN.CENTER)
    # QR placeholders
    for x, label, url in [
        (Inches(3.5), "簡報", "is.gd/1rLH3E"),
        (Inches(7.8), "教學網頁", "is.gd/tENFQN"),
    ]:
        add_rect(slide, x, Inches(5.4), Inches(1.5), Inches(1.5), CARD_BG)
        add_text(slide, x, Inches(6.0), Inches(1.5), Inches(0.4),
                 "QR", size=18, color=SUBTITLE_GRAY, align=PP_ALIGN.CENTER)
        add_text(slide, x, Inches(7.0), Inches(1.5), Inches(0.3),
                 label, size=14, color=LIGHT_GRAY, align=PP_ALIGN.CENTER)
        add_text(slide, x + Inches(1.6), Inches(5.7), Inches(2.5), Inches(0.4),
                 url, size=18, color=ACCENT, bold=True)
    return slide


def section_divider(label_top, title, subtitle, color=ACCENT):
    slide = new_slide(BG_SECTION)
    # Side colour bar
    add_rect(slide, Inches(0), Inches(0), Inches(0.4), SLIDE_H, color)
    add_text(slide, Inches(1.2), Inches(2.5), Inches(11), Inches(0.6),
             label_top, size=20, color=color, bold=True)
    add_text(slide, Inches(1.2), Inches(3.2), Inches(11), Inches(1.4),
             title, size=54, bold=True, color=WHITE, font=FONT_TITLE)
    add_text(slide, Inches(1.2), Inches(4.7), Inches(11), Inches(0.6),
             subtitle, size=22, color=LIGHT_GRAY)
    return slide


def step_divider(step_num, tool_name, subtitle, color=ACCENT):
    slide = new_slide(BG_SECTION)
    add_rect(slide, Inches(0), Inches(0), Inches(0.4), SLIDE_H, color)
    add_text(slide, Inches(1.2), Inches(2.4), Inches(11), Inches(0.5),
             f"STEP {step_num}", size=20, color=color, bold=True)
    add_text(slide, Inches(1.2), Inches(3.0), Inches(11), Inches(1.6),
             tool_name, size=66, bold=True, color=WHITE, font=FONT_TITLE)
    add_text(slide, Inches(1.2), Inches(4.8), Inches(11), Inches(0.6),
             subtitle, size=24, color=LIGHT_GRAY)
    return slide


def content_slide(step_label, title, color=ACCENT):
    """回傳 slide，呼叫端再加 body 元件"""
    slide = new_slide(BG_DARK)
    # Step badge (small circle with number)
    if step_label:
        badge = slide.shapes.add_shape(
            MSO_SHAPE.OVAL, Inches(0.8), Inches(0.6), Inches(0.55), Inches(0.55))
        badge.fill.solid()
        badge.fill.fore_color.rgb = color
        badge.line.fill.background()
        tf = badge.text_frame
        tf.margin_left = tf.margin_right = Pt(0)
        tf.margin_top = tf.margin_bottom = Pt(0)
        p = tf.paragraphs[0]
        p.text = str(step_label)
        p.font.size = Pt(20)
        p.font.bold = True
        p.font.color.rgb = WHITE
        p.alignment = PP_ALIGN.CENTER
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    # Title
    add_text(slide, Inches(1.6) if step_label else Inches(0.8), Inches(0.65),
             Inches(11), Inches(0.7),
             title, size=32, bold=True, color=WHITE, font=FONT_TITLE)
    # Underline accent
    add_rect(slide, Inches(0.8), Inches(1.45), Inches(11.5), Pt(2), color)
    return slide


# ═══════════════════════════════════════════
# 1. 封面
# ═══════════════════════════════════════════
title_slide()


# ═══════════════════════════════════════════
# 2. 今天要做什麼？
# ═══════════════════════════════════════════
slide = content_slide(None, "今天要做什麼？", ACCENT)
add_text(slide, Inches(0.8), Inches(1.8), Inches(11.5), Inches(0.5),
         "用 AI 工具，從零完成一份台灣史小論文",
         size=22, color=LIGHT_GRAY)
add_text(slide, Inches(0.8), Inches(2.4), Inches(11.5), Inches(0.6),
         "主題：為什麼日本人選雲林蓋糖廠？",
         size=26, bold=True, color=ACCENT, font=FONT_TITLE)

# 5-step table
steps = [
    ("1", "NotebookLM", "研究筆記（含出處）", ACCENT),
    ("2", "Napkin", "因果關係圖", ACCENT2),
    ("3", "Claude", "小論文 Word 檔", ACCENT3),
    ("4", "ChatGPT", "改好的反思段落", ACCENT4),
    ("5", "Gamma 或 Claude", "報告簡報", ACCENT5),
]
for i, (num, tool, output, color) in enumerate(steps):
    y = Inches(3.4 + i * 0.72)
    add_rect(slide, Inches(0.8), y, Inches(11.5), Inches(0.6), CARD_BG)
    add_rect(slide, Inches(0.8), y, Pt(5), Inches(0.6), color)
    add_text(slide, Inches(1.1), y + Inches(0.1), Inches(0.6), Inches(0.4),
             num, size=22, bold=True, color=color, font=FONT_TITLE)
    add_text(slide, Inches(1.9), y + Inches(0.13), Inches(4.5), Inches(0.4),
             tool, size=20, bold=True, color=WHITE, font=FONT_TITLE)
    add_text(slide, Inches(6.5), y + Inches(0.13), Inches(5.5), Inches(0.4),
             output, size=18, color=LIGHT_GRAY)


# ═══════════════════════════════════════════
# Step 1: NotebookLM
# ═══════════════════════════════════════════
step_divider(1, "NotebookLM", "餵素材、問問題、產出研究筆記", ACCENT)

# Step 1 - 怎麼用？
slide = content_slide("1", "NotebookLM 怎麼用？", ACCENT)
# 兩欄：上傳素材 / 問問題
add_text(slide, Inches(0.8), Inches(1.9), Inches(5.5), Inches(0.5),
         "上傳素材", size=22, bold=True, color=ACCENT, font=FONT_TITLE)
add_text(slide, Inches(0.8), Inches(2.5), Inches(5.5), Inches(2),
         "把文章存成 PDF，上傳到 NotebookLM\n\n素材清單見教學網頁底部",
         size=18, color=LIGHT_GRAY, line_spacing=1.5)
add_text(slide, Inches(7.0), Inches(1.9), Inches(5.5), Inches(0.5),
         "問問題", size=22, bold=True, color=ACCENT, font=FONT_TITLE)
add_bullets(slide, Inches(7.0), Inches(2.5), Inches(5.7), Inches(4), [
    "日本人為什麼選雲林？",
    "戊戌大水災跟糖廠的關係？",
    "虎尾 vs. 橋頭有什麼不同？",
    "小論文大綱可以怎麼寫？",
], size=18, bullet_color=ACCENT)

# Step 1 - 打包指令
slide = content_slide("1", "最後一步：請它幫你打包", ACCENT)
add_text(slide, Inches(0.8), Inches(1.8), Inches(11.5), Inches(0.5),
         "問完所有問題後，貼這段指令：", size=20, color=LIGHT_GRAY)
add_callout(slide, Inches(0.8), Inches(2.5), Inches(11.5), Inches(2.5),
            "請把我們剛才所有對話整合成一份「研究筆記」：\n\n"
            "一、研究問題\n"
            "二、關鍵發現（每點標明出處）\n"
            "三、比較分析\n"
            "四、建議大綱\n"
            "五、參考資料清單",
            color=ACCENT, size=16)
add_callout(slide, Inches(0.8), Inches(5.4), Inches(11.5), Inches(0.9),
            "▸  只要複製這一份，就是等一下餵給 Claude 的全部素材",
            color=ACCENT3, size=18)


# ═══════════════════════════════════════════
# Step 2: Napkin
# ═══════════════════════════════════════════
step_divider(2, "Napkin", "把文字變成圖表", ACCENT2)

slide = content_slide("2", "Napkin——畫一張圖", ACCENT2)
add_text(slide, Inches(0.8), Inches(1.8), Inches(11.5), Inches(0.5),
         "把研究重點貼進 napkin.ai，自動生成圖表",
         size=20, color=LIGHT_GRAY)
add_callout(slide, Inches(0.8), Inches(2.5), Inches(11.5), Inches(3.3),
            "日本選擇在雲林虎尾設立糖廠的原因：\n\n"
            "1. 天災創造機會：1897 戊戌大水災沖出大量溪埔地\n"
            "2. 自然條件適合：雲林平原日照充足、雨量適中\n"
            "3. 政策推動：1902 年糖業獎勵規則\n"
            "4. 企業家眼光：鈴木藤三郎，1906 年設廠\n"
            "5. 國家需求：日俄戰爭後需要糖的自給自足",
            color=ACCENT2, size=15)
add_text(slide, Inches(0.8), Inches(6.1), Inches(11.5), Inches(0.5),
         "→ 產出的圖放進小論文或簡報",
         size=18, color=ACCENT, bold=True)


# ═══════════════════════════════════════════
# Step 3: Claude
# ═══════════════════════════════════════════
step_divider(3, "Claude", "生成小論文 Word 檔", ACCENT3)

# Step 3 - 怎麼用
slide = content_slide("3", "Claude——生成 Word", ACCENT3)
add_text(slide, Inches(0.8), Inches(1.8), Inches(11.5), Inches(0.5),
         "打開 claude.ai（免費帳號），依序貼入：",
         size=20, color=LIGHT_GRAY)
order = [
    ("1", "小論文格式規範（見教學網頁附錄）", ACCENT3),
    ("2", "Step 1 的研究筆記（含出處）", ACCENT3),
    ("3", "指令：「整理成小論文，輸出 Word 檔」", ACCENT3),
]
for i, (num, text, color) in enumerate(order):
    y = Inches(2.6 + i * 0.85)
    badge = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.8), y, Inches(0.45), Inches(0.45))
    badge.fill.solid(); badge.fill.fore_color.rgb = color; badge.line.fill.background()
    tf = badge.text_frame
    tf.margin_left = tf.margin_right = Pt(0); tf.margin_top = tf.margin_bottom = Pt(0)
    tf.paragraphs[0].text = num
    tf.paragraphs[0].font.size = Pt(18); tf.paragraphs[0].font.bold = True
    tf.paragraphs[0].font.color.rgb = WHITE
    tf.paragraphs[0].alignment = PP_ALIGN.CENTER
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    add_text(slide, Inches(1.5), y + Inches(0.05), Inches(11), Inches(0.5),
             text, size=20, color=WHITE)
add_callout(slide, Inches(0.8), Inches(5.6), Inches(11.5), Inches(0.9),
            "▸  Claude 生成 Artifact → 點「下載」→ 拿到 .docx",
            color=ACCENT4, size=18)

# Step 3 - 下載後
slide = content_slide("3", "下載後一定要做的事", ACCENT3)
checks = [
    ("換成你的口吻", "教授看得出 AI 味", ACCENT),
    ("檢查事實", "人名、年份、事件", ACCENT4),
    ("反思先留著", "下一步會改", ACCENT3),
]
for i, (title, desc, color) in enumerate(checks):
    x = Inches(0.8 + i * 4.0)
    add_rect(slide, x, Inches(2.5), Inches(3.7), Inches(3.5), CARD_BG)
    add_rect(slide, x, Inches(2.5), Inches(3.7), Pt(4), color)
    add_text(slide, x, Inches(3.0), Inches(3.7), Inches(0.6),
             title, size=24, bold=True, color=WHITE,
             font=FONT_TITLE, align=PP_ALIGN.CENTER)
    add_text(slide, x, Inches(4.0), Inches(3.7), Inches(0.4),
             desc, size=16, color=LIGHT_GRAY, align=PP_ALIGN.CENTER)


# ═══════════════════════════════════════════
# Step 4: ChatGPT
# ═══════════════════════════════════════════
step_divider(4, "ChatGPT", "用 AI 當教練改反思", ACCENT4)

# 好/爛 prompt
slide = content_slide("4", "好 prompt vs. 爛 prompt", ACCENT4)
add_callout(slide, Inches(0.8), Inches(1.9), Inches(11.5), Inches(1.0),
            "「幫我寫一段學習歷程反思」  →  罐頭、流水帳",
            color=ACCENT2, label="✗  爛 prompt", size=16)
add_callout(slide, Inches(0.8), Inches(3.1), Inches(11.5), Inches(2.7),
            "我住雲林，常經過虎尾糖廠，以前只覺得是老地方。研究後才知道日本人選雲林"
            "不是巧合，是一場水災沖出了空地。幫我用「從不在意到驚訝」的轉變寫反思，"
            "像高中生寫的，300 字以內。",
            color=ACCENT3, label="✓  好 prompt", size=15)
add_callout(slide, Inches(0.8), Inches(6.0), Inches(11.5), Inches(0.9),
            "你給 AI 多少「你的樣子」，它就回給你多少「你的樣子」",
            color=ACCENT, size=18)

# Coach
slide = content_slide("4", "用 AI 改，不是用 AI 寫", ACCENT4)
add_text(slide, Inches(0.8), Inches(1.8), Inches(11.5), Inches(0.5),
         "拿到初稿後，繼續用 AI 當教練：", size=20, color=LIGHT_GRAY)
prompts = [
    ("這段反思沒有連結到研究問題，幫我指出哪裡可以改。", ACCENT),
    ("給我三種角度：學到什麼 / 對家鄉看法改變 / 還想探索什麼", ACCENT4),
    ("這句話怎麼改更有力？", ACCENT3),
]
for i, (text, color) in enumerate(prompts):
    y = Inches(2.6 + i * 1.05)
    add_callout(slide, Inches(0.8), y, Inches(11.5), Inches(0.85),
                f"「{text}」", color=color, size=16)
add_text(slide, Inches(0.8), Inches(6.1), Inches(11.5), Inches(0.5),
         "至少來回三輪，改好後貼回 Word 檔",
         size=20, bold=True, color=WHITE)


# ═══════════════════════════════════════════
# Step 5: Gamma
# ═══════════════════════════════════════════
step_divider(5, "Gamma", "做簡報", ACCENT5)

slide = content_slide("5", "Gamma——做簡報", ACCENT5)
add_text(slide, Inches(0.8), Inches(1.8), Inches(11.5), Inches(0.5),
         "把 Word 內容貼進 gamma.app，輸入：",
         size=20, color=LIGHT_GRAY)
add_callout(slide, Inches(0.8), Inches(2.5), Inches(11.5), Inches(2.0),
            "高中歷史課程學習成果報告簡報。\n"
            "標題：（你的題目）\n"
            "一、研究動機  二、研究發現  三、學習反思",
            color=ACCENT5, size=16)
add_text(slide, Inches(0.8), Inches(4.8), Inches(11.5), Inches(0.5),
         "→ 30 秒生成簡報",
         size=20, bold=True, color=ACCENT)
add_text(slide, Inches(0.8), Inches(5.4), Inches(11.5), Inches(0.5),
         "→ 記得刪掉 AI 亂加的內容、補上自己的話",
         size=20, bold=True, color=ACCENT)
add_text(slide, Inches(0.8), Inches(6.2), Inches(11.5), Inches(0.5),
         "也可以用 claude.ai 直接生成 .pptx",
         size=18, color=SUBTITLE_GRAY)


# ═══════════════════════════════════════════
# AI 怎麼辦到的？
# ═══════════════════════════════════════════
section_divider("揭開黑盒子", "AI 怎麼辦到的？", "輸入 → 函數 → 輸出", ACCENT5)

# AI = 函數
slide = content_slide(None, "AI 的本質是「函數」", ACCENT5)
add_text(slide, Inches(0.8), Inches(1.9), Inches(5.5), Inches(0.5),
         "國中數學學過：", size=22, color=LIGHT_GRAY)
add_text(slide, Inches(0.8), Inches(2.5), Inches(5.5), Inches(0.6),
         "輸入  →  規則  →  輸出",
         size=28, bold=True, color=ACCENT, font=FONT_TITLE)
add_text(slide, Inches(0.8), Inches(3.4), Inches(5.5), Inches(0.5),
         "AI 做的事情一模一樣", size=22, color=LIGHT_GRAY)

# 右側舉例
examples = [
    ("手寫辨識", "圖片 → f → 「這是 5」（92%）"),
    ("貓狗辨識", "照片 → f → 貓 90% / 狗 10%"),
    ("文字生成", "提示詞 → f → 生成的文字"),
]
for i, (name, desc) in enumerate(examples):
    y = Inches(2.0 + i * 1.0)
    add_rect(slide, Inches(7.0), y, Inches(5.5), Inches(0.85), CARD_BG)
    add_rect(slide, Inches(7.0), y, Pt(4), Inches(0.85), ACCENT5)
    add_text(slide, Inches(7.2), y + Inches(0.1), Inches(5.2), Inches(0.4),
             name, size=18, bold=True, color=WHITE, font=FONT_TITLE)
    add_text(slide, Inches(7.2), y + Inches(0.45), Inches(5.2), Inches(0.4),
             desc, size=14, color=LIGHT_GRAY)
add_callout(slide, Inches(0.8), Inches(5.6), Inches(11.5), Inches(0.9),
            "不管多厲害，骨子裡都是：輸入 → 函數 → 輸出",
            color=ACCENT3, size=20)

# 接龍函數
slide = content_slide(None, "生成式 AI 是「會接龍的函數」", ACCENT5)
add_text(slide, Inches(0.8), Inches(1.9), Inches(11.5), Inches(0.5),
         "每次只猜下一個字，再把輸出接回輸入",
         size=22, color=LIGHT_GRAY)
add_callout(slide, Inches(0.8), Inches(2.7), Inches(11.5), Inches(3.0),
            'f("什麼是臺灣最美的風景？")           →  臺\n'
            'f("什麼是臺灣最美的風景？臺")         →  灣\n'
            'f("什麼是臺灣最美的風景？臺灣")       →  最\n'
            'f("什麼是臺灣最美的風景？臺灣最")     →  美\n'
            '   ……\n'
            'f("什麼是臺灣最美的風景？臺灣最美的風景是")  →  人',
            color=ACCENT, size=15)
add_text(slide, Inches(0.8), Inches(6.1), Inches(11.5), Inches(0.5),
         "就像 LINE 鍵盤自動建議下一個字——AI 做一樣的事，只是猜得非常準",
         size=18, color=LIGHT_GRAY)

# 為什麼同一題不同答案 / 輪盤
slide = content_slide(None, "為什麼同一題會得到不同答案？", ACCENT5)
add_text(slide, Inches(0.8), Inches(1.9), Inches(11.5), Inches(0.5),
         "AI 不是每次都選機率最高的字",
         size=24, bold=True, color=WHITE, font=FONT_TITLE)
add_text(slide, Inches(0.8), Inches(2.6), Inches(11.5), Inches(0.5),
         "像轉「輪盤」——機率高的比較容易中，但不是一定",
         size=20, color=LIGHT_GRAY)
add_callout(slide, Inches(0.8), Inches(3.6), Inches(11.5), Inches(2.5),
            "溫度低 → 回答比較固定\n\n"
            "溫度高 → 回答有創意，但比較容易出錯",
            color=ACCENT4, label="🌡  溫度參數（Temperature）控制隨機性",
            size=18)
add_text(slide, Inches(0.8), Inches(6.4), Inches(11.5), Inches(0.5),
         "→ 這個「猜字+輪盤」也是幻覺的根源（下一張）",
         size=16, color=ACCENT, bold=True)

# Token
slide = content_slide(None, "Token：AI 眼中的世界", ACCENT5)
add_text(slide, Inches(0.8), Inches(1.9), Inches(11.5), Inches(0.5),
         "AI 看到的不是文字，是 Token（碎片）",
         size=24, bold=True, color=WHITE, font=FONT_TITLE)
add_callout(slide, Inches(0.8), Inches(2.8), Inches(11.5), Inches(2.5),
            "▸  輸入「虎尾糖廠」  →  看它被切成幾個 token\n"
            "▸  輸入「314159」  →  數字的切法跟你想的完全不一樣",
            color=ACCENT, label="打開 platform.openai.com/tokenizer 試試", size=17)
add_text(slide, Inches(0.8), Inches(5.8), Inches(11.5), Inches(0.5),
         "這就是為什麼 AI 有時候數學會算錯——它看到的不是數字，是碎片",
         size=18, color=ACCENT2, bold=True)

# 訓練三階段
slide = content_slide(None, "AI 怎麼被訓練出來的？", ACCENT5)
stages = [
    ("1", "預訓練", "讀了整個網路的文字\n→ 學會語言模式", ACCENT),
    ("2", "微調", "針對任務練習\n→ 學會當助理", ACCENT4),
    ("3", "RLHF", "人類回饋好壞\n→ 學會避免有害回應", ACCENT3),
]
for i, (num, title, desc, color) in enumerate(stages):
    x = Inches(0.8 + i * 4.0)
    add_rect(slide, x, Inches(2.3), Inches(3.7), Inches(4.2), CARD_BG)
    add_rect(slide, x, Inches(2.3), Inches(3.7), Pt(4), color)
    add_text(slide, x, Inches(2.7), Inches(3.7), Inches(1.2),
             num, size=56, bold=True, color=color,
             font=FONT_TITLE, align=PP_ALIGN.CENTER)
    add_text(slide, x, Inches(4.3), Inches(3.7), Inches(0.6),
             title, size=24, bold=True, color=WHITE,
             font=FONT_TITLE, align=PP_ALIGN.CENTER)
    add_text(slide, x, Inches(5.1), Inches(3.7), Inches(1.2),
             desc, size=15, color=LIGHT_GRAY,
             align=PP_ALIGN.CENTER, line_spacing=1.5)

# 幻覺（連結回輪盤）
slide = content_slide(None, "幻覺——AI 也會騙你", ACCENT2)
add_text(slide, Inches(0.8), Inches(1.9), Inches(11.5), Inches(0.5),
         "AI 有時候會一本正經地「編故事」，這叫幻覺",
         size=24, bold=True, color=WHITE, font=FONT_TITLE)
add_callout(slide, Inches(0.8), Inches(2.8), Inches(11.5), Inches(2.0),
            "還記得剛才的「輪盤」嗎？AI 不是在「回憶事實」，是在「猜下一個最可能的字」"
            "——它沒辦法分辨「真的」跟「聽起來像真的」",
            color=ACCENT4, label="❓  為什麼會有幻覺？", size=16)
add_callout(slide, Inches(0.8), Inches(5.1), Inches(11.5), Inches(1.9),
            "分別打開 ChatGPT、Claude、Gemini、Grok，貼上完全相同的問題：\n"
            "「請描述 1920 年雲林農民組合抗議虎尾糖廠壓低甘蔗收購價格的經過」",
            color=ACCENT2, label="實驗：用同一個問題測四個 AI", size=15)

# 幻覺實驗結論
slide = content_slide(None, "觀察重點", ACCENT2)
add_bullets(slide, Inches(0.8), Inches(1.9), Inches(11.5), Inches(2.5), [
    "四個 AI 講的是同一件事嗎？",
    "有沒有 AI 承認「我不確定」？",
    "跟 NotebookLM 的素材比對——哪些是真的？哪些是編的？",
], size=20, bullet_color=ACCENT2)
add_callout(slide, Inches(0.8), Inches(4.7), Inches(11.5), Inches(2.0),
            "AI 不是知識庫，是語言模型。\n"
            "所以才需要 NotebookLM——有出處的回答才可信。",
            color=ACCENT3, label="結論", size=20)

# 不該丟給 AI
slide = content_slide(None, "什麼東西不該丟給 AI？", ACCENT2)
add_text(slide, Inches(0.8), Inches(1.9), Inches(11.5), Inches(0.5),
         "你打進去的每一個字，都會離開你的電腦",
         size=22, color=LIGHT_GRAY)
add_callout(slide, Inches(0.8), Inches(2.7), Inches(5.6), Inches(3.5),
            "▸  日記、心情、家裡狀況\n\n"
            "▸  同學個資、輔導紀錄\n\n"
            "▸  未公開的競賽作品",
            color=ACCENT2, label="不要丟", size=17)
add_callout(slide, Inches(6.7), Inches(2.7), Inches(5.6), Inches(3.5),
            "「這段話被截圖貼到網路上，\n我會崩潰嗎？」\n\n"
            "會  →  別丟雲端",
            color=ACCENT3, label="判準", size=17)


# ═══════════════════════════════════════════
# 你現在手上有什麼？
# ═══════════════════════════════════════════
slide = content_slide(None, "你現在手上有什麼？", ACCENT)
products = [
    ("1", "研究筆記（含出處）", "NotebookLM", ACCENT),
    ("2", "因果關係圖", "Napkin", ACCENT2),
    ("3", "小論文 Word 檔", "Claude", ACCENT3),
    ("4", "反思定稿（貼回 Word）", "ChatGPT", ACCENT4),
    ("5", "報告簡報", "Gamma 或 Claude", ACCENT5),
]
for i, (num, what, tool, color) in enumerate(products):
    y = Inches(2.0 + i * 0.8)
    add_rect(slide, Inches(0.8), y, Inches(11.5), Inches(0.65), CARD_BG)
    add_rect(slide, Inches(0.8), y, Pt(5), Inches(0.65), color)
    add_text(slide, Inches(1.1), y + Inches(0.13), Inches(0.5), Inches(0.4),
             num, size=22, bold=True, color=color, font=FONT_TITLE)
    add_text(slide, Inches(1.8), y + Inches(0.15), Inches(6.5), Inches(0.4),
             what, size=20, bold=True, color=WHITE, font=FONT_TITLE)
    add_text(slide, Inches(8.3), y + Inches(0.17), Inches(4.0), Inches(0.4),
             tool, size=16, color=LIGHT_GRAY, align=PP_ALIGN.RIGHT)


# ═══════════════════════════════════════════
# 老師示範作品
# ═══════════════════════════════════════════
slide = content_slide(None, "老師示範：跑完流程長這樣", ACCENT3)
add_text(slide, Inches(0.8), Inches(1.85), Inches(11.5), Inches(0.5),
         "以「為什麼是雲林？」為題跑完整套流程，產出：",
         size=20, color=LIGHT_GRAY)
demos = [
    ("⏱  互動時間軸", "虎尾糖廠時間軸.html", ACCENT),
    ("📄  小論文 Word 檔", "Step 3 + 4 定稿", ACCENT4),
    ("🎞  Gamma 簡報", "Step 5 路徑 A", ACCENT3),
    ("🎞  Claude 簡報", "Step 5 路徑 B", ACCENT5),
]
for i, (title, desc, color) in enumerate(demos):
    col = i % 2
    row = i // 2
    x = Inches(0.8 + col * 6.0)
    y = Inches(2.6 + row * 1.7)
    add_rect(slide, x, y, Inches(5.7), Inches(1.5), CARD_BG)
    add_rect(slide, x, y, Inches(5.7), Pt(4), color)
    add_text(slide, x + Inches(0.3), y + Inches(0.25), Inches(5.3), Inches(0.5),
             title, size=22, bold=True, color=WHITE, font=FONT_TITLE)
    add_text(slide, x + Inches(0.3), y + Inches(0.85), Inches(5.3), Inches(0.5),
             desc, size=15, color=LIGHT_GRAY)
add_callout(slide, Inches(0.8), Inches(6.3), Inches(11.5), Inches(0.7),
            "三小時後，你手上也會有類似的東西——而且是你自己選的題目、你自己的口吻",
            color=ACCENT3, size=16)


# ═══════════════════════════════════════════
# 三個帶走的觀念
# ═══════════════════════════════════════════
slide = content_slide(None, "三個帶走的觀念", ACCENT)
takeaways = [
    ("1", "AI 幫你加速，不是幫你代寫",
     "最後掛名的是你", ACCENT),
    ("2", "好 prompt 比好工具重要",
     "你給 AI 多少「你的樣子」，它就回給你多少", ACCENT4),
    ("3", "一定要自己檢查",
     "四個 AI 可能給你四個版本的「歷史」", ACCENT3),
]
for i, (num, title, desc, color) in enumerate(takeaways):
    y = Inches(2.0 + i * 1.55)
    add_rect(slide, Inches(0.8), y, Inches(11.5), Inches(1.3), CARD_BG)
    add_rect(slide, Inches(0.8), y, Pt(5), Inches(1.3), color)
    badge = slide.shapes.add_shape(
        MSO_SHAPE.OVAL, Inches(1.1), y + Inches(0.3), Inches(0.7), Inches(0.7))
    badge.fill.solid(); badge.fill.fore_color.rgb = color; badge.line.fill.background()
    tf = badge.text_frame
    tf.margin_left = tf.margin_right = Pt(0); tf.margin_top = tf.margin_bottom = Pt(0)
    tf.paragraphs[0].text = num
    tf.paragraphs[0].font.size = Pt(28); tf.paragraphs[0].font.bold = True
    tf.paragraphs[0].font.color.rgb = WHITE
    tf.paragraphs[0].alignment = PP_ALIGN.CENTER
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    add_text(slide, Inches(2.1), y + Inches(0.2), Inches(10), Inches(0.6),
             title, size=24, bold=True, color=WHITE, font=FONT_TITLE)
    add_text(slide, Inches(2.1), y + Inches(0.75), Inches(10), Inches(0.4),
             desc, size=16, color=LIGHT_GRAY)


# ═══════════════════════════════════════════
# 謝謝
# ═══════════════════════════════════════════
slide = new_slide(BG_DARK)
add_rect(slide, Inches(5.67), Inches(2.8), Inches(2), Pt(3), ACCENT)
add_text(slide, Inches(0), Inches(3.0), SLIDE_W, Inches(1.4),
         "謝謝！", size=72, bold=True, color=WHITE,
         font=FONT_TITLE, align=PP_ALIGN.CENTER)
add_text(slide, Inches(0), Inches(4.5), SLIDE_W, Inches(0.6),
         "現在開始跟 AI 作朋友吧",
         size=24, color=ACCENT, align=PP_ALIGN.CENTER)
add_text(slide, Inches(0), Inches(5.8), SLIDE_W, Inches(0.5),
         "教學網頁：letranger.github.io/2026-04-30-AI-work/",
         size=16, color=SUBTITLE_GRAY, align=PP_ALIGN.CENTER)


# ═══════════════════════════════════════════
output_path = os.path.join(os.path.dirname(__file__), "和AI作朋友-工作坊簡報.pptx")
prs.save(output_path)
print(f"✅ 簡報已儲存：{output_path}")
print(f"📊 共 {len(prs.slides)} 張投影片")
