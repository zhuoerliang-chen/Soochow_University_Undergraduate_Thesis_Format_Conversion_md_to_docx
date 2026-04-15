from __future__ import annotations

from dataclasses import dataclass


"""
字号（pt）常用对照：
- 小二：18pt
- 二号：22pt
- 小三：15pt
- 三号：16pt
- 四号：14pt
- 小四：12pt
- 五号：10.5pt
- 小五：9pt

说明：
- python-docx 使用 Pt(...) 以“磅（point）”为单位设置字号。
- 行距 line_spacing 使用倍数（例如 1.5 表示 1.5 倍行距）。
- 缩进/页边距使用 Cm(...) 以“厘米”为单位。

本配置文件用于集中管理论文排版关键参数（按当前配置默认值归纳）：
- 封面标题：中文小二（18pt）黑体；外文小二（18pt）Times New Roman
- 摘要：摘要标题小三（15pt）黑体；摘要内容使用正文样式（小四宋体/Times New Roman）
- 正文章节标题：标题加粗；章/节标题小三（15pt）黑体（Heading 1/2）
- 正文：中文宋体、外文 Times New Roman；小四（12pt）；行距 1.5 倍；段前段后 0；首行缩进约 2 字符（0.74cm）
- 目录标题：目录页顶部“目录”二字视为章节标题——小三（15pt）黑体加粗、居中
- 目录条目：章题目四号（14pt）黑体；节题目四号（14pt）宋体；页码用“……”点引导连接（由代码使用制表位 + leader dots + PAGEREF 实现）
- 页面设置：页边距上 3.3cm、下 2.7cm、左 2.5cm、右 2.5cm；左侧装订线 0.5cm
- 页眉页脚：页眉距 2.6cm、页脚距 2.0cm；页眉居中小五（9pt）宋体“苏州大学本科生毕业设计（论文）”并带横线；页脚居中页码
- 代码块：Times New Roman 字体；代码字号 9pt；行号字号 8pt；行号栏宽 0.65cm
"""


@dataclass(frozen=True)
class WordFormatConfig:
    body_east_asia_font: str = "宋体"  # 正文中文字体（宋体）
    body_ascii_font: str = "Times New Roman"  # 正文外文字体（Times New Roman）
    body_font_size_pt: float = 12.0  # 正文字号：小四=12pt
    body_line_spacing: float = 1.5  # 正文行距：1.5 倍
    body_first_line_indent_cm: float = 0.74  # 正文首行缩进：约等于“小四 2 字符”的常用折算（cm）

    title_east_asia_font: str = "黑体"  # 中文标题字体：黑体
    title_ascii_font: str = "Times New Roman"  # 外文标题字体：Times New Roman
    title_font_size_pt: float = 18.0  # 标题字号：小二=18pt

    abstract_heading_text_zh: str = "摘要"  # 中文摘要标题文字
    abstract_heading_text_en: str = "Abstract"  # 外文摘要标题文字
    abstract_heading_font_east_asia: str = "黑体"  # 摘要标题中文字体（不约束摘要内容；摘要内容使用正文样式）
    abstract_heading_font_ascii: str = "Times New Roman"  # 摘要标题外文字体
    abstract_heading_size_pt: float = 15.0  # 摘要标题字号：小三=15pt（不影响摘要内容）

    heading_east_asia_font: str = "黑体"  # 正文章节标题中文字体：黑体（标题加粗由样式设定）
    heading_ascii_font: str = "Times New Roman"  # 正文章节标题外文字体：Times New Roman
    heading_1_size_pt: float = 15.0  # 章节题目字号：小三=15pt（对应“第X章 …”）
    heading_2_size_pt: float = 15.0  # 节题目字号：小三=15pt（对应“X.X …”）
    heading_3_size_pt: float = 15.0  # 三级标题字号（如有需要可调整）
    heading_4_size_pt: float = 15.0  # 四级标题字号（如有需要可调整）

    toc_title_text: str = "目录"  # 目录标题文字（目录页顶部的“目录”二字）
    toc_title_east_asia_font: str = "黑体"  # 目录标题中文字体：黑体（按章节标题要求）
    toc_title_ascii_font: str = "Times New Roman"  # 目录标题外文字体
    toc_title_size_pt: float = 15.0  # 目录标题字号：小三=15pt（视为章节标题）
    toc_title_bold: bool = True  # 目录标题是否加粗

    toc_heading_text: str = "目录"  # Word 内置样式“TOC Heading”的文字（当前实现主要使用 toc_title_* 作为目录页标题）
    toc_heading_east_asia_font: str = "宋体"  # “TOC Heading”样式中文字体
    toc_heading_ascii_font: str = "Times New Roman"  # “TOC Heading”样式外文字体
    toc_heading_size_pt: float = 14.0  # “TOC Heading”样式字号：四号=14pt
    toc_heading_space_after_pt: float = 12.0  # “TOC Heading”样式段后间距（pt）
    toc_level1_east_asia_font: str = "黑体"  # 目录中“章题目”中文字体：黑体（四号）
    toc_level2_east_asia_font: str = "宋体"  # 目录中“节题目”中文字体：宋体（四号）
    toc_level3_east_asia_font: str = "宋体"  # 目录中更深层级条目的中文字体（如有）
    toc_level_font_ascii: str = "Times New Roman"  # 目录条目外文字体
    toc_level1_size_pt: float = 14.0  # 目录章题目字号：四号=14pt
    toc_level2_size_pt: float = 14.0  # 目录节题目字号：四号=14pt
    toc_level3_size_pt: float = 14.0  # 目录更深层级字号：四号=14pt
    toc_level2_left_indent_cm: float = 0.74  # 目录中“节题目”左缩进（与正文首行缩进一致的常用值）

    margin_top_cm: float = 3.3  # 页边距：上 3.3cm
    margin_bottom_cm: float = 2.7  # 页边距：下 2.7cm
    margin_left_cm: float = 2.5  # 页边距：左 2.5cm
    margin_right_cm: float = 2.5  # 页边距：右 2.5cm
    gutter_cm: float = 0.5  # 左侧装订线：0.5cm
    header_distance_cm: float = 2.6  # 页眉距顶端：2.6cm
    footer_distance_cm: float = 2.0  # 页脚距底端：2.0cm

    header_text: str = "苏州大学本科生毕业设计（论文）"  # 页眉内容（居中显示）
    header_font_east_asia: str = "宋体"  # 页眉中文字体：宋体
    header_font_ascii: str = "Times New Roman"  # 页眉外文字体：Times New Roman
    header_footer_font_size_pt: float = 9.0  # 页眉/页脚字号：小五=9pt
    header_bottom_border_size: str = "6"  # 页眉下横线粗细（Word 边框尺寸，常用 6≈0.75pt）
    header_bottom_border_space: str = "1"  # 页眉下横线与文字的间距（Word 边框 space 值）

    code_font_name: str = "Times New Roman"  # 代码块字体（Times New Roman）
    code_font_size_pt: float = 9.0  # 代码块字号（pt）
    code_line_number_font_size_pt: float = 8.0  # 代码块行号字号（pt）
    code_gutter_width_cm: float = 0.65  # 代码块行号栏宽度（cm）


DEFAULT_WORD_FORMAT = WordFormatConfig()
