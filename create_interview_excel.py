#!/usr/bin/env python3
"""
Generate an Excel (.xlsx) file for interview evaluation system.
Design Philosophy: Approachable Luxury
Color Strategy: Monochromatic Morandi watercolor wash tones
Layout: Card-based design with layered elements
Style: Editorial sketch style (人文类)

Uses only Python standard library (zipfile + xml) to build Open XML spreadsheet.
"""

import zipfile
import os

OUTPUT_FILE = "面试评估表.xlsx"

# ─── Morandi Watercolor Palette (Monochromatic warm gray-rose) ────────────────
# Base: A warm dusty rose/mauve Morandi tone, expanded monochromatically
#
# Darkest (text/accent):     #5B4E51 - deep warm charcoal
# Dark (headers):            #7D6B6E - muted mauve-brown
# Medium (card borders):     #A89396 - soft rose-gray
# Light (card fill):         #D4C4C7 - pale watercolor wash
# Lighter (alt rows):        #E8DFE1 - whisper blush
# Lightest (background):     #F5F0F1 - barely-there warmth
# Accent warm:               #C9A9A0 - terracotta watercolor
# Accent cool:               #B8C4C9 - sage mist (complementary)

MORANDI = {
    "darkest":  "FF5B4E51",
    "dark":     "FF7D6B6E",
    "medium":   "FFA89396",
    "light":    "FFD4C4C7",
    "lighter":  "FFE8DFE1",
    "lightest": "FFF5F0F1",
    "accent":   "FFC9A9A0",
    "cool":     "FFB8C4C9",
    "white":    "FFFFFFFF",
    "cream":    "FFFAF8F7",
}

# ─── XML Infrastructure ──────────────────────────────────────────────────────

CONTENT_TYPES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>'''

RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>'''

WORKBOOK_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>'''

WORKBOOK = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="\u603b\u8868" sheetId="1" r:id="rId1"/>
    <sheet name="\u5355\u4eba\u8bc4\u4ef7\u9875" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>'''



# ─── Styles: Morandi Watercolor + Card-Based Layered Design ──────────────────
# 
# Design concept:
# - "Card" effect via medium borders on content blocks with soft fill
# - Layered depth: header layer (dark) → subheader (medium) → content (light/lighter alternating)
# - Typography: serif-like feel via Georgia for titles, clean sans for body
# - Generous whitespace (row heights) for breathing room
# - Monochromatic progression creates visual hierarchy without competing hues
#
# Font indices:
#   0 = Body (11pt, dark text)
#   1 = Body bold (11pt, dark text, bold)
#   2 = Title (16pt, darkest, bold)
#   3 = Subtitle (12pt, medium tone)
#   4 = Header (11pt, white/cream text, bold)
#   5 = Small/meta (9pt, medium tone)
#   6 = Accent bold (11pt, accent color, bold)
#
# Fill indices:
#   0 = none
#   1 = gray125 (required by spec)
#   2 = darkest (header bar)
#   3 = dark (sub-header)
#   4 = medium (card border tone)
#   5 = light (card fill)
#   6 = lighter (alt row)
#   7 = lightest (page background)
#   8 = cream (title area)
#   9 = accent (highlight)
#   10 = cool (subtle accent)
#   11 = white
#
# Border indices:
#   0 = none
#   1 = thin all (medium color) - card outline
#   2 = bottom only thick (dark) - title underline
#   3 = thin bottom (light) - subtle row separator
#   4 = medium all (dark) - card frame
#   5 = dashed bottom (medium) - editorial sketch line

STYLES = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="7">
    <font><sz val="11"/><color rgb="{MORANDI["darkest"]}"/><name val="Segoe UI"/></font>
    <font><b/><sz val="11"/><color rgb="{MORANDI["darkest"]}"/><name val="Segoe UI"/></font>
    <font><b/><sz val="16"/><color rgb="{MORANDI["darkest"]}"/><name val="Georgia"/></font>
    <font><sz val="12"/><color rgb="{MORANDI["medium"]}"/><name val="Georgia"/></font>
    <font><b/><sz val="11"/><color rgb="{MORANDI["cream"]}"/><name val="Segoe UI"/></font>
    <font><sz val="9"/><color rgb="{MORANDI["medium"]}"/><name val="Segoe UI"/></font>
    <font><b/><sz val="11"/><color rgb="{MORANDI["accent"]}"/><name val="Segoe UI"/></font>
  </fonts>
  <fills count="12">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="{MORANDI["darkest"]}"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="{MORANDI["dark"]}"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="{MORANDI["medium"]}"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="{MORANDI["light"]}"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="{MORANDI["lighter"]}"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="{MORANDI["lightest"]}"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="{MORANDI["cream"]}"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="{MORANDI["accent"]}"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="{MORANDI["cool"]}"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="{MORANDI["white"]}"/></patternFill></fill>
  </fills>
  <borders count="6">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border>
      <left style="thin"><color rgb="{MORANDI["medium"]}"/></left>
      <right style="thin"><color rgb="{MORANDI["medium"]}"/></right>
      <top style="thin"><color rgb="{MORANDI["medium"]}"/></top>
      <bottom style="thin"><color rgb="{MORANDI["medium"]}"/></bottom>
      <diagonal/>
    </border>
    <border>
      <left/><right/><top/>
      <bottom style="medium"><color rgb="{MORANDI["dark"]}"/></bottom>
      <diagonal/>
    </border>
    <border>
      <left/><right/><top/>
      <bottom style="thin"><color rgb="{MORANDI["light"]}"/></bottom>
      <diagonal/>
    </border>
    <border>
      <left style="medium"><color rgb="{MORANDI["dark"]}"/></left>
      <right style="medium"><color rgb="{MORANDI["dark"]}"/></right>
      <top style="medium"><color rgb="{MORANDI["dark"]}"/></top>
      <bottom style="medium"><color rgb="{MORANDI["dark"]}"/></bottom>
      <diagonal/>
    </border>
    <border>
      <left/><right/><top/>
      <bottom style="dashed"><color rgb="{MORANDI["medium"]}"/></bottom>
      <diagonal/>
    </border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="16">
    <!-- 0: Default/page background -->
    <xf numFmtId="0" fontId="0" fillId="7" borderId="0" xfId="0" applyFill="1"/>
    <!-- 1: Title (large serif, cream bg, underline border) -->
    <xf numFmtId="0" fontId="2" fillId="8" borderId="2" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="left" vertical="center"/>
    </xf>
    <!-- 2: Subtitle (medium serif, no fill) -->
    <xf numFmtId="0" fontId="3" fillId="7" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1">
      <alignment horizontal="left" vertical="center"/>
    </xf>
    <!-- 3: Card header (dark fill, white text, card border) -->
    <xf numFmtId="0" fontId="4" fillId="2" borderId="4" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center" wrapText="1"/>
    </xf>
    <!-- 4: Card body - normal row (light fill, thin border) -->
    <xf numFmtId="0" fontId="0" fillId="5" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center" wrapText="1"/>
    </xf>
    <!-- 5: Card body - alt row (lighter fill, thin border) -->
    <xf numFmtId="0" fontId="0" fillId="6" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center" wrapText="1"/>
    </xf>
    <!-- 6: Card body - left aligned (light fill) -->
    <xf numFmtId="0" fontId="0" fillId="5" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="left" vertical="center" wrapText="1" indent="1"/>
    </xf>
    <!-- 7: Card body - left aligned alt (lighter fill) -->
    <xf numFmtId="0" fontId="0" fillId="6" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="left" vertical="center" wrapText="1" indent="1"/>
    </xf>
    <!-- 8: Accent bold cell (for totals/highlights) -->
    <xf numFmtId="0" fontId="6" fillId="8" borderId="4" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center"/>
    </xf>
    <!-- 9: Section label (bold, sketch underline) -->
    <xf numFmtId="0" fontId="1" fillId="7" borderId="5" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="left" vertical="center"/>
    </xf>
    <!-- 10: Meta/small text -->
    <xf numFmtId="0" fontId="5" fillId="7" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1">
      <alignment horizontal="left" vertical="center"/>
    </xf>
    <!-- 11: Card sub-header (medium fill, dark border) -->
    <xf numFmtId="0" fontId="1" fillId="4" borderId="4" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center" wrapText="1"/>
    </xf>
    <!-- 12: Info field label (bold, subtle bottom) -->
    <xf numFmtId="0" fontId="1" fillId="8" borderId="3" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="left" vertical="center"/>
    </xf>
    <!-- 13: Info field value (normal, subtle bottom) -->
    <xf numFmtId="0" fontId="0" fillId="8" borderId="3" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="left" vertical="center"/>
    </xf>
    <!-- 14: Score input cell (white fill, card border, center) -->
    <xf numFmtId="0" fontId="1" fillId="11" borderId="4" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center"/>
    </xf>
    <!-- 15: Cool accent cell (sage mist) -->
    <xf numFmtId="0" fontId="0" fillId="10" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center"/>
    </xf>
  </cellXfs>
</styleSheet>'''



# ─── Shared Strings ──────────────────────────────────────────────────────────

shared_strings = []
string_index_map = {}

def get_string_index(s):
    """Get or create shared string index."""
    if s not in string_index_map:
        string_index_map[s] = len(shared_strings)
        shared_strings.append(s)
    return string_index_map[s]


def build_shared_strings_xml():
    """Build the sharedStrings.xml content."""
    items = ""
    for s in shared_strings:
        escaped = s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        items += f'  <si><t>{escaped}</t></si>\n'
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="{len(shared_strings)}" uniqueCount="{len(shared_strings)}">
{items}</sst>'''


# ─── Sheet Building Helpers ──────────────────────────────────────────────────

def col_letter(col_idx):
    """Convert 0-based column index to Excel column letter."""
    result = ""
    idx = col_idx
    while True:
        result = chr(ord('A') + idx % 26) + result
        idx = idx // 26 - 1
        if idx < 0:
            break
    return result


def build_cell(row, col, value=None, style=0, formula=None, is_string=False):
    """Build a single cell XML element."""
    ref = f"{col_letter(col)}{row}"
    attrs = f'r="{ref}"'
    if style > 0:
        attrs += f' s="{style}"'

    if formula is not None:
        return f'<c {attrs}><f>{formula}</f></c>'
    elif is_string and value is not None:
        idx = get_string_index(str(value))
        return f'<c {attrs} t="s"><v>{idx}</v></c>'
    elif value is not None:
        return f'<c {attrs}><v>{value}</v></c>'
    else:
        return f'<c {attrs}/>'


def build_row(row_num, cells_xml, ht=None, custom_height=False):
    """Wrap cells in a row element with optional height."""
    attrs = f'r="{row_num}"'
    if ht:
        attrs += f' ht="{ht}" customHeight="1"'
    return f'<row {attrs}>{cells_xml}</row>\n'



# ─── Sheet1: 总表 (Card-based summary with layered hierarchy) ────────────────

def build_sheet1():
    """
    Build the summary sheet with card-based design.
    Layout layers:
      Row 1: Title bar (large serif, underline - editorial masthead feel)
      Row 2: Subtitle / description  
      Row 3: Spacer
      Row 4: Card header (dark layer - column labels)
      Rows 5-23: Card body (alternating light/lighter - data rows)
    """
    rows_xml = ""

    # ── Layer 1: Editorial Title ──
    cells = build_cell(1, 0, value="面试评估总表", style=1, is_string=True)
    cells += build_cell(1, 1, style=1)
    cells += build_cell(1, 2, style=1)
    rows_xml += build_row(1, cells, ht=36)

    # ── Layer 2: Subtitle ──
    cells = build_cell(2, 0, value="Candidate Evaluation Summary  ·  Approachable Luxury", style=2, is_string=True)
    rows_xml += build_row(2, cells, ht=22)

    # ── Layer 3: Spacer ──
    cells = build_cell(3, 0, style=0)
    rows_xml += build_row(3, cells, ht=10)

    # ── Layer 4: Card Header (dark band) ──
    headers = [
        "排名", "候选人", "应聘岗位", "当前公司", "工作年限",
        "专业能力(40)", "沟通表达(25)", "组织协调(20)",
        "责任心(10)", "文化匹配(5)", "稳定性(5)", "总分", "面试结论"
    ]
    cells = ""
    for col_idx, h in enumerate(headers):
        cells += build_cell(4, col_idx, value=h, style=3, is_string=True)
    rows_xml += build_row(4, cells, ht=28)

    # ── Layer 5: Card Body (data rows with alternating fills) ──
    for i in range(19):
        row = 5 + i
        is_alt = (i % 2 == 1)
        cell_style = 5 if is_alt else 4  # alternating card body styles

        cells = ""
        # A: 排名
        formula_a = f'RANK(L{row},$L$5:$L$23,0)'
        cells += build_cell(row, 0, formula=formula_a, style=cell_style)

        # B-E: Manual input fields
        for col in range(1, 5):
            cells += build_cell(row, col, style=cell_style)

        # F: 专业能力
        formula_f = "&#x27;\u5355\u4eba\u8bc4\u4ef7\u9875&#x27;!B5+&#x27;\u5355\u4eba\u8bc4\u4ef7\u9875&#x27;!B7+&#x27;\u5355\u4eba\u8bc4\u4ef7\u9875&#x27;!B8"
        cells += build_cell(row, 5, formula=formula_f, style=cell_style)

        # G: 沟通表达
        formula_g = "&#x27;\u5355\u4eba\u8bc4\u4ef7\u9875&#x27;!B6+&#x27;\u5355\u4eba\u8bc4\u4ef7\u9875&#x27;!B9"
        cells += build_cell(row, 6, formula=formula_g, style=cell_style)

        # H: 组织协调
        formula_h = "&#x27;\u5355\u4eba\u8bc4\u4ef7\u9875&#x27;!B7+&#x27;\u5355\u4eba\u8bc4\u4ef7\u9875&#x27;!B13"
        cells += build_cell(row, 7, formula=formula_h, style=cell_style)

        # I: 责任心
        formula_i = "&#x27;\u5355\u4eba\u8bc4\u4ef7\u9875&#x27;!B8"
        cells += build_cell(row, 8, formula=formula_i, style=cell_style)

        # J: 文化匹配
        formula_j = "&#x27;\u5355\u4eba\u8bc4\u4ef7\u9875&#x27;!B10"
        cells += build_cell(row, 9, formula=formula_j, style=cell_style)

        # K: 稳定性
        formula_k = "&#x27;\u5355\u4eba\u8bc4\u4ef7\u9875&#x27;!B11+&#x27;\u5355\u4eba\u8bc4\u4ef7\u9875&#x27;!B12"
        cells += build_cell(row, 10, formula=formula_k, style=cell_style)

        # L: 总分
        formula_l = f'SUM(F{row}:K{row})'
        cells += build_cell(row, 11, formula=formula_l, style=8)  # accent total

        # M: 面试结论
        formula_m = f'IF(L{row}&gt;=90,"\u5f3a\u70c8\u63a8\u8350",IF(L{row}&gt;=80,"\u63a8\u8350\u5f55\u7528",IF(L{row}&gt;=70,"\u4fdd\u7559\u89c2\u5bdf","\u4e0d\u63a8\u8350")))'
        cells += build_cell(row, 12, formula=formula_m, style=cell_style)

        rows_xml += build_row(row, cells, ht=24)

    # Column widths - generous spacing for approachable luxury
    cols_xml = '<cols>\n'
    widths = [7, 14, 16, 16, 10, 14, 14, 14, 12, 12, 12, 10, 14]
    for i, w in enumerate(widths):
        cols_xml += f'  <col min="{i+1}" max="{i+1}" width="{w}" customWidth="1"/>\n'
    cols_xml += '</cols>\n'

    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetViews>
    <sheetView tabSelected="1" workbookViewId="0" showGridLines="0">
      <pane ySplit="4" topLeftCell="A5" activePane="bottomLeft" state="frozen"/>
    </sheetView>
  </sheetViews>
  {cols_xml}
  <sheetData>
{rows_xml}  </sheetData>
</worksheet>'''



# ─── Sheet2: 单人评价页 (Editorial Sketch Style Card Layout) ─────────────────

def build_sheet2():
    """
    Build individual evaluation sheet with editorial sketch aesthetic.
    Layout:
      - Masthead title with editorial underline
      - Info card (candidate details) 
      - Evaluation card with alternating rows
      - Score input cells stand out (white fill, bold border)
      - Sketch-style dashed separators between sections
    """
    rows_xml = ""

    # ── Editorial Masthead ──
    cells = build_cell(1, 0, value="面试候选人评价表", style=1, is_string=True)
    cells += build_cell(1, 1, style=1)
    cells += build_cell(1, 2, style=1)
    rows_xml += build_row(1, cells, ht=40)

    # Tagline
    cells = build_cell(2, 0, value="Individual Candidate Assessment  ·  Editorial Sketch", style=2, is_string=True)
    rows_xml += build_row(2, cells, ht=20)

    # ── Info Card Section ──
    # Section label with sketch underline
    cells = build_cell(3, 0, value="基本信息", style=9, is_string=True)
    rows_xml += build_row(3, cells, ht=24)

    # Info fields (label + value pairs)
    info_labels = [
        ("候选人姓名", 0), ("应聘岗位", 2), ("面试日期", 4),
    ]
    cells = ""
    cells += build_cell(4, 0, value="候选人姓名：", style=12, is_string=True)
    cells += build_cell(4, 1, style=13)  # value cell
    cells += build_cell(4, 2, value="应聘岗位：", style=12, is_string=True)
    cells += build_cell(4, 3, style=13)
    cells += build_cell(4, 4, value="面试日期：", style=12, is_string=True)
    cells += build_cell(4, 5, style=13)
    rows_xml += build_row(4, cells, ht=26)

    # Spacer
    rows_xml += build_row(5, build_cell(5, 0, style=0), ht=8)

    # ── Evaluation Card ──
    # Section label
    cells = build_cell(6, 0, value="评价维度", style=9, is_string=True)
    cells += build_cell(6, 1, value="", style=9, is_string=True)
    cells += build_cell(6, 2, value="", style=9, is_string=True)
    rows_xml += build_row(6, cells, ht=24)

    # Card header row
    card_headers = ["评价维度", "评分(0~10)", "评价说明"]
    cells = ""
    for col_idx, h in enumerate(card_headers):
        cells += build_cell(7, col_idx, value=h, style=3, is_string=True)
    rows_xml += build_row(7, cells, ht=26)

    # Evaluation dimensions with descriptions
    dimensions = [
        ("文书处理能力", "公文写作、材料整理、文档格式规范"),
        ("沟通协调能力", "语言表达、沟通技巧、协调各方关系"),
        ("组织安排能力", "工作计划制定、任务分配、流程管理"),
        ("时间管理能力", "时间规划、优先级判断、多任务处理"),
        ("责任心与执行力", "工作态度、任务完成度、主动承担"),
        ("适应力与学习力", "新环境适应速度、学习意愿与能力"),
        ("团队协作能力", "团队合作意识、配合度、共同目标"),
        ("企业文化匹配度", "价值观、工作风格与公司文化契合"),
        ("稳定性/职业规划匹配", "职业规划清晰度、岗位匹配、长期意愿"),
        ("稳定性评估（HR模块）", "HR综合评估候选人稳定性风险"),
        ("执行力评估（HR模块）", "HR综合评估候选人执行力水平"),
        ("办公技能评估（HR模块）", "Office/OA系统等办公软件熟练度"),
    ]

    for i, (dim, desc) in enumerate(dimensions):
        row_num = 8 + i  # rows 8-19
        is_alt = (i % 2 == 1)
        dim_style = 7 if is_alt else 6   # left-aligned card body
        score_style = 14                   # white input cell (stands out)
        desc_style = 7 if is_alt else 6   # left-aligned

        cells = ""
        cells += build_cell(row_num, 0, value=dim, style=dim_style, is_string=True)
        cells += build_cell(row_num, 1, value=0, style=score_style)
        cells += build_cell(row_num, 2, value=desc, style=desc_style, is_string=True)
        rows_xml += build_row(row_num, cells, ht=24)

    # ── Summary Section ──
    # Spacer
    rows_xml += build_row(20, build_cell(20, 0, style=0), ht=8)

    # Section label
    cells = build_cell(21, 0, value="评估汇总", style=9, is_string=True)
    rows_xml += build_row(21, cells, ht=24)

    # Total score row
    cells = ""
    cells += build_cell(22, 0, value="合计得分", style=8, is_string=True)
    cells += build_cell(22, 1, formula="SUM(B8:B19)", style=8)
    cells += build_cell(22, 2, value="满分120分，由总表按权重折算", style=10, is_string=True)
    rows_xml += build_row(22, cells, ht=28)

    # Spacer
    rows_xml += build_row(23, build_cell(23, 0, style=0), ht=8)

    # ── Notes Section (editorial feel) ──
    cells = build_cell(24, 0, value="综合评语", style=9, is_string=True)
    rows_xml += build_row(24, cells, ht=24)

    # Empty note area with subtle border
    cells = build_cell(25, 0, style=13)
    cells += build_cell(25, 1, style=13)
    cells += build_cell(25, 2, style=13)
    rows_xml += build_row(25, cells, ht=60)

    # Spacer
    rows_xml += build_row(26, build_cell(26, 0, style=0), ht=8)

    # ── Signature Section ──
    cells = ""
    cells += build_cell(27, 0, value="面试官签名：", style=12, is_string=True)
    cells += build_cell(27, 1, style=13)
    cells += build_cell(27, 2, value="日期：", style=12, is_string=True)
    rows_xml += build_row(27, cells, ht=26)

    # Footer meta
    cells = build_cell(28, 0, value="※ 评分标准：0~10分，总表公式自动引用计算", style=10, is_string=True)
    rows_xml += build_row(28, cells, ht=18)

    # Column widths - generous for editorial readability
    cols_xml = '''<cols>
  <col min="1" max="1" width="26" customWidth="1"/>
  <col min="2" max="2" width="14" customWidth="1"/>
  <col min="3" max="3" width="48" customWidth="1"/>
  <col min="4" max="6" width="14" customWidth="1"/>
</cols>'''

    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetViews>
    <sheetView workbookViewId="0" showGridLines="0"/>
  </sheetViews>
  {cols_xml}
  <sheetData>
{rows_xml}  </sheetData>
</worksheet>'''



# ─── Main: Assemble the XLSX ─────────────────────────────────────────────────

def main():
    # Build sheets (this also populates shared_strings)
    sheet1_xml = build_sheet1()
    sheet2_xml = build_sheet2()
    shared_strings_xml = build_shared_strings_xml()

    output_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), OUTPUT_FILE)

    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr('[Content_Types].xml', CONTENT_TYPES)
        zf.writestr('_rels/.rels', RELS)
        zf.writestr('xl/_rels/workbook.xml.rels', WORKBOOK_RELS)
        zf.writestr('xl/workbook.xml', WORKBOOK)
        zf.writestr('xl/styles.xml', STYLES)
        zf.writestr('xl/sharedStrings.xml', shared_strings_xml)
        zf.writestr('xl/worksheets/sheet1.xml', sheet1_xml)
        zf.writestr('xl/worksheets/sheet2.xml', sheet2_xml)

    print(f"✅ 面试评估表已生成: {output_path}")
    print()
    print("  🎨 设计风格: Approachable Luxury")
    print("  🖌️  配色方案: Morandi Watercolor Wash (Monochromatic)")
    print("  📐 布局策略: Card-Based Design with Layered Elements")
    print("  ✒️  视觉语言: Editorial Sketch Style")
    print()
    print(f"  📄 Sheet1: 总表 (Summary - frozen header, auto-rank)")
    print(f"  📄 Sheet2: 单人评价页 (Individual - editorial layout)")
    print(f"  🔤 Shared strings: {len(shared_strings)}")
    print()
    print("  色彩层次:")
    print("    ███ Darkest  #5B4E51 (headers/text)")
    print("    ███ Dark     #7D6B6E (sub-headers)")
    print("    ███ Medium   #A89396 (borders/accents)")
    print("    ░░░ Light    #D4C4C7 (card fill)")
    print("    ░░░ Lighter  #E8DFE1 (alt rows)")
    print("    ··· Lightest #F5F0F1 (background)")
    print("    ▓▓▓ Accent   #C9A9A0 (highlights)")


if __name__ == "__main__":
    main()
