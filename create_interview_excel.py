#!/usr/bin/env python3
"""
Generate an Excel (.xlsx) file for interview evaluation system.
Uses only Python standard library (zipfile + xml) to build Open XML spreadsheet.
"""

import zipfile
import os

OUTPUT_FILE = "面试评估表.xlsx"

# ─── XML Templates ───────────────────────────────────────────────────────────

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
    <sheet name="总表" sheetId="1" r:id="rId1"/>
    <sheet name="单人评价页" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>'''

STYLES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font><sz val="11"/><name val="Microsoft YaHei"/></font>
    <font><b/><sz val="11"/><name val="Microsoft YaHei"/></font>
  </fonts>
  <fills count="3">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF4472C4"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border>
      <left style="thin"><color auto="1"/></left>
      <right style="thin"><color auto="1"/></right>
      <top style="thin"><color auto="1"/></top>
      <bottom style="thin"><color auto="1"/></bottom>
      <diagonal/>
    </border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="4">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center" wrapText="1"/>
    </xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center"/>
    </xf>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center" wrapText="1"/>
    </xf>
  </cellXfs>
</styleSheet>'''


# ─── Shared Strings ──────────────────────────────────────────────────────────

# All text strings used in both sheets
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
        # Escape XML special characters
        escaped = s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        items += f'  <si><t>{escaped}</t></si>\n'
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="{len(shared_strings)}" uniqueCount="{len(shared_strings)}">
{items}</sst>'''


# ─── Sheet Building Helpers ──────────────────────────────────────────────────

def col_letter(col_idx):
    """Convert 0-based column index to Excel column letter (A, B, ..., Z, AA, ...)."""
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


def build_row(row_num, cells_xml):
    """Wrap cells in a row element."""
    return f'<row r="{row_num}">{cells_xml}</row>\n'


# ─── Sheet1: 总表 ────────────────────────────────────────────────────────────

def build_sheet1():
    """Build the summary sheet (总表)."""
    headers = [
        "排名", "候选人", "应聘岗位", "当前公司", "工作年限",
        "专业能力(40)", "沟通表达(25)", "组织协调(20)",
        "责任心(10)", "文化匹配(5)", "稳定性(5)", "总分", "面试结论"
    ]

    # Column widths
    cols_xml = '<cols>\n'
    widths = [6, 12, 14, 14, 10, 14, 14, 14, 12, 12, 12, 8, 14]
    for i, w in enumerate(widths):
        cols_xml += f'  <col min="{i+1}" max="{i+1}" width="{w}" customWidth="1"/>\n'
    cols_xml += '</cols>\n'

    rows_xml = ""

    # Header row (row 1)
    cells = ""
    for col_idx, h in enumerate(headers):
        cells += build_cell(1, col_idx, value=h, style=1, is_string=True)
    rows_xml += build_row(1, cells)

    # Data rows (rows 2-20) with formulas
    for row in range(2, 21):
        cells = ""
        # A: 排名 - RANK formula
        formula_a = f'RANK(L{row},$L$2:$L$20,0)'
        cells += build_cell(row, 0, formula=formula_a, style=2)

        # B: 候选人 (empty, to be filled)
        cells += build_cell(row, 1, style=2)
        # C: 应聘岗位 (empty)
        cells += build_cell(row, 2, style=2)
        # D: 当前公司 (empty)
        cells += build_cell(row, 3, style=2)
        # E: 工作年限 (empty)
        cells += build_cell(row, 4, style=2)

        # F: 专业能力 = 单人评价页!B4 + B6 + B7
        formula_f = "&#x27;单人评价页&#x27;!B4+&#x27;单人评价页&#x27;!B6+&#x27;单人评价页&#x27;!B7"
        cells += build_cell(row, 5, formula=formula_f, style=2)

        # G: 沟通表达 = 单人评价页!B5 + B8
        formula_g = "&#x27;单人评价页&#x27;!B5+&#x27;单人评价页&#x27;!B8"
        cells += build_cell(row, 6, formula=formula_g, style=2)

        # H: 组织协调 = 单人评价页!B6 + B12
        formula_h = "&#x27;单人评价页&#x27;!B6+&#x27;单人评价页&#x27;!B12"
        cells += build_cell(row, 7, formula=formula_h, style=2)

        # I: 责任心 = 单人评价页!B7
        formula_i = "&#x27;单人评价页&#x27;!B7"
        cells += build_cell(row, 8, formula=formula_i, style=2)

        # J: 文化匹配 = 单人评价页!B8
        formula_j = "&#x27;单人评价页&#x27;!B8"
        cells += build_cell(row, 9, formula=formula_j, style=2)

        # K: 稳定性 = 单人评价页!B9 + B10
        formula_k = "&#x27;单人评价页&#x27;!B9+&#x27;单人评价页&#x27;!B10"
        cells += build_cell(row, 10, formula=formula_k, style=2)

        # L: 总分 = SUM(F:K)
        formula_l = f'SUM(F{row}:K{row})'
        cells += build_cell(row, 11, formula=formula_l, style=2)

        # M: 面试结论
        formula_m = f'IF(L{row}&gt;=90,"强烈推荐",IF(L{row}&gt;=80,"推荐录用",IF(L{row}&gt;=70,"保留观察","不推荐")))'
        cells += build_cell(row, 12, formula=formula_m, style=2)

        rows_xml += build_row(row, cells)

    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetViews>
    <sheetView tabSelected="1" workbookViewId="0">
      <pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>
    </sheetView>
  </sheetViews>
  {cols_xml}
  <sheetData>
{rows_xml}  </sheetData>
</worksheet>'''


# ─── Sheet2: 单人评价页 ──────────────────────────────────────────────────────

def build_sheet2():
    """Build the individual evaluation sheet (单人评价页)."""

    # Title row
    rows_xml = ""

    # Row 1: Title
    cells = build_cell(1, 0, value="面试候选人评价表", style=3, is_string=True)
    rows_xml += build_row(1, cells)

    # Row 2: Candidate info area
    cells = build_cell(2, 0, value="候选人姓名：", style=3, is_string=True)
    cells += build_cell(2, 2, value="应聘岗位：", style=3, is_string=True)
    cells += build_cell(2, 4, value="面试日期：", style=3, is_string=True)
    rows_xml += build_row(2, cells)

    # Row 3: Headers
    headers = ["评价维度", "评分(0~10)", "评价说明"]
    cells = ""
    for col_idx, h in enumerate(headers):
        cells += build_cell(3, col_idx, value=h, style=1, is_string=True)
    rows_xml += build_row(3, cells)

    # Row 4-15: Evaluation dimensions
    dimensions = [
        ("文书处理能力", "考察公文写作、材料整理、文档格式规范等能力"),
        ("沟通协调能力", "考察语言表达、沟通技巧、协调各方关系的能力"),
        ("组织安排能力", "考察工作计划制定、任务分配、流程管理能力"),
        ("时间管理能力", "考察时间规划、优先级判断、多任务处理能力"),
        ("责任心与执行力", "考察工作态度、任务完成度、主动承担意识"),
        ("适应力与学习力", "考察对新环境/新任务的适应速度和学习意愿"),
        ("团队协作能力", "考察团队合作意识、配合度、共同目标感"),
        ("企业文化匹配度", "考察价值观、工作风格与公司文化的契合程度"),
        ("稳定性/职业规划匹配", "考察职业规划清晰度、岗位匹配度、长期稳定意愿"),
        ("稳定性评估（HR模块）", "HR综合评估候选人的稳定性风险"),
        ("执行力评估（HR模块）", "HR综合评估候选人的执行力水平"),
        ("办公技能评估（HR模块）", "HR评估Office/OA系统等办公软件熟练度"),
    ]

    for i, (dim, desc) in enumerate(dimensions):
        row_num = 4 + i
        cells = ""
        cells += build_cell(row_num, 0, value=dim, style=2, is_string=True)
        cells += build_cell(row_num, 1, value=0, style=2)  # Default score 0
        cells += build_cell(row_num, 2, value=desc, style=2, is_string=True)
        rows_xml += build_row(row_num, cells)

    # Row 16: blank
    # Row 17: Summary
    row_sum = 4 + len(dimensions) + 1
    cells = build_cell(row_sum, 0, value="合计得分", style=3, is_string=True)
    cells += build_cell(row_sum, 1, formula="SUM(B4:B15)", style=3)
    rows_xml += build_row(row_sum, cells)

    # Row 18: Notes section
    row_notes = row_sum + 1
    cells = build_cell(row_notes, 0, value="综合评语：", style=3, is_string=True)
    rows_xml += build_row(row_notes, cells)

    # Row 19: Interviewer
    row_int = row_notes + 2
    cells = build_cell(row_int, 0, value="面试官签名：", style=3, is_string=True)
    cells += build_cell(row_int, 2, value="日期：", style=3, is_string=True)
    rows_xml += build_row(row_int, cells)

    # Column widths
    cols_xml = '''<cols>
  <col min="1" max="1" width="24" customWidth="1"/>
  <col min="2" max="2" width="12" customWidth="1"/>
  <col min="3" max="3" width="45" customWidth="1"/>
</cols>'''

    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetViews>
    <sheetView workbookViewId="0"/>
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

    print(f"✅ Excel file created: {output_path}")
    print(f"   - Sheet1: 总表 (Summary with formulas)")
    print(f"   - Sheet2: 单人评价页 (Individual evaluation)")
    print(f"   - {len(shared_strings)} shared strings")


if __name__ == "__main__":
    main()
