from __future__ import annotations

from datetime import date, timedelta
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

OUT = Path('generated/建筑设计基础2')
OUT.mkdir(parents=True, exist_ok=True)
OUTPUT = OUT / '2023级建筑学_建筑设计基础2_教学进度表_2024春.docx'

COURSE = '建筑设计基础2'
CLASS_NAME = '2023级建筑学'
TEACHERS = '杨巧梅、周枳辛'
ROOM = '1B-305'
START = date(2024, 2, 26)

THEORY = [
    '建筑抄绘理论与案例剖析：课程导入；建筑抄绘的目的、方法、图纸内容及经典建筑案例整体认知',
    '建筑抄绘实践（一）理论指导：建筑平面图识读；比例、尺度、轴线、空间组合及功能关系分析',
    '建筑抄绘实践（一）理论指导：建筑立面、剖面与透视表达；形体、构造和空间关系的对应',
    '建筑抄绘实践（二）深化指导：经典建筑空间序列、流线组织、比例控制及线条层级',
    '建筑抄绘实践（二）总结：图纸完善、建筑分析、设计元素提炼及抄绘成果汇报方法',
    '平面构成原理与练习（一）：平面构成概念；点、线、面基本元素及视觉特征',
    '平面构成原理与练习（二）：对称、均衡、重复、渐变、发射及特异等形式法则',
    '平面构成原理与练习（三）：对比与调和、节奏韵律、比例、正负形及共生图形',
    '立体构成原理与练习（一）：立体构成概念、空间限定、块体组合及体量关系',
    '立体构成原理与练习（二）：线材、面材构成；材料特性、连接方式与结构稳定',
    '立体构成原理与练习（三）：形态组合、空间序列、光影关系及作品表达评析',
    '模型制作材料与工具讲解：模型比例、材料性能、刀具胶黏剂使用、安全操作及制作流程',
    '模型制作实践（一）指导：设计图纸识读、模型分解、底板定位、主体结构与墙体制作',
    '模型制作实践（一）深化：楼板、屋面、门窗、立面构件及主要空间关系的模型表达',
    '模型制作实践（二）指导：细部完善、材质色彩、场地环境、景观构件及整体效果控制',
    '模型制作实践（二）与课程总结：成果展示、设计说明、模型摄影、汇报评价与学习复盘',
]

PRACTICE = [
    '建筑抄绘准备：熟悉工具与线型，选择案例，确定比例和版面，完成图纸资料整理',
    '建筑抄绘实践（一）：完成经典建筑平面图抄绘，校核比例、尺度、轴线和空间关系',
    '建筑抄绘实践（一）：完成立面图、剖面图及外观透视图抄绘，统一图线和标注',
    '建筑抄绘实践（二）：深化平立剖图纸，修正比例、构造细节和图面层级',
    '建筑抄绘实践（二）：完成建筑分析图、抄绘心得和成果版面，组织阶段展示与讲评',
    '平面构成练习（一）：点、线、面组合及基本骨格练习，形成多组构成草案',
    '平面构成练习（二）：运用对称、均衡、重复、渐变等法则完成主题构成练习',
    '平面构成练习（三）：完成正负形、共生图形或综合色彩构成，优化版面并讲评',
    '立体构成练习（一）：完成块体切割、组合与空间限定练习，比较不同体量关系',
    '立体构成练习（二）：完成线材、面材构成，重点检查节点连接与结构稳定性',
    '立体构成练习（三）：完成综合立体构成作品，调整空间、光影、比例和展示方式',
    '模型制作准备：工具安全训练、材料切割与粘接试验，确定模型比例和加工计划',
    '模型制作实践（一）：制作底板、定位轴线、主体结构、墙体及主要空间构件',
    '模型制作实践（一）：完成楼板、屋面、门窗和立面构件，校核比例与空间关系',
    '模型制作实践（二）：完善材质、色彩、场地与景观细部，修正模型工艺问题',
    '模型制作实践（二）与总结：完成最终模型、成果摄影、展示汇报、评价反馈与归档',
]

assert len(THEORY) == 16 and len(PRACTICE) == 16


def font(run, size=8.0, bold=False, name='宋体'):
    run.font.name = name
    run._element.rPr.rFonts.set(qn('w:eastAsia'), name)
    run.font.size = Pt(size)
    run.bold = bold


def cell_text(cell, text, size=7.5, bold=False, align=WD_ALIGN_PARAGRAPH.CENTER):
    cell.text = ''
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.0
    r = p.add_run(str(text))
    font(r, size=size, bold=bold)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def shade(cell, fill='EAF2F8'):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = tcPr.find(qn('w:shd'))
    if shd is None:
        shd = OxmlElement('w:shd')
        tcPr.append(shd)
    shd.set(qn('w:fill'), fill)


def margins(cell, value=35):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = tcPr.first_child_found_in('w:tcMar')
    if tcMar is None:
        tcMar = OxmlElement('w:tcMar')
        tcPr.append(tcMar)
    for side in ('top', 'start', 'bottom', 'end'):
        node = tcMar.find(qn(f'w:{side}'))
        if node is None:
            node = OxmlElement(f'w:{side}')
            tcMar.append(node)
        node.set(qn('w:w'), str(value))
        node.set(qn('w:type'), 'dxa')


def repeat_header(row):
    trPr = row._tr.get_or_add_trPr()
    node = OxmlElement('w:tblHeader')
    node.set(qn('w:val'), 'true')
    trPr.append(node)


def fixed_layout(table):
    tblPr = table._tbl.tblPr
    layout = tblPr.find(qn('w:tblLayout'))
    if layout is None:
        layout = OxmlElement('w:tblLayout')
        tblPr.append(layout)
    layout.set(qn('w:type'), 'fixed')


def set_cell_width(cell, cm):
    tcPr = cell._tc.get_or_add_tcPr()
    tcW = tcPr.find(qn('w:tcW'))
    if tcW is None:
        tcW = OxmlElement('w:tcW')
        tcPr.append(tcW)
    tcW.set(qn('w:w'), str(int(cm * 567)))
    tcW.set(qn('w:type'), 'dxa')


def add_heading(doc, continuation=False):
    top = doc.add_table(rows=1, cols=2)
    top.alignment = WD_TABLE_ALIGNMENT.CENTER
    top.autofit = False
    set_cell_width(top.cell(0,0), 23.5)
    set_cell_width(top.cell(0,1), 4.5)
    top.cell(0,0).text = ''
    p = top.cell(0,0).paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run('附件3：城乡规划教研室\n教学进度表' + ('（续）' if continuation else ''))
    font(r, size=15 if not continuation else 13, bold=True, name='黑体')

    summary = top.cell(0,1)
    summary.text = ''
    st = summary.add_table(rows=3, cols=2)
    st.style = 'Table Grid'
    for i, (a,b) in enumerate([('计划总学时','96学时'),('理论课学时','48学时'),('实验课学时','48学时')]):
        cell_text(st.cell(i,0), a, 8.2, True)
        cell_text(st.cell(i,1), b, 8.2)
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run('课程名称  建筑设计基础2    班级  2023级建筑学    从2024年2月26日至2024年6月14日止')
    font(r, size=9.2)


def add_progress_table(doc, weeks):
    widths = [1.35, 1.0, 0.62, 0.62, 0.48, 0.48, 0.52, 0.62, 6.15, 0.75,
              1.35, 1.0, 0.62, 0.48, 0.48, 0.62, 5.55, 0.8]
    table = doc.add_table(rows=2, cols=18)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    fixed_layout(table)
    for row in table.rows:
        for i, c in enumerate(row.cells):
            set_cell_width(c, widths[i])
            margins(c)

    cell_text(table.cell(0,0), '理论课教学安排', 9, True)
    table.cell(0,0).merge(table.cell(0,9))
    shade(table.cell(0,0), 'D9EAF7')
    cell_text(table.cell(0,10), '实验课教学安排', 9, True)
    table.cell(0,10).merge(table.cell(0,17))
    shade(table.cell(0,10), 'D9EAF7')

    headers = ['教师姓名','职称','周次','年','月','日','时数','讲课章节及内容','备注','',
               '教师姓名','职称','周次','月','日','时数','实验、实践内容','学生分组']
    # Merge empty auxiliary column into theory content/remarks area to match wide template proportions.
    for i, h in enumerate(headers):
        cell_text(table.cell(1,i), h, 7.4, True)
        shade(table.cell(1,i), 'EAF2F8')
    # Use columns 7-8 for theory content and remarks; col9 is kept as narrow spacer/remarks continuation.
    cell_text(table.cell(1,8), '讲课章节及内容', 7.4, True)
    cell_text(table.cell(1,9), '备注', 7.4, True)
    repeat_header(table.rows[0]); repeat_header(table.rows[1])

    for w in weeks:
        monday = START + timedelta(weeks=w-1)
        wednesday = monday + timedelta(days=2)
        row = table.add_row()
        cells = row.cells
        for i, c in enumerate(cells):
            set_cell_width(c, widths[i]); margins(c)
        values = [
            TEACHERS, '', w, monday.year, monday.month, monday.day, 3, '', THEORY[w-1], '',
            TEACHERS, '', w, wednesday.month, wednesday.day, 3, PRACTICE[w-1], ''
        ]
        for i, value in enumerate(values):
            align = WD_ALIGN_PARAGRAPH.LEFT if i in (8,16) else WD_ALIGN_PARAGRAPH.CENTER
            size = 7.0 if i in (8,16) else 7.2
            cell_text(cells[i], value, size=size, align=align)
        row.height = Cm(1.38)
    return table


def build():
    doc = Document()
    sec = doc.sections[0]
    sec.orientation = WD_ORIENT.LANDSCAPE
    sec.page_width, sec.page_height = sec.page_height, sec.page_width
    sec.top_margin = Cm(0.65)
    sec.bottom_margin = Cm(0.65)
    sec.left_margin = Cm(0.55)
    sec.right_margin = Cm(0.55)
    doc.styles['Normal'].font.name = '宋体'
    doc.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    doc.styles['Normal'].font.size = Pt(8)

    add_heading(doc, False)
    add_progress_table(doc, range(1,9))

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(0)
    r = p.add_run('上课时间/地点：周一下午第5、6、7节（理论课），周三第7、8、9节（实验、实践课）/1B-305')
    font(r, size=8.6)

    doc.add_page_break()
    add_heading(doc, True)
    add_progress_table(doc, range(9,17))
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(0)
    r = p.add_run('上课时间/地点：周一下午第5、6、7节（理论课），周三第7、8、9节（实验、实践课）/1B-305')
    font(r, size=8.6)

    doc.core_properties.title = '建筑设计基础2教学进度表'
    doc.core_properties.subject = '2023级建筑学，2024年春季学期，96学时'
    doc.core_properties.author = '杨巧梅、周枳辛'
    doc.save(OUTPUT)
    print(OUTPUT, OUTPUT.stat().st_size)


if __name__ == '__main__':
    build()
