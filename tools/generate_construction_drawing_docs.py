from __future__ import annotations

from datetime import date, timedelta
from pathlib import Path
from typing import Iterable

from docx import Document
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

OUT = Path("generated/施工图设计")
OUT.mkdir(parents=True, exist_ok=True)

COURSE = "施工图设计"
TEACHER = "尚谢俊"
TITLE = "工程师"
COLLEGE = "建筑工程学院"
DEPARTMENT = "建筑学教研室"
MAJOR = "建筑学"
TEXTBOOK = "《建筑施工图设计》"
TEXTBOOK_AUTHOR = "单立欣、穆丽丽"
TEXTBOOK_PUBLISHER = "机械工业出版社"
TEXTBOOK_DATE = "2019年1月"

WEEK_TOPICS = [
    ("理论", "第一章 建筑工程设计阶段与施工图概述（一）：课程导论；建筑工程设计阶段、基本工作流程及施工图阶段的职责与成果深度"),
    ("理论", "第一章 建筑工程设计阶段与施工图概述（二）：施工图的概念、作用和特点；全套图纸组成、图纸目录、图号体系及编排顺序"),
    ("理论", "第二章 房屋建筑制图规定（一）：图纸幅面、标题栏与会签栏；图线、字体、比例、常用符号和图例"),
    ("理论", "第二章 房屋建筑制图规定（二）：定位轴线、标高、尺寸标注、索引与详图符号、剖切符号及常见制图错误辨析"),
    ("理论", "第三章 建筑专业施工图设计（一）：封面、图纸目录、建筑设计说明、主要技术指标、材料做法表、门窗表及专项说明"),
    ("理论", "第三章 建筑专业施工图设计（二）：建筑总平面图和建筑平面图的内容、绘制步骤、尺寸系统、标高、索引及构造定位"),
    ("理论", "第三章 建筑专业施工图设计（三）：平面图深化；建筑立面图、剖面图的形体、材料、标高及跨图纸对应关系"),
    ("理论", "第四章 常用建筑详图：楼梯、台阶、坡道、卫生间、门窗、墙身、屋面等重点部位的构造与大样表达"),
    ("理论", "第五章 多专业配合与施工图校审：建筑与结构、给排水、暖通、电气专业接口；校对审核流程、校审提纲及常见问题"),
    ("实践", "第六章 施工图绘制实践（一）：整理方案资料，确定图纸目录、图号、图幅和比例，编制建筑设计说明及主要表格"),
    ("实践", "第六章 施工图绘制实践（二）：建筑总平面图——建筑定位、道路出入口、消防场地、停车组织、竖向及经济技术指标"),
    ("实践", "第六章 施工图绘制实践（三）：建筑平面图第一阶段——轴网、墙柱、房间功能、门窗、楼梯及首层/标准层基本表达"),
    ("实践", "第六章 施工图绘制实践（四）：建筑平面图第二阶段——尺寸链、标高、剖切与索引、各层对应、设备空间及专业条件预留"),
    ("实践", "第六章 施工图绘制实践（五）：建筑立面图——外轮廓、门窗、檐口、雨篷、栏杆、材料分格、轴线和主要标高"),
    ("实践", "第六章 施工图绘制实践（六）：建筑剖面图——剖切位置、层高净高、楼地面与屋面、楼梯及竖向尺寸标高体系"),
    ("实践", "第六章 施工图绘制实践（七）：建筑大样图——楼梯、墙身、门窗或卫生间等节点的构造层次、材料、尺寸与索引"),
    ("实践", "第六章 施工图绘制实践（八）：全套施工图自检、互检和集中修改；设计说明与总平面、平立剖及大样的一致性校核"),
    ("实践", "第六章 施工图绘制实践（九）：教师终审、成果统一修改、图纸整理与归档；成果展示、问题复盘和课程总结"),
]

LESSONS = [
    {
        "title": WEEK_TOPICS[0][1],
        "goals": ["了解课程任务、学习方法及施工图成果在工程建设中的作用。", "掌握建筑工程方案设计、初步设计和施工图设计等主要阶段及其工作边界。", "能够识别施工图阶段设计人员的基本职责、成果深度与质量责任。", "树立依法设计、规范作图和对工程实施负责的职业意识。"],
        "key": "建筑工程设计阶段、施工图阶段职责及成果深度。",
        "hard": "理解工程流程、设计责任与施工图成果之间的逻辑关系。",
        "process": [
            ("课程导入与任务说明", 10, "展示方案图、施工图和建成照片，说明方案向工程实施转化的必要性；介绍36学时、18周的理论—实践安排及成果要求。"),
            ("建筑工程设计全过程", 25, "讲解项目策划、方案设计、初步设计、施工图设计、施工配合和竣工验收等环节，明确建筑专业在各阶段的主要任务。"),
            ("施工图阶段的职责与深度", 25, "结合工程案例分析设计依据、技术条件、规范责任、图纸深度以及设计人员、校对人、审核人的岗位分工。"),
            ("案例辨析与课堂讨论", 20, "对比同一项目的方案图和施工图，引导学生找出新增的尺寸、材料、构造、标高和专业条件信息。"),
            ("总结与作业", 10, "梳理工程流程和施工图阶段责任；布置施工图案例收集与图纸组成预习任务。"),
        ],
        "discussion": "为什么一个视觉效果较完整的建筑方案仍不能直接用于施工？",
        "homework": "收集一套小型公共建筑或住宅施工图，列出其中的图纸类别及每类图纸承担的主要信息。",
        "reflection": "关注学生能否从“表达方案”转向“服务实施”理解施工图；下次课根据案例收集情况强化图纸体系认知。",
    },
    {
        "title": WEEK_TOPICS[1][1],
        "goals": ["掌握建筑施工图的概念、作用、特点和基本设计深度。", "熟悉全套建筑施工图的组成、图纸目录、图号体系和编排顺序。", "能够依据项目规模初步建立施工图成果框架。", "形成完整表达和系统组织图纸的意识。"],
        "key": "全套建筑施工图的组成、图号体系与编排逻辑。",
        "hard": "根据项目特点确定图纸取舍，并建立各图种之间的信息关联。",
        "process": [
            ("复习导入", 5, "回顾建筑工程设计阶段，检查学生收集的施工图案例。"),
            ("施工图概念与特点", 20, "说明施工图的法律技术属性、实施依据作用以及完整性、准确性、一致性和可操作性要求。"),
            ("全套图纸组成与编排", 30, "讲解封面、目录、设计说明、总平面、平面、立面、剖面、大样和表格的组成及常用编排顺序。"),
            ("完整图纸导读", 25, "以一套实际项目为例追踪图号、索引和详图之间的关联，分析图纸缺项对施工的影响。"),
            ("课堂练习与总结", 10, "学生为给定小型建筑拟定图纸目录，教师点评并布置修订。"),
        ],
        "discussion": "图纸数量越多是否代表施工图质量越高？",
        "homework": "为一个两层小型公共建筑编制初步图纸目录，注明图号、图名和建议比例。",
        "reflection": "检查学生能否按“先总后分、先整体后局部”的逻辑组织图纸，并识别目录与实际图纸不一致的问题。",
    },
    {
        "title": WEEK_TOPICS[2][1],
        "goals": ["掌握图幅、标题栏、会签栏、图线、字体和比例的基本规定。", "熟悉建筑施工图常用符号和图例。", "能够正确设置基本图面并识别图线、比例和文字表达错误。", "养成统一、清晰、规范的制图习惯。"],
        "key": "图线系统、比例选择、字体与图面组织。",
        "hard": "在信息密集的完整图纸中建立清晰的表达层级。",
        "process": [
            ("错误图样导入", 10, "展示图线混乱、比例不当和文字不统一的图纸，请学生判断阅读困难的原因。"),
            ("图幅与图框系统", 20, "讲解常用图幅、加长规则、标题栏、会签栏、装订边和图纸方向。"),
            ("图线、字体与比例", 30, "示范粗中细图线层级、字体高度、比例标注及不同图种的比例选择。"),
            ("符号图例与案例对照", 20, "讲解指北针、风玫瑰、材料图例、引出线等常用表达，比较规范图与问题图。"),
            ("课堂练习与作业", 10, "学生完成一张A3图框及图线样式练习，教师现场纠错。"),
        ],
        "discussion": "电子出图条件下，为什么仍需严格区分图线宽度和层级？",
        "homework": "按A3图幅建立施工图标准图框，设置标题栏、文字和常用图线图层。",
        "reflection": "关注学生是否仅记忆参数而忽略图面层级；下一次课将制图规定落实到轴线、尺寸和索引系统。",
    },
    {
        "title": WEEK_TOPICS[3][1],
        "goals": ["掌握定位轴线、标高、尺寸、索引、详图和剖切符号的表达方法。", "理解不同图种之间编号和标注的一致性要求。", "能够诊断常见制图错误并提出修改方案。", "形成主动校核和精确表达意识。"],
        "key": "轴线、尺寸、标高、索引和剖切符号的规范表达。",
        "hard": "保持平面、立面、剖面和大样图中编号、标高与索引关系一致。",
        "process": [
            ("复习与问题导入", 5, "检查图框作业，展示轴号重复、尺寸链缺项和索引失效等问题。"),
            ("定位轴线与编号", 20, "讲解轴线类型、编号顺序、附加轴线、圆形和复杂平面轴网表达。"),
            ("尺寸与标高系统", 25, "讲解外部尺寸链、内部尺寸、定位尺寸、总尺寸及绝对/相对标高的使用。"),
            ("索引、详图与剖切符号", 25, "示范索引符号、详图编号、剖切位置线、剖视方向及跨图纸索引。"),
            ("错误诊断练习", 15, "分组核查问题图样，列出错误、依据和修改方式，教师集中讲评。"),
        ],
        "discussion": "哪些尺寸必须在建筑图中明确，哪些信息应由结构或设备专业表达？",
        "homework": "修改给定问题图样，统一轴线、尺寸、标高和索引系统，并形成错误清单。",
        "reflection": "通过错误清单判断学生是否建立“表达—校核—修正”的闭环，而非仅完成单项标注。",
    },
    {
        "title": WEEK_TOPICS[4][1],
        "goals": ["掌握建筑设计说明的组成、编写逻辑和常用表格内容。", "了解主要技术指标、材料做法、门窗表及节能、消防等专项说明的关系。", "能够将方案信息转化为规范、准确、可核查的文字和表格。", "形成设计依据明确、文字表达严谨的职业习惯。"],
        "key": "建筑设计说明、技术指标和材料门窗表的完整性与一致性。",
        "hard": "避免照抄通用文本，使说明内容与具体项目及图纸相对应。",
        "process": [
            ("说明文件导入", 10, "比较完整设计说明与模板化说明，指出项目依据、指标和做法缺项。"),
            ("设计说明组成", 25, "讲解工程概况、设计依据、设计范围、分类等级、主要技术指标、构造做法和安全要求。"),
            ("表格与专项内容", 25, "讲解材料做法表、房间用料表、门窗表及节能、消防、无障碍等内容的衔接。"),
            ("案例核查", 20, "依据一套施工图核对说明、门窗编号和材料做法是否与图纸一致。"),
            ("总结与作业", 10, "形成设计说明编写清单，布置项目说明框架。"),
        ],
        "discussion": "通用设计说明可以直接复制到所有项目中吗？哪些内容必须项目化？",
        "homework": "选择一个既有方案，编写建筑设计说明框架和主要技术指标表。",
        "reflection": "重点检查设计依据、项目指标和材料做法能否与方案相互印证，及时纠正空泛表述。",
    },
    {
        "title": WEEK_TOPICS[5][1],
        "goals": ["掌握建筑总平面图和平面图的主要内容及绘制步骤。", "理解定位、交通、竖向、尺寸、门窗、标高、索引和构造定位的表达要求。", "能够核查总平面与各层平面的对应关系。", "形成从总体到单体、从功能到实施信息的系统思维。"],
        "key": "总平面定位与平面图尺寸、功能、构造信息的完整表达。",
        "hard": "将方案平面深化为轴网清晰、尺寸闭合、构造可定位的施工图。",
        "process": [
            ("图种关系导入", 5, "从建筑定位追踪到首层轴网，说明总平面与单体平面的关联。"),
            ("建筑总平面图", 25, "讲解建筑定位、坐标、道路、出入口、消防场地、停车、室外标高、指标及图例。"),
            ("建筑平面图内容", 30, "讲解轴网、墙柱、房间、门窗、固定设施、楼梯、尺寸链、标高、剖切与索引。"),
            ("案例拆解与联动核查", 20, "核查总平面建筑轮廓、首层出入口、室内外高差和道路关系。"),
            ("练习与作业", 10, "学生标注给定平面的轴网和三道外部尺寸，教师点评。"),
        ],
        "discussion": "方案平面中哪些“看起来合理”的信息在施工图阶段仍需重新确定？",
        "homework": "完成给定建筑的总平面信息清单和平面图深化问题清单。",
        "reflection": "检查学生是否理解总平面与首层平面之间的定位、出入口和标高对应关系。",
    },
    {
        "title": WEEK_TOPICS[6][1],
        "goals": ["进一步掌握建筑平面图深化要求。", "掌握立面图和剖面图的内容、绘制方法及标高体系。", "能够校核平、立、剖图之间的形体、门窗、尺寸和标高关系。", "培养跨图纸协调和整体表达能力。"],
        "key": "立面、剖面表达及平立剖之间的一致性。",
        "hard": "由二维平面信息准确推导建筑外部形态和竖向空间构造。",
        "process": [
            ("平面深化复盘", 10, "总结轴网、尺寸、门窗和楼梯等对立剖面绘制的前置作用。"),
            ("建筑立面图", 25, "讲解立面方向、轮廓、门窗、檐口、雨篷、栏杆、材料分格、轴线和标高。"),
            ("建筑剖面图", 30, "讲解剖切位置、层高净高、楼地面屋面、楼梯、基础关系及竖向尺寸标高。"),
            ("跨图纸校核", 15, "通过叠合与追踪检查门窗位置、层高、屋面和楼梯是否一致。"),
            ("总结与作业", 10, "学生为给定方案选择剖切位置并说明理由。"),
        ],
        "discussion": "怎样选择最能反映建筑空间和构造关系的剖切位置？",
        "homework": "完成一个主要立面和一个代表性剖面的信息草图，标注需与平面核对的项目。",
        "reflection": "关注学生是否从平面机械投影转向空间与构造表达，并能主动发现平立剖矛盾。",
    },
    {
        "title": WEEK_TOPICS[7][1],
        "goals": ["掌握楼梯、墙身、门窗、卫生间和屋面等常用部位的构造关系。", "理解大样图比例、尺寸、材料、连接和索引要求。", "能够从设计意图推导基本合理的节点并规范表达。", "培养设计、构造与制图一体化意识。"],
        "key": "常用构件的构造层次及大样图完整表达。",
        "hard": "将抽象设计意图转化为材料、尺寸和连接关系明确的可实施节点。",
        "process": [
            ("节点问题导入", 10, "展示仅有轮廓而缺少构造层次的大样图，分析其无法施工的原因。"),
            ("楼梯与重点空间", 25, "讲解楼梯平剖大样、台阶坡道、卫生间降板防水和设备用房构造要点。"),
            ("围护构造节点", 25, "讲解墙身、门窗、屋面、顶棚等节点的材料层次、收口、防水和保温表达。"),
            ("大样表达与索引", 20, "示范比例选择、详图编号、材料引注、尺寸标高及与平立剖图的索引对应。"),
            ("课堂练习", 10, "学生补全一个墙身或卫生间节点的材料层次与标注。"),
        ],
        "discussion": "建筑大样应表达多少信息，如何避免既缺项又过度表达？",
        "homework": "选择楼梯、墙身或卫生间之一，绘制节点草图并列出材料和关键尺寸。",
        "reflection": "判断学生能否解释构造层次的功能与先后关系，而不是机械临摹标准图。",
    },
    {
        "title": WEEK_TOPICS[8][1],
        "goals": ["了解建筑与结构、给排水、暖通、电气等专业的主要接口。", "掌握施工图设计、校对、审核的基本流程和校审提纲。", "能够识别跨专业、跨图种的常见冲突并提出修改方向。", "形成协同设计、主动校审和成果负责意识。"],
        "key": "专业接口、校审程序和全套图纸检查要点。",
        "hard": "发现管井、洞口、设备空间、标高等跨专业隐性矛盾。",
        "process": [
            ("冲突案例导入", 10, "展示梁与门洞、风管与净高、管井与功能空间冲突等案例。"),
            ("多专业配合", 25, "讲解结构柱网、管井、设备用房、洞口预留、标高和竖向系统协调。"),
            ("校对审核流程", 20, "介绍设计、校对、审核职责及图纸审查流程。"),
            ("校审提纲与问题诊断", 25, "按图号、轴线、尺寸、标高、索引、材料和专业条件检查问题图纸。"),
            ("实践任务布置", 10, "明确后9周绘图对象、成果清单、阶段节点、提交格式和评价标准。"),
        ],
        "discussion": "跨专业冲突应由哪个专业负责解决？建筑师应承担什么协调责任？",
        "homework": "确定实践项目，整理方案资料，建立电子文件夹、图纸目录初稿和问题清单。",
        "reflection": "检查学生是否掌握校审清单的使用方法，并为连续绘图实践准备完整基础资料。",
    },
]

# Append practical lesson definitions.
_practical = [
    (WEEK_TOPICS[9][1], ["明确实践项目、图纸成果和阶段进度要求。", "能够建立图纸目录、图号、图幅和出图比例体系。", "完成建筑设计说明、技术指标及主要表格初稿。", "形成持续修改和对成果完整性负责的习惯。"], "图纸框架和设计说明与方案的一致性。", "将零散方案信息转化为规范、准确的文字与表格。", "图纸目录、设计说明和主要表格初稿。"),
    (WEEK_TOPICS[10][1], ["掌握总平面图的定位、交通、消防、竖向和指标表达。", "能够完成内容较完整、比例合理的建筑总平面图。", "能够核查总平面与首层平面的对应关系。", "强化总体统筹和场地安全意识。"], "建筑定位、场地交通和竖向信息表达。", "完整处理场地关系并保持与建筑平面一致。", "建筑总平面图阶段成果。"),
    (WEEK_TOPICS[11][1], ["掌握平面施工图的轴网、墙柱、功能和门窗表达。", "完成首层及主要标准层的基本平面施工图。", "能够建立初步尺寸和编号体系。", "形成先定位、后深化的规范绘图习惯。"], "轴网、墙柱、门窗、楼梯和功能空间表达。", "由方案图转化为可定位的施工图基础。", "首层和主要标准层平面图初稿。"),
    (WEEK_TOPICS[12][1], ["完善各层平面图尺寸、标高、剖切和索引系统。", "处理各层上下对应、设备空间和专业条件预留。", "能够开展平面图自检、互检并根据意见修改。", "培养协同与持续改进意识。"], "尺寸链、标高、索引、各层一致性和专业条件。", "减少跨楼层、跨专业和跨图种矛盾。", "完整各层建筑平面图及自检清单。"),
    (WEEK_TOPICS[13][1], ["掌握建筑立面图的轮廓、构件、材料和标高表达。", "完成主要方向建筑立面图。", "能够核查立面与平面门窗和建筑形体的一致性。", "提升图面精确性和整体形象控制能力。"], "立面构件、材料分格、轴线和主要标高。", "保证立面图与平面图及建筑形体准确对应。", "主要建筑立面图阶段成果。"),
    (WEEK_TOPICS[14][1], ["掌握剖切位置选择和竖向空间表达方法。", "完成代表性建筑剖面图。", "能够校核层高、净高、楼梯、门窗和屋面关系。", "形成竖向设计和构造协调意识。"], "竖向尺寸、标高、空间和主要构造关系。", "剖面图与平面、立面准确对应。", "代表性建筑剖面图阶段成果。"),
    (WEEK_TOPICS[15][1], ["掌握大样图比例、材料、尺寸、连接和索引表达。", "完成若干重点部位建筑大样图。", "能够将设计意图深化为基本合理的构造节点。", "培养精益求精和工程实施意识。"], "构造层次、材料、尺寸、连接及索引关系。", "方案意图向合理构造节点的转化。", "楼梯、墙身、门窗或卫生间等大样图。"),
    (WEEK_TOPICS[16][1], ["掌握全套施工图自检和互检方法。", "能够核对设计说明、总平面、平立剖及大样之间的对应关系。", "形成问题清单并完成集中修改。", "强化主动校审和成果质量责任。"], "全套图纸完整性、准确性和一致性检查。", "系统发现并修正跨图纸问题。", "全套图纸自检表、互检表和修改版。"),
    (WEEK_TOPICS[17][1], ["完成教师终审意见落实和全套成果统一。", "掌握图纸整理、打印检查和电子归档要求。", "能够汇报设计深化过程、主要问题及修改依据。", "形成对最终成果负责和持续复盘的职业习惯。"], "终期校审、成果完整性及统一归档。", "在有限时间内系统修正遗留问题并清晰说明依据。", "完整课程成果、汇报文件和归档材料。"),
]

for idx, (title, goals, key, hard, deliverable) in enumerate(_practical, start=10):
    LESSONS.append({
        "title": title,
        "goals": goals,
        "key": key,
        "hard": hard,
        "process": [
            ("任务检查与问题梳理", 10, f"检查上周成果和本周准备材料，明确第{idx}周阶段目标、图面要求和提交节点。"),
            ("教师示范与共性问题讲解", 15, "结合案例示范关键绘制步骤、图层和标注方法，集中解释上一阶段发现的共性问题。"),
            ("学生连续绘图（一）", 25, "学生依据方案和规范独立推进图纸；教师巡回检查，重点核查内容完整性和关键技术关系。"),
            ("学生连续绘图（二）及一对一辅导", 30, "教师逐人进行图面批改，形成明确修改清单；学生现场修正主要问题并记录依据。"),
            ("集中讲评与阶段布置", 10, f"抽取典型成果讲评，明确课后需完成的修改和下周衔接任务。本周成果：{deliverable}"),
        ],
        "discussion": "本周图纸中最影响后续图种的基础性问题是什么，应该优先如何修改？",
        "homework": f"根据教师修改清单完成并提交：{deliverable}；同时更新个人施工图问题清单。",
        "reflection": "记录学生完成度、集中出现的制图或技术问题及个别辅导情况；下周开课前检查修改落实率并调整示范重点。",
    })

assert len(LESSONS) == 18


def set_run_font(run, size=10.5, bold=False, name="宋体"):
    run.font.name = name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name)
    run.font.size = Pt(size)
    run.bold = bold


def set_paragraph(p, size=10.5, bold=False, align=None, first_line=False, space_after=0, line=1.25):
    if align is not None:
        p.alignment = align
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = line
    if first_line:
        p.paragraph_format.first_line_indent = Pt(size * 2)
    for r in p.runs:
        set_run_font(r, size=size, bold=bold)
    return p


def set_cell_text(cell, text, size=9.5, bold=False, align=WD_ALIGN_PARAGRAPH.CENTER):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.0
    r = p.add_run(str(text))
    set_run_font(r, size=size, bold=bold)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    return cell


def set_cell_shading(cell, fill="D9EAF7"):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def set_cell_margins(cell, top=70, start=70, bottom=70, end=70):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = tcPr.first_child_found_in("w:tcMar")
    if tcMar is None:
        tcMar = OxmlElement("w:tcMar")
        tcPr.append(tcMar)
    for m, v in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tcMar.find(qn(f"w:{m}"))
        if node is None:
            node = OxmlElement(f"w:{m}")
            tcMar.append(node)
        node.set(qn("w:w"), str(v))
        node.set(qn("w:type"), "dxa")


def set_repeat_table_header(row):
    trPr = row._tr.get_or_add_trPr()
    tblHeader = OxmlElement("w:tblHeader")
    tblHeader.set(qn("w:val"), "true")
    trPr.append(tblHeader)


def set_fixed_layout(table):
    tblPr = table._tbl.tblPr
    layout = tblPr.find(qn("w:tblLayout"))
    if layout is None:
        layout = OxmlElement("w:tblLayout")
        tblPr.append(layout)
    layout.set(qn("w:type"), "fixed")


def configure_doc(doc, landscape=False, margins=(2.2, 2.0, 2.0, 2.0)):
    sec = doc.sections[0]
    if landscape:
        sec.orientation = WD_ORIENT.LANDSCAPE
        sec.page_width, sec.page_height = sec.page_height, sec.page_width
    sec.top_margin = Cm(margins[0])
    sec.bottom_margin = Cm(margins[1])
    sec.left_margin = Cm(margins[2])
    sec.right_margin = Cm(margins[3])
    styles = doc.styles
    styles["Normal"].font.name = "宋体"
    styles["Normal"]._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
    styles["Normal"].font.size = Pt(10.5)


def add_title(doc, text, size=18, bold=True, space_after=8):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(space_after)
    r = p.add_run(text)
    set_run_font(r, size=size, bold=bold, name="黑体" if bold else "宋体")
    return p


def add_body(doc, text, size=10.5, bold=False, first_line=True, align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=2):
    p = doc.add_paragraph()
    r = p.add_run(text)
    set_run_font(r, size=size, bold=bold)
    set_paragraph(p, size=size, bold=bold, align=align, first_line=first_line, space_after=space_after, line=1.35)
    return p


def add_heading_cn(doc, text, level=1):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6 if level == 1 else 3)
    p.paragraph_format.space_after = Pt(3)
    r = p.add_run(text)
    set_run_font(r, size=14 if level == 1 else 12, bold=True, name="黑体")
    return p


def add_bullets(doc, items: Iterable[str], size=10.5):
    for i, item in enumerate(items, 1):
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.4)
        p.paragraph_format.first_line_indent = Cm(-0.4)
        p.paragraph_format.space_after = Pt(1)
        r = p.add_run(f"{i}. {item}")
        set_run_font(r, size=size)
        p.paragraph_format.line_spacing = 1.25


def create_syllabus(path: Path):
    doc = Document()
    configure_doc(doc)
    add_title(doc, f"贵阳人文科技学院《{COURSE}》理论与实践课课程教学大纲", 16)
    add_heading_cn(doc, "第一部分 课程基本信息")
    table = doc.add_table(rows=6, cols=4)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"
    set_fixed_layout(table)
    info = [
        ("课程名称", COURSE, "开设单位", COLLEGE),
        ("总 学 时", "36", "课程学分", "2"),
        ("课程性质", "专业选修课", "先修课程", "计算机辅助设计、建筑设计基础Ⅰ、Ⅱ、建筑设计Ⅰ、建筑构造（上）"),
        ("理论学时", "18", "实践学时", "18"),
        ("适用学期", "第五学期", "考核方式", "考查"),
        ("执 笔 人", TEACHER, "审 核 人", "李保华"),
    ]
    for r, vals in enumerate(info):
        for c, val in enumerate(vals):
            set_cell_text(table.cell(r, c), val, size=9.5, bold=(c % 2 == 0))
            if c % 2 == 0:
                set_cell_shading(table.cell(r, c), "D9EAF7")
    for row in table.rows:
        row.cells[0].width = Cm(2.4); row.cells[1].width = Cm(5.3); row.cells[2].width = Cm(2.4); row.cells[3].width = Cm(7.0)

    add_heading_cn(doc, "一、课程性质")
    add_body(doc, "本课程是建筑学专业的专业选修课程，以规范作图、职业对接和工程实施为导向，围绕建筑施工图的编制逻辑、制图标准、绘制方法和校审要求展开教学。课程将建筑方案设计、建筑构造、建筑规范和计算机辅助制图等知识贯通起来，引导学生把既有建筑方案深化为内容较完整、表达准确、具有基本实施依据的建筑施工图成果。")
    add_body(doc, "课程采用理论讲授、案例示范、学生连续绘图、一对一点评修改、共性问题集中讲解和成果校审相结合的方式，强化学生的工程思维、规范意识、质量责任和严谨求实的职业态度，为后续综合建筑设计、毕业设计以及建筑设计岗位实践奠定基础。")

    add_heading_cn(doc, "二、课程目标")
    add_body(doc, "（一）知识目标", bold=True, first_line=False)
    add_bullets(doc, [
        "系统掌握建筑施工图设计的基本概念、组成内容、编排逻辑和房屋建筑制图标准。",
        "理解建筑设计说明、总平面图、平面图、立面图、剖面图和建筑大样图的设计要点与表达要求。",
        "了解建筑、结构、给排水、暖通和电气等专业在施工图阶段的基本配合关系以及校审要求。",
    ])
    add_body(doc, "（二）能力目标", bold=True, first_line=False)
    add_bullets(doc, [
        "能够依据既有建筑方案、现行规范和制图标准，建立图纸目录并编制建筑设计说明。",
        "能够完成建筑总平面图、各层平面图、主要立面图、代表性剖面图和重点部位大样图。",
        "能够分析方案向施工图转化过程中出现的功能、尺寸、构造和图纸表达问题，并通过持续修改形成较完整成果。",
        "能够依据校审提纲开展自检与互检，识别跨图纸、跨专业的常见错误并进行修正。",
    ])
    add_body(doc, "（三）素质目标", bold=True, first_line=False)
    add_bullets(doc, [
        "形成遵守规范、严谨作图、主动校审和精益求精的专业态度，理解施工图质量与工程安全、实施质量之间的关系。",
        "在一对一辅导、集中讲评和反复修改中形成主动沟通、持续改进和对设计成果负责的职业习惯。",
    ])

    add_heading_cn(doc, "三、指定参考书（教材）信息")
    add_body(doc, f"教材：{TEXTBOOK}；主编：{TEXTBOOK_AUTHOR}；出版社：{TEXTBOOK_PUBLISHER}；出版时间：{TEXTBOOK_DATE}。", first_line=False)
    add_heading_cn(doc, "四、其他参考资料")
    add_bullets(doc, ["现行国家及地方房屋建筑制图标准。", "现行国家及地方建筑设计、防火、无障碍、节能等相关规范。", "与建筑专业相关的结构、给排水、暖通和电气专业规范及施工图案例。"])

    doc.add_page_break()
    add_heading_cn(doc, "第二部分 课程教学内容与基本要求")
    chapters = [
        ("第一章 建筑工程设计阶段与施工图概述（4课时）",
         "了解建筑工程设计全过程以及施工图阶段的职责，掌握施工图的概念、作用、特点、设计深度、图纸组成和编排逻辑。",
         ["建筑工程设计阶段与基本工作流程。", "建筑工程法律制度、项目审批及设计人员责任。", "施工图的概念、作用、特点和设计深度。", "全套施工图组成、图纸目录、图号体系与编排顺序。"],
         "施工图阶段的工作内容及全套图纸的组成与编排。",
         "工程流程、设计责任与施工图成果之间的逻辑关系。"),
        ("第二章 房屋建筑制图规定（4课时）",
         "系统掌握房屋建筑制图的一般规定，能够规范识读和表达施工图中的图幅、图线、轴线、尺寸、标高、索引和符号等基本信息。",
         ["图纸幅面、标题栏、会签栏和图纸组织。", "图线、比例、字体、常用符号和图例。", "定位坐标、定位轴线、标高及编号方法。", "尺寸标注、索引符号、详图符号和剖切符号。", "常见制图错误辨析及不同图种的综合运用。"],
         "轴线、尺寸、标高、索引及图线系统的规范表达。",
         "将分散的制图规定综合运用于完整图纸并保持各图种一致。"),
        ("第三章 建筑专业施工图设计（6课时）",
         "掌握建筑专业施工图各组成文件的内容、绘制要求及图纸之间的对应关系，建立完整的施工图信息框架。",
         ["封面、图纸目录、建筑设计说明及主要技术指标。", "材料做法表、房间用料表、门窗表和专项说明。", "建筑总平面图的定位、交通、竖向及指标表达。", "建筑平面图的轴网、功能、尺寸、门窗、标高、索引及构造定位。", "建筑立面图和剖面图的形体、材料、标高、竖向空间及构造关系。", "总平面、平面、立面、剖面、说明及表格之间的一致性校核。"],
         "建筑设计说明和主要图种的内容、绘制方法与相互对应关系。",
         "将方案信息深化为具有基本实施依据且跨图纸一致的施工图信息。"),
        ("第四章 常用建筑详图（2课时）",
         "掌握常用建筑构件和重点部位的大样设计要点及规范表达方法。",
         ["楼梯、台阶、坡道及垂直交通构件。", "卫生间、厨房、设备用房等重点空间。", "门窗、墙身、屋面、顶棚等常用构造节点。", "大样比例、尺寸、材料、连接及索引关系。"],
         "楼梯、墙身、门窗等常用大样的构造与表达要点。",
         "从建筑设计意图推导合理构造，并在大样图中准确表达。"),
        ("第五章 多专业配合与施工图校审（2课时）",
         "了解施工图阶段的多专业配合、校对审核和审查要求，掌握常见冲突与错误的检查方法。",
         ["建筑与结构、给排水、暖通和电气专业的配合内容。", "管井、设备空间、洞口、标高及竖向系统协调。", "设计、校对和审核人员职责及审查流程。", "建筑专业施工图校审提纲和图纸对应检查。", "常见冲突类型、问题诊断及修改方法。"],
         "专业接口、校审程序和全套图纸检查要点。",
         "识别跨专业、跨图种的隐性矛盾并提出合理修改方案。"),
        ("第六章 施工图绘制实践（18课时）",
         "综合运用理论知识，将既有建筑方案转化为较完整的建筑施工图，并通过一对一辅导、自检、互检和终审持续修改。",
         ["图纸组织、图纸目录、建筑设计说明及主要表格。", "建筑总平面图绘制与修改。", "各层建筑平面图绘制、深化及专业条件协调。", "建筑立面图和剖面图绘制与对应校核。", "重点部位建筑大样图绘制。", "全套施工图自检、互检、统一修改、终审、成果展示与归档。"],
         "全套施工图的完整性、准确性、一致性和基本实施依据。",
         "系统发现并修正跨图纸、跨专业问题，形成完整成果。"),
    ]
    for title, purpose, contents, key, hard in chapters:
        add_heading_cn(doc, title)
        add_body(doc, "一、学习目的和要求", bold=True, first_line=False)
        add_body(doc, purpose)
        add_body(doc, "二、讲授或实践内容", bold=True, first_line=False)
        add_bullets(doc, contents)
        add_body(doc, "三、教学重点、难点", bold=True, first_line=False)
        add_body(doc, f"（一）重点：{key}", first_line=False)
        add_body(doc, f"（二）难点：{hard}", first_line=False)
        add_body(doc, "四、教学方法与手段", bold=True, first_line=False)
        add_body(doc, "采用多媒体讲授、规范导读、完整施工图案例拆解、教师示范、学生连续绘图、一对一辅导、集中讲评、自检互检和成果校审等方式。", first_line=False)

    add_heading_cn(doc, "第三部分 成绩考核评定及课时分配")
    add_heading_cn(doc, "一、成绩的考核与评定办法", level=2)
    add_body(doc, "本课程采用过程性考核和施工图实践成果考核相结合的方式。总成绩由平时成绩和作业实践成绩组成，其中平时成绩占30%，作业实践成绩占70%。平时成绩主要依据出勤、课堂参与、阶段任务完成、修改落实和自检互检情况评定；作业实践成绩依据设计说明及主要建筑施工图的完整性、规范性、准确性、一致性和修改质量评定。")
    add_heading_cn(doc, "二、授课内容课时分配", level=2)
    t = doc.add_table(rows=1, cols=5)
    t.style = "Table Grid"; t.alignment = WD_TABLE_ALIGNMENT.CENTER; set_fixed_layout(t)
    headers = ["周次", "授课内容", "理论课时", "实践课时", "备注"]
    for i, h in enumerate(headers):
        set_cell_text(t.cell(0, i), h, bold=True); set_cell_shading(t.cell(0, i));
    set_repeat_table_header(t.rows[0])
    for i, (kind, topic) in enumerate(WEEK_TOPICS, 1):
        row = t.add_row().cells
        vals = [i, topic, 2 if kind == "理论" else 0, 2 if kind == "实践" else 0, kind]
        for j, v in enumerate(vals):
            set_cell_text(row[j], v, size=9, align=WD_ALIGN_PARAGRAPH.LEFT if j == 1 else WD_ALIGN_PARAGRAPH.CENTER)
    total = t.add_row().cells
    set_cell_text(total[0], "合计", bold=True); total[0].merge(total[1])
    set_cell_text(total[2], "18", bold=True); set_cell_text(total[3], "18", bold=True); set_cell_text(total[4], "36学时", bold=True)
    doc.save(path)


def weekly_dates(semester_monday: date, weekday_index: int):
    # weekday_index: Monday=0, Wednesday=2
    first = semester_monday + timedelta(days=(weekday_index - semester_monday.weekday()) % 7)
    return [first + timedelta(weeks=i) for i in range(18)]


def create_progress(path: Path, class_name: str, dates: list[date], time_text: str, room: str):
    doc = Document()
    configure_doc(doc, landscape=True, margins=(1.2, 1.2, 1.2, 1.2))
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("附件3：建筑学教研室教学进度表")
    set_run_font(r, 15, True, "黑体")
    p.paragraph_format.space_after = Pt(4)
    info = doc.add_table(rows=1, cols=4)
    info.style = "Table Grid"; info.alignment = WD_TABLE_ALIGNMENT.CENTER
    for c, txt in enumerate(["课程名称", COURSE, "班级", class_name]):
        set_cell_text(info.cell(0, c), txt, size=10, bold=(c % 2 == 0))
        if c % 2 == 0: set_cell_shading(info.cell(0, c))
    doc.add_paragraph().paragraph_format.space_after = Pt(0)
    table = doc.add_table(rows=2, cols=10)
    table.style = "Table Grid"; table.alignment = WD_TABLE_ALIGNMENT.CENTER; set_fixed_layout(table)
    # grouped headers
    set_cell_text(table.cell(0,0), "理论课教学安排", 10, True); table.cell(0,0).merge(table.cell(0,6)); set_cell_shading(table.cell(0,0))
    set_cell_text(table.cell(0,7), "实验、实践课教学安排", 10, True); table.cell(0,7).merge(table.cell(0,9)); set_cell_shading(table.cell(0,7))
    headers = ["教师姓名", "职称", "周次", "年 月 日", "时数", "讲课章节及内容", "备注", "周次", "时数", "实验、实践内容"]
    for i, h in enumerate(headers):
        set_cell_text(table.cell(1,i), h, 8.5, True); set_cell_shading(table.cell(1,i), "EAF2F8")
    set_repeat_table_header(table.rows[0]); set_repeat_table_header(table.rows[1])
    for w, ((kind, topic), d) in enumerate(zip(WEEK_TOPICS, dates), 1):
        cells = table.add_row().cells
        vals = [TEACHER if w == 1 else "", TITLE if w == 1 else "", w, f"{d.year}.{d.month}.{d.day}", 2 if kind == "理论" else 0, topic if kind == "理论" else "", kind, w if kind == "实践" else "", 2 if kind == "实践" else "", topic if kind == "实践" else ""]
        for i, v in enumerate(vals):
            align = WD_ALIGN_PARAGRAPH.LEFT if i in (5,9) else WD_ALIGN_PARAGRAPH.CENTER
            set_cell_text(cells[i], v, 8.0 if i in (5,9) else 8.5, align=align)
            set_cell_margins(cells[i], 40, 45, 40, 45)
    # Merge teacher/title across rows for traditional template appearance.
    table.cell(2,0).merge(table.cell(19,0)); table.cell(2,1).merge(table.cell(19,1))
    set_cell_text(table.cell(2,0), TEACHER, 9, True); set_cell_text(table.cell(2,1), TITLE, 9, True)
    footer = doc.add_table(rows=2, cols=4)
    footer.style = "Table Grid"; footer.alignment = WD_TABLE_ALIGNMENT.CENTER
    totals = [("计划总学时", "36学时", "理论课学时", "18学时"), ("实践课学时", "18学时", "上课时间/地点", f"{time_text}，{room}")]
    for r, vals in enumerate(totals):
        for c, val in enumerate(vals):
            set_cell_text(footer.cell(r,c), val, 9, bold=(c%2==0), align=WD_ALIGN_PARAGRAPH.LEFT if c==3 else WD_ALIGN_PARAGRAPH.CENTER)
            if c%2==0: set_cell_shading(footer.cell(r,c), "EAF2F8")
    doc.save(path)


def add_cover(doc: Document, academic_year: str, semester: str, grade: str):
    for _ in range(3): doc.add_paragraph()
    add_title(doc, "贵阳人文科技学院", 22, True, 20)
    add_title(doc, "教  案", 30, True, 45)
    lines = [("课    程", COURSE), ("专    业", MAJOR), ("年    级", grade), ("教    师", TEACHER), ("职    称", TITLE)]
    for label, val in lines:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(12)
        r = p.add_run(f"{label}    {val}")
        set_run_font(r, 16, False, "宋体")
    for _ in range(3): doc.add_paragraph()
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(f"{academic_year}学年  第{semester}学期")
    set_run_font(r, 14, False)
    doc.add_page_break()


def add_homepage(doc: Document, class_name: str, time_text: str, academic_year: str, semester: str):
    add_title(doc, "贵阳人文科技学院教案", 16, True, 2)
    add_title(doc, "【首页】", 12, True, 6)
    t = doc.add_table(rows=9, cols=4); t.style = "Table Grid"; t.alignment = WD_TABLE_ALIGNMENT.CENTER; set_fixed_layout(t)
    rows = [
        ("院 系", COLLEGE, "教研室", DEPARTMENT),
        ("课程名称", COURSE, "授课年级及专业", class_name.replace("1班", "建筑学")),
        ("班级", "1班", "课程类型", "专业选修课（*）"),
        ("学生层次", "本科生（*）", "授课方式", "理论课（*）；实践课（*）"),
        ("考核方式", "考查", "课程教学总学时数", "36学时"),
        ("学分数", "2", "学时分配", "理论课18学时；实践课18学时"),
        ("教材名称", TEXTBOOK, "作者", TEXTBOOK_AUTHOR),
        ("出版社及出版时间", f"{TEXTBOOK_PUBLISHER}，{TEXTBOOK_DATE}", "授课时间", f"{time_text}（每节45分钟，共90分钟）"),
        ("参考资料", "现行建筑制图、建筑设计、防火、无障碍、节能及相关专业规范；完整施工图案例。", "学年学期", f"{academic_year}学年第{semester}学期"),
    ]
    for r, vals in enumerate(rows):
        for c, val in enumerate(vals):
            set_cell_text(t.cell(r,c), val, 9.2, bold=(c%2==0), align=WD_ALIGN_PARAGRAPH.LEFT if c%2 else WD_ALIGN_PARAGRAPH.CENTER)
            if c%2==0: set_cell_shading(t.cell(r,c), "EAF2F8")
            set_cell_margins(t.cell(r,c), 80, 90, 80, 90)
    doc.add_paragraph()
    p = doc.add_paragraph()
    r = p.add_run("注：教案按授课次数填写，每次授课均填写一份；重复班可根据实际授课日期分别归档。")
    set_run_font(r, 9)
    doc.add_page_break()


def create_lesson_plan(path: Path, class_name: str, grade: str, dates: list[date], time_text: str, room: str, academic_year: str, semester: str):
    doc = Document()
    configure_doc(doc, margins=(1.7, 1.7, 1.8, 1.8))
    add_cover(doc, academic_year, semester, grade)
    add_homepage(doc, class_name, f"{time_text}，{room}", academic_year, semester)
    for week, (lesson, d) in enumerate(zip(LESSONS, dates), 1):
        t = doc.add_table(rows=6, cols=4); t.style = "Table Grid"; t.alignment = WD_TABLE_ALIGNMENT.CENTER; set_fixed_layout(t)
        # Header row
        set_cell_text(t.cell(0,0), "授课章节", 9.5, True); set_cell_shading(t.cell(0,0))
        set_cell_text(t.cell(0,1), lesson["title"], 9.5, True, WD_ALIGN_PARAGRAPH.LEFT)
        set_cell_text(t.cell(0,2), "授课时长", 9.5, True); set_cell_shading(t.cell(0,2))
        set_cell_text(t.cell(0,3), "2课时（90分钟）", 9.5)
        # Date row
        set_cell_text(t.cell(1,0), "授课日期", 9.5, True); set_cell_shading(t.cell(1,0))
        set_cell_text(t.cell(1,1), f"{d.year}年{d.month:02d}月{d.day:02d}日", 9.5)
        set_cell_text(t.cell(1,2), "周次", 9.5, True); set_cell_shading(t.cell(1,2))
        set_cell_text(t.cell(1,3), week, 9.5)
        # Goals merged
        set_cell_text(t.cell(2,0), "教学目标", 9.5, True); set_cell_shading(t.cell(2,0))
        goal_cell = t.cell(2,1).merge(t.cell(2,3)); goal_cell.text = ""
        for i, g in enumerate(lesson["goals"], 1):
            p = goal_cell.add_paragraph() if i > 1 else goal_cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0); p.paragraph_format.line_spacing=1.15
            r = p.add_run(f"{i}. {g}"); set_run_font(r, 9.2)
        # Key/difficult
        set_cell_text(t.cell(3,0), "教学重难点", 9.5, True); set_cell_shading(t.cell(3,0))
        kd = t.cell(3,1).merge(t.cell(3,3)); kd.text=""
        p=kd.paragraphs[0]; r=p.add_run(f"教学重点：{lesson['key']}"); set_run_font(r,9.2,True)
        p=kd.add_paragraph(); r=p.add_run(f"教学难点：{lesson['hard']}"); set_run_font(r,9.2)
        # Process
        set_cell_text(t.cell(4,0), "教学内容与过程", 9.5, True); set_cell_shading(t.cell(4,0))
        pc = t.cell(4,1).merge(t.cell(4,3)); pc.text=""
        p=pc.paragraphs[0]; r=p.add_run("一、教学组织（每节45分钟，共90分钟）"); set_run_font(r,9.2,True)
        for i,(name,mins,desc) in enumerate(lesson["process"],1):
            p=pc.add_paragraph(); p.paragraph_format.space_after=Pt(0); p.paragraph_format.line_spacing=1.1
            r=p.add_run(f"{i}. {name}（{mins}分钟）：{desc}"); set_run_font(r,8.9)
        p=pc.add_paragraph(); r=p.add_run("二、教学方法与手段："); set_run_font(r,9.2,True)
        r=p.add_run(" 多媒体讲授、完整施工图案例分析、教师示范、课堂讨论、学生绘图、一对一辅导和集中讲评。" if week <= 9 else " 任务驱动、学生连续绘图、教师巡回指导、一对一图面批改、同伴互检和共性问题集中讲评。")
        set_run_font(r,8.9)
        p=pc.add_paragraph(); r=p.add_run(f"三、课堂讨论/实践检查：{lesson['discussion']}"); set_run_font(r,8.9)
        p=pc.add_paragraph(); r=p.add_run(f"四、作业及阶段成果：{lesson['homework']}"); set_run_font(r,8.9)
        # Reflection
        set_cell_text(t.cell(5,0), "课后反思", 9.5, True); set_cell_shading(t.cell(5,0))
        rc = t.cell(5,1).merge(t.cell(5,3)); set_cell_text(rc, f"反思重点：{lesson['reflection']}\n\n实际教学记录：____________________________________________________________________\n__________________________________________________________________________________", 8.8, False, WD_ALIGN_PARAGRAPH.LEFT)
        # widths and margins
        for row in t.rows:
            for cell in row.cells:
                set_cell_margins(cell, 70, 80, 70, 80)
        t.cell(0,0).width=Cm(2.4); t.cell(0,1).width=Cm(11.5); t.cell(0,2).width=Cm(2.4); t.cell(0,3).width=Cm(3.0)
        if week != 18:
            doc.add_page_break()
    doc.save(path)


def main():
    # Common 36-hour syllabus.
    create_syllabus(OUT / "施工图设计_课程教学大纲_36学时18周.docx")

    dates_2020 = weekly_dates(date(2023, 8, 21), 2)  # Wednesday
    dates_2021 = weekly_dates(date(2024, 2, 26), 0)  # Monday

    create_progress(OUT / "2020级建筑学1班_施工图设计_教学进度表_18周.docx", "2020级建筑学1班", dates_2020, "周三第3、4节", "自信515")
    create_progress(OUT / "2021级建筑学1班_施工图设计_教学进度表_18周.docx", "2021级建筑学1班", dates_2021, "周一第7、8节", "自信519")

    create_lesson_plan(OUT / "2020级建筑学1班_施工图设计_教案_18周完整版.docx", "2020级建筑学1班", "2020级", dates_2020, "周三第3、4节", "自信515", "2023—2024", "一")
    create_lesson_plan(OUT / "2021级建筑学1班_施工图设计_教案_18周完整版.docx", "2021级建筑学1班", "2021级", dates_2021, "周一第7、8节", "自信519", "2023—2024", "二")

    print("Generated:")
    for p in sorted(OUT.glob("*.docx")):
        print(p, p.stat().st_size)


if __name__ == "__main__":
    main()
