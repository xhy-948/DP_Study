# -*- coding: utf-8 -*-
"""生成《差分隐私学习总结》Word 文档"""
import docx
from docx import Document
from docx.shared import Pt, Cm, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

OUT = r'E:\研究生\差分隐私学习\差分隐私学习总结.docx'

doc = Document()

# ---------- 页面设置 ----------
sec = doc.sections[0]
sec.page_width = Cm(21.0)
sec.page_height = Cm(29.7)
sec.top_margin = Cm(2.5)
sec.bottom_margin = Cm(2.5)
sec.left_margin = Cm(2.5)
sec.right_margin = Cm(2.5)

# ---------- 样式工具 ----------
def set_run_font(run, cn_font='宋体', en_font='Arial', size=12, bold=False, color=RGBColor(0, 0, 0)):
    run.font.name = en_font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.append(rFonts)
    rFonts.set(qn('w:ascii'), en_font)
    rFonts.set(qn('w:hAnsi'), en_font)
    rFonts.set(qn('w:eastAsia'), cn_font)

def para_fmt(p, align=WD_ALIGN_PARAGRAPH.JUSTIFY, line=1.5, first_indent=True,
             before=0, after=0, keep_with_next=False):
    pf = p.paragraph_format
    pf.alignment = align
    pf.line_spacing = line
    if first_indent:
        pf.first_line_indent = Pt(24)  # 12pt*2字符
    else:
        pf.first_line_indent = Pt(0)
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.keep_with_next = keep_with_next
    return p

# Normal 样式
normal = doc.styles['Normal']
normal.font.name = 'Arial'
normal.font.size = Pt(12)
normal.font.color.rgb = RGBColor(0, 0, 0)
rpr = normal.element.get_or_add_rPr()
rfonts = rpr.find(qn('w:rFonts'))
if rfonts is None:
    rfonts = OxmlElement('w:rFonts'); rpr.append(rfonts)
rfonts.set(qn('w:ascii'), 'Arial'); rfonts.set(qn('w:hAnsi'), 'Arial'); rfonts.set(qn('w:eastAsia'), '宋体')

# 标题样式定制
def style_heading(name, size):
    st = doc.styles[name]
    st.font.name = 'Arial'
    st.font.size = Pt(size)
    st.font.bold = True
    st.font.color.rgb = RGBColor(0, 0, 0)
    rpr = st.element.get_or_add_rPr()
    rf = rpr.find(qn('w:rFonts'))
    if rf is None:
        rf = OxmlElement('w:rFonts'); rpr.append(rf)
    rf.set(qn('w:ascii'), 'Arial'); rf.set(qn('w:hAnsi'), 'Arial'); rf.set(qn('w:eastAsia'), '黑体')
    # 段落间距
    pf = st.paragraph_format
    pf.keep_with_next = True
    return st

h1 = style_heading('Heading 1', 16)
h1.paragraph_format.space_before = Pt(14)
h1.paragraph_format.space_after = Pt(6)
h2 = style_heading('Heading 2', 14)
h2.paragraph_format.space_before = Pt(11)
h2.paragraph_format.space_after = Pt(5)
h3 = style_heading('Heading 3', 12)
h3.paragraph_format.space_before = Pt(9)
h3.paragraph_format.space_after = Pt(4)

# Title 样式
title_st = doc.styles['Title']
title_st.font.name = 'Arial'
title_st.font.size = Pt(18)
title_st.font.bold = True
title_st.font.color.rgb = RGBColor(0, 0, 0)
rpr = title_st.element.get_or_add_rPr()
rf = rpr.find(qn('w:rFonts'))
if rf is None:
    rf = OxmlElement('w:rFonts'); rpr.append(rf)
rf.set(qn('w:ascii'), 'Arial'); rf.set(qn('w:hAnsi'), 'Arial'); rf.set(qn('w:eastAsia'), '黑体')
title_st.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
title_st.paragraph_format.space_after = Pt(18)

# ---------- 正文工具 ----------
def H1(text):
    p = doc.add_paragraph(style='Heading 1')
    r = p.add_run(text)
    set_run_font(r, cn_font='黑体', size=16, bold=True)
    return p

def H2(text):
    p = doc.add_paragraph(style='Heading 2')
    r = p.add_run(text)
    set_run_font(r, cn_font='黑体', size=14, bold=True)
    return p

def H3(text):
    p = doc.add_paragraph(style='Heading 3')
    r = p.add_run(text)
    set_run_font(r, cn_font='黑体', size=12, bold=True)
    return p

def P(text, indent=True, bold_prefix=None):
    p = doc.add_paragraph()
    para_fmt(p, first_indent=indent)
    if bold_prefix:
        r0 = p.add_run(bold_prefix)
        set_run_font(r0, size=12, bold=True)
    r = p.add_run(text)
    set_run_font(r, size=12)
    return p

def bullet(text, bold_prefix=None):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.left_indent = Pt(24)
    pf.line_spacing = 1.5
    pf.first_line_indent = Pt(0)
    r0 = p.add_run('• ')
    set_run_font(r0, size=12, bold=True)
    if bold_prefix:
        rb = p.add_run(bold_prefix)
        set_run_font(rb, size=12, bold=True)
    r = p.add_run(text)
    set_run_font(r, size=12)
    return p

def caption(text):
    p = doc.add_paragraph()
    para_fmt(p, align=WD_ALIGN_PARAGRAPH.CENTER, first_indent=False)
    r = p.add_run(text)
    set_run_font(r, size=10.5)
    return p

def set_cell_bg(cell, color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color)
    tcPr.append(shd)

def add_table(headers, rows, widths=None, header_bg='D9D9D9'):
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = 'Table Grid'
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = t.rows[0].cells
    for i, h in enumerate(headers):
        set_cell_bg(hdr[i], header_bg)
        hdr[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = hdr[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.first_line_indent = Pt(0)
        p.paragraph_format.line_spacing = 1.0
        r = p.add_run(h)
        set_run_font(r, size=10.5, bold=True)
    for row in rows:
        cells = t.add_row().cells
        for i, val in enumerate(row):
            cells[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cells[i].paragraphs[0]
            p.paragraph_format.first_line_indent = Pt(0)
            p.paragraph_format.line_spacing = 1.0
            short = len(str(val)) <= 12
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if short else WD_ALIGN_PARAGRAPH.LEFT
            r = p.add_run(str(val))
            set_run_font(r, size=10.5)
    if widths:
        for i, w in enumerate(widths):
            for row in t.rows:
                row.cells[i].width = Cm(w)
    return t

# ---------- 封面标题 ----------
p = doc.add_paragraph(style='Title')
r = p.add_run('差分隐私学习总结')
set_run_font(r, cn_font='黑体', size=18, bold=True)

p = doc.add_paragraph()
para_fmt(p, align=WD_ALIGN_PARAGRAPH.CENTER, first_indent=False)
r = p.add_run('基于 Dwork–Roth 专著与两篇中文综述，及 2025 年拆分学习应用论文的综合整理')
set_run_font(r, cn_font='楷体', size=12)
p.paragraph_format.space_after = Pt(12)

# ---------- 目录（域） ----------
p = doc.add_paragraph()
para_fmt(p, align=WD_ALIGN_PARAGRAPH.CENTER, first_indent=False)
r = p.add_run('目  录')
set_run_font(r, cn_font='黑体', size=16, bold=True)

toc_p = doc.add_paragraph()
fldChar1 = OxmlElement('w:fldChar'); fldChar1.set(qn('w:fldCharType'), 'begin')
instrText = OxmlElement('w:instrText'); instrText.set(qn('xml:space'), 'preserve')
instrText.text = 'TOC \\o "1-2" \\h \\z \\u'
fldChar2 = OxmlElement('w:fldChar'); fldChar2.set(qn('w:fldCharType'), 'separate')
t = OxmlElement('w:t'); t.text = '（在 Word/WPS 中打开时自动更新目录）'
fldChar3 = OxmlElement('w:fldChar'); fldChar3.set(qn('w:fldCharType'), 'end')
run = toc_p.add_run()
run._r.append(fldChar1); run._r.append(instrText); run._r.append(fldChar2)
run._r.append(t); run._r.append(fldChar3)

doc.add_page_break()

# =========================================================
# 1 学习材料与知识全景
# =========================================================
H1('1  学习材料与知识全景')
P('本总结围绕四份差分隐私（Differential Privacy，DP）学习材料整理，覆盖从理论定义、实现机制到数据发布、数据挖掘与机器学习应用，再到最新工程落地案例的完整知识链。材料角色不同、互为补充：Dwork–Roth 专著提供权威理论基础，两篇 2014 年中文综述分别从“数据发布/数据挖掘”和“发布与分析技术细节”两个视角展开全景综述，2025 年的 IEEE 论文则展示了 DP 与拆分学习结合的前沿应用。', bold_prefix='')

H2('1.1  四份材料总览')
add_table(
    ['编号', '材料', '类型/来源', '核心内容'],
    [
        ['A1', 'Iqbal, Gope, Sikdar. Privacy-Preserving Collaborative Split Learning Framework for Smart Grid Load Forecasting', '期刊论文，IEEE TDSC 2025 年 22 卷 6 期', '面向智能电网负荷预测的双层拆分学习框架；FEDformer 骨干；SplitGlobal/SplitPersonal 两种训练策略；用互信息估计（MINE）与差分隐私分析激活泄露'],
        ['A2', 'Dwork, Roth. The Algorithmic Foundations of Differential Privacy', '专著，Foundations and Trends in Theoretical Computer Science, 2014', '差分隐私理论奠基之作：定义、组合定理、拉普拉斯/指数机制、查询发布、下界、机制设计、机器学习、本地模型等 13 章'],
        ['A3', '熊平, 朱天清, 王晓峰. 差分隐私保护及其应用', '中文综述，计算机学报 2014, 37(1): 101–113', 'DP 基础理论（定义、敏感度、组合性质、机制）；隐私保护数据发布（交互式/非交互式）；隐私保护数据挖掘（接口/完全访问模式）'],
        ['A4', '张啸剑, 孟小峰. 面向数据发布和分析的差分隐私保护', '中文综述，计算机学报 2014, 37(4): 927–949', '直方图发布（策略一/策略二）、划分发布（树/网格）、频繁模式/分类/回归/SVM 等数据分析方法、批量查询处理、原型系统'],
    ],
    widths=[1.2, 4.6, 4.0, 6.2],
)
caption('表 1  四份学习材料总览')

H2('1.2  材料间关系与推荐学习顺序')
P('四份材料构成“理论→方法→应用”三级结构：A2 回答“差分隐私是什么、为什么成立”（定义、随机化必要性、组合理论、下界）；A3 与 A4 回答“差分隐私能用来做什么”（发布、挖掘、机器学习），两者分别侧重两条应用主线的机制对比与技术细节分类；A1 回答“差分隐私在真实场景中如何落地”（把 ε-DP、(ε,δ)-DP 与互信息估计用于拆分学习的激活保护）。')
P('推荐学习顺序：先读 A2 第 1–3 章建立严格定义与基本机制，再并行读 A3、A4 掌握两大应用方向的全景与代表算法，最后精读 A1 验证理论如何转化为工程权衡（如 ε 的取值对隐私强度与模型精度的折衷）。')

# =========================================================
# 2 差分隐私基础理论
# =========================================================
H1('2  差分隐私基础理论')
P('本节依据 A2（Dwork–Roth）、A3（熊平等）、A4（张啸剑等）整理差分隐私的定义体系。理解要点：差分隐私是一种“定义”而非某个具体算法（A2 明确指出）；它不依赖攻击者的背景知识假设，而是通过输出随机化保证“任一记录加入或不加入数据集，对输出分布的影响都被 ε 界住”。')

H2('2.1  动机：为什么需要差分隐私')
P('早期的 k-匿名及其扩展模型（l-diversity、t-closeness、(α,k)-anonymity、M-invariance 等）属于基于分组的模型：通过对准标识符泛化/压缩，把记录隐藏到等价类中。A3 指出这类模型有两个根本缺陷：其一，安全性依赖攻击者背景知识假设，而“合成式攻击”“前景知识攻击”“deFinetti 攻击”等新型攻击不断出现，模型需反复修补；其二，无法提供严格、可量化的隐私保护水平证明。')
P('A2 用多个经典反例说明“删除标识符”或“发布汇总统计”并不安全：链接攻击可将“匿名”医疗记录与选民登记记录匹配（马萨诸塞州州长案例），也可把 Netflix 匿名观影记录与 IMDb 关联；差分攻击通过两个大体量查询的差即可推断个体属性（“数据库中有多少人携带镰状细胞性状”与“除 X 外有多少人……”）；查询审计不可行（拒绝回答本身可能泄露信息）；基因组关联研究发布等位基因频率也被证实可判定个体是否在病例组，美国 NIH 与 Wellcome Trust 因此终止公开访问聚合频率数据。')
P('差分隐私正是在这一背景下提出的：在最大背景知识假设下（攻击者已知除目标记录外所有信息），DP 依然能控制单条记录对查询结果的影响，从而中和链接攻击等各类攻击，且隐私保护水平可被 ε 定量比较。')

H2('2.2  形式化定义')
P('设数据集 D 与 D′ 具有相同属性结构，其对称差 |D△D′|=1 时称 D 与 D′ 为邻近数据集（相邻数据库）。A3 定义 1：对任意随机算法 M 及其任意输出子集 S_M，若')
P('Pr[M(D)∈S_M] ≤ exp(ε)·Pr[M(D′)∈S_M]', indent=False)
P('恒成立，则称 M 提供 ε-差分隐私保护，参数 ε 称为隐私保护预算。直观含义：删除或加入任一记录，算法输出同一结果的概率之比不超过 e^ε，攻击者无法通过观察输出来判定个体是否在数据集中。')
P('A2 的形式化定义（定义 2.4）采用更一般的 (ε,δ)-差分隐私：对任意 ||x−y||₁≤1 的数据库对 x、y，Pr[M(x)∈S] ≤ exp(ε)·Pr[M(y)∈S]+δ。δ 允许小概率（最坏情形）违约，ε 控制平均隐私损失；δ=0 时退化为纯 ε-DP。A1 指出，由于 (ε,δ)-DP 拥有更优的高级组合定理，能在相同 ε 下支持更多训练迭代，当前差分隐私机器学习研究大多转向 (ε,δ)-DP。')
P('随机化是必要条件：A2 用混合论证（hybrid argument）证明，任何非平凡的确定性算法都无法在任意辅助信息（其他数据库、网站、新闻等）下提供隐私保证——总存在一对仅差一行、却输出不同的相邻数据库，攻击者即可判定未知行的取值。')

H2('2.3  隐私保护预算与敏感度')
bullet('ε 越小隐私保护越强；ε=0 时两个邻近数据集输出分布完全相同，结果不再携带数据信息。实际中 ε 常取 0.01、0.1、ln2、ln3 等小值，需在隐私强度与数据可用性之间权衡（A3）。')
bullet('全局敏感度：GS_f = max_{D,D′} ||f(D)−f(D′)||₁，即删除任一记录对查询结果的最大改变，决定所需噪声量。计数查询敏感度为 1，只需少量噪声；中位数、均值等函数敏感度可能很大（如区间 [a,b] 上极端取值使中位数敏感度达 b−a），直接加噪会淹没真实结果（A3）。')
bullet('局部敏感度 LS_f(D) 利用具体数据的分布特征，通常远小于全局敏感度；但直接使用局部敏感度加噪会泄露数据分布信息，故引入 β-平滑上界 S(D)（满足 S(D)≥LS_f(D) 且 S(D)≤e^β S(D′)），据此构造平滑敏感度 S_{f,β}(D)=max(LS_f(D′)·e^{−βd(D,D′)}) 来计算噪声（A3，Nissim 等提出）。')

H2('2.4  组合性质')
P('复杂任务需多次调用 DP 算法，隐私预算需在步骤间分配，此时依赖两条基本组合性质（A3 性质 1、2；A4 性质 1、2）：')
bullet('序列组合：对同一数据集依次运行预算为 ε₁,…,ε_L 的 L 个算法，组合算法满足 (Σε_i)-差分隐私。预算逐次累加，因此“预算耗尽”后算法不再满足 DP。')
bullet('并行组合：对互不相交的数据集 D₁,…,D_L 分别运行预算为 ε_i 的算法，组合算法满足 (max ε_i)-差分隐私。可用于将预算按数据分区复用，提高利用率（PINQ 的 Partition 算子即利用此性质）。')
P('A4 进一步总结了常用预算分配策略：线性分配、均匀分配、指数分配、自适应分配与混合策略分配，目标是让 ε 的“生命周期”尽可能长。')

H2('2.5  基本实现机制')
add_table(
    ['机制', '适用对象', '实现方式', '噪声量/特点'],
    [
        ['拉普拉斯机制 (Laplace)', '数值型查询结果', 'M(D)=f(D)+Y，Y~Lap(Δf/ε)', '绝对误差 √2Δf/ε，方差 2(Δf)²/ε²；ε 越小噪声越大'],
        ['指数机制 (Exponential)', '非数值型结果（方案/选择/实体）', '以正比于 exp(εq(D,r)/(2Δq)) 的概率输出 r，q 为可用性打分函数', 'ε 大时放大高分项概率（如 ε=1 时最高票项目概率达 0.924）；ε=0 时退化为均匀选择'],
        ['高斯机制 (Gaussian)', '(ε,δ)-DP 场景', 'M(D)=f(D)+η，η~N(0, 2s²/ε²·log(2/δ))', '配合高级组合定理，支持更多次迭代；A1 用于 (ε,δ)-DP 实验'],
    ],
    widths=[3.0, 3.0, 4.8, 5.2],
)
caption('表 2  三种基本差分隐私机制')

P('随机化响应（randomized response）是另一经典基础技术（A2 第 3 章）：受访者以一定概率如实回答、以其余概率回答反值，分析者用无偏估计还原总体比例，个体答案本身不可信。')

H2('2.6  核心要点小结')
bullet('差分隐私是定义而非算法，任何满足该输出分布约束的机制都可被称为 DP 机制。')
bullet('DP 与攻击者背景知识无关：即使攻击者已知除目标记录外的一切，个体隐私仍受保护。')
bullet('预算 ε 与敏感度 Δf 是两大核心参数：ε 决定隐私强度，Δf 决定所需噪声规模。')
bullet('组合性质是算法设计与预算分配的理论基础：串行累加、并行取最大。')
bullet('(ε,δ)-DP 通过引入 δ 换取高级组合定理，是当前 DP 机器学习的主流形式。')

# =========================================================
# 3 差分隐私数据发布
# =========================================================
H1('3  差分隐私数据发布')
P('数据发布是 DP 研究最核心的应用方向，目标是在满足差分隐私的条件下，以尽量高的精度发布查询结果或“净化”后的数据集。依据实现环境可分为交互式与非交互式两类（A3 图 5、A4 第 4 节）。')

H2('3.1  交互式与非交互式框架')
bullet('交互式（在线查询）：用户逐次提交查询，数据管理者加噪后返回；查询次数受预算约束，预算耗尽即失效。代表性机制有 Laplace、Median、PMW、K-norm、IDC 等。')
bullet('非交互式（离线发布）：一次性发布所有可查询的结果，或发布净化数据集供任意查询。研究重点是设计满足 DP 且可用性高的发布算法，常借助数据压缩、转换、采样过滤等手段。')
P('A3 给出交互式发布机制的精度度量方式：以精确度 α（Pr[|f(D)−M(f(D))|≤α]≥1−δ）与对应样本复杂度评估机制性能。')

H2('3.2  交互式发布机制比较')
add_table(
    ['机制', '可响应查询数量级', '计算效率', '精确度（误差界）'],
    [
        ['Median（中位数机制）', 'n 的指数', '非高效', 'n^{1/2}(log k)(log|Z|)^{1/2}/ε^{1/2}'],
        ['PMW（私有乘性权重）', 'n 的指数', '非高效', 'n^{1/2}(log k)(log|Z|)^{1/2}/ε^{1/2}'],
        ['K-norm（K 范数机制）', 'n 的次线性', '非高效', '√(k/ε)·(log(|Z|/k))^{1/2}'],
        ['IDC（迭代数据生成框架）', 'n 的指数', '非高效', 'n^{1/2}(log k)^{3/2}(log|Z|)^{1/2}/ε^{1/2}'],
    ],
    widths=[3.2, 3.2, 2.6, 7.0],
)
caption('表 3  交互式发布机制比较（A3 表 3）')
P('要点：Laplace 机制简单通用但查询数量受限且噪声大；Median 通过把查询分为“难/易”两类节省预算；PMW 借助加权多数投票思想、仅在差异超阈值时才消耗预算；K-norm 机制把噪声复杂度刻画为高维凸体几何属性；IDC 框架是 Median 与 PMW 的统一架构。这些机制大多计算代价高，是理论分析多于工程实现的典型。')

H2('3.3  直方图发布')
P('直方图是 DP 发布最常用的数据结构：删除一条记录最多影响一个（或少数）桶的计数，敏感度易于计算。A4 将直方图发布方法归纳为两种策略。')
H3('3.3.1  策略一：先加噪、后优化')
P('直接为每个桶计数添加 Laplace 噪声，再通过后置处理（一致性约束、二次规划）提升精度。代表方法：')
bullet('LP：等宽直方图逐桶加噪 Ĥ=H+Lap(1/ε)，响应长范围查询时噪声累加，方差 2n/ε²，误差界 O(n/ε²)，可用性低。')
bullet('Boost1：在一致性约束下求 min||H−Ĥ||₂（最小二乘），误差界 O(log³n/ε²)，精度显著提升，但仅支持一维直方图。')
bullet('NoiseFirst：借鉴 V-优化直方图（动态规划合并相邻相似桶、最小化平方误差和 SSE），同时考虑重构误差与噪声误差，支持较长范围查询，误差界 O((n−k)/ε²)。')
bullet('DPCube：先按单元加噪、再用 k-d 树后置处理获得多维 V-优化直方图，支持多维查询，但精度不稳定、误差大（O(dn²/ε²)）。')
H3('3.3.2  策略二：先转换、后加噪')
P('先重组直方图结构降低敏感度再加噪，能更精确响应长范围查询，但转换带来信息缺损。转换方式分三类：')
bullet('层次树变换：Boost2 用 m-ary 树（敏感度 1+log_m n）；Privlet 用哈尔小波变换（Δf=1+log₂ n），每个查询由 O(log n) 个小波系数的线性组合表示，噪声方差 O((log n)²/ε²)。')
bullet('聚类变换：StructureFirst 用 V-优化直方图合并邻近相似桶，把敏感度从 Δf 降到 Δf/p，并以指数机制选择合并边界；P-HPartition 用自适应层次聚类（贪婪二分+指数机制）平衡重构误差与噪声误差，Error(Ck)=RECk+k/ε。')
bullet('傅里叶变换：FPA 对直方图做离散傅里叶变换后只保留 k 个低频系数加噪再逆变换，误差由噪声误差（2k²/ε²）与重组误差构成；EFPA 用指数机制自适应挑选 k。')
P('对比结论（A4）：策略一方法中 Boost1、NoiseFirst 一维效果较好；策略二方法中树变换与傅里叶变换只支持一维、可比性差；聚类变换可启发式发布但未处理离群点与计数顺序。支持多维、携带离群点、均衡两种误差是未来方向。')

H2('3.4  批查询与列联表发布')
P('批查询指一次性发布所有可能查询的结果。由于查询间相关，删除一条记录可能影响多个查询，批查询集合的敏感度远高于单一查询（A3 示例：n 个变量上所有范围查询的集合敏感度可达 6 量级）。')
add_table(
    ['方法', '总敏感度', '每个查询噪声方差'],
    [
        ['Lap1（逐查询独立加噪）', 'O(n²)', 'O(n⁴/ε²)'],
        ['Lap2（对单点加噪后推算）', '1', 'O(n/ε²)'],
        ['Privelet（小波变换）', 'log n + 1', 'O((log n)²/ε²)'],
        ['层次查询（层次树）', 'log n + 1', 'O((log n)²/ε²)'],
    ],
    widths=[5.6, 3.0, 4.6],
)
caption('表 4  批查询的敏感度与噪声方差对比（A3 表 4）')
P('列联表（列表）发布有两类做法：对每个单元格加噪（保持边缘频数一致性但噪声大）或对边缘频数直接加噪（噪声小但可能破坏一致性）；Barak 等提出傅里叶域加噪的整体方案以兼顾准确性与一致性。')

H2('3.5  分组发布与净化数据集')
P('分组发布把 k-匿名思想与 DP 结合：安全 k-匿名模型通过对原始数据随机抽样增加不确定性，理论证明当抽样函数满足 ε-DP 时可提供 (ε,δ)-DP；DiffGen 采用“自顶向下、逐步细分”策略，用指数机制选择细分方案生成满足 DP 的决策树（A3）。')
P('净化数据集发布指直接发布加噪后的数据集供任意查询，是公认难题（A3 引 Dinur–Nissim：精确回答超过线性个查询即可高概率还原数据集）。学习理论给出了突破：Kasiviswanathan 等定义差分隐私 PAC 学习并证明——无隐私可学的概念类在 DP 下同样可学；Blum 等证明用指数机制遍历数据域可找到满足 DP 的净化数据集。但此类方法计算复杂度为超多项式，难以实用。')

H2('3.6  基于划分的发布')
P('A4 系统梳理了基于树结构与网格结构的划分发布方法，两类都需区分数据依赖划分（划分结构本身可能泄露隐私）与数据独立划分。')
bullet('树划分：Quad-Post（完全四分树，预算均匀/几何分配，利用序列+并行组合性质）；kd-standard（指数机制保护中位数分割线，预算两份：选分割点+计数加噪）；DiffPart（分类树分割，Laplace 扰动非叶节点计数，自适应预算分配，用二项分布控制空节点）；另有 kd-noisemean、DP-tree、n-gram、Hybrid-Bus 等。')
bullet('网格划分：UG 均匀划分 m×m 格，均衡噪声误差 √(2rm/ε) 与均匀假设误差 √(rN/mc₀)，取 m=√(N/c)；AG 自适应两层划分，先用 αε 粗分、再对稠密单元细分，避免过稀（噪声大）与过密（均匀假设误差大）问题。')
P('共同挑战：树划分的停止条件、高维扩展、动态数据支持；网格划分的启发式稠密/稀疏边界判定。')

H2('3.7  发布技术小结')
P('发布技术的主线是“在隐私与可用性之间权衡”：降低敏感度（转换、压缩、分区）与合理分配预算（组合性质、自适应策略）是两大抓手；后置处理（一致性约束、最小二乘）是普遍有效的精度提升手段。')

# =========================================================
# 4 差分隐私数据挖掘与机器学习
# =========================================================
H1('4  差分隐私数据挖掘与机器学习')

H2('4.1  两种实现模式')
bullet('接口模式：数据管理者不发布原始数据，仅提供加噪查询接口（SuLQ、PINQ），挖掘者无需掌握 DP 知识，查询数受预算限制。')
bullet('完全访问模式：挖掘者可信、可直接访问数据，但需改造算法使其满足 DP（随机决策树、差分逻辑回归、FIM 等）。')
P('SuLQ 框架把布尔查询定义为查询原语，并扩展到连续数据与 k-means、ID3 等复杂算法；PINQ 基于 LINQ 语言，其 Partition 算子利用并行组合性提高预算利用率。')

H2('4.2  分类算法')
add_table(
    ['算法', '实现机制', '噪声', '数据类型', '特点'],
    [
        ['SuLQ-based ID3', 'Laplace 机制', '高', '离散', '逐计数加噪算信息增益；ε<1 时准确率较无隐私 ID3 降约 30%'],
        ['PINQ-based ID3', 'Laplace 机制', '高', '离散', 'Partition 算子分割数据，但每属性须单独预算，噪声仍大'],
        ['DiffPID3', 'Laplace+指数机制', '低', '离散', '一次指数机制评估全部属性，单次分裂只耗一次预算'],
        ['DiffP-C4.5', 'Laplace+指数机制', '低', '离散或连续', '连续属性先选分裂点再选分裂方案，每次迭代调用两次指数机制'],
    ],
    widths=[3.0, 3.0, 1.6, 2.6, 5.8],
)
caption('表 5  基于接口模式的分类算法比较（A3 表 6）')
P('完全访问模式下，Jagannathan 等的随机决策树方法先随机构建树结构（与数据无关），再对叶节点类标签计数向量 V（敏感度为 1）加 Lap(1/ε) 噪声，在 UCI 三个数据集上平均分类准确率超过 85%。Chaudhuri 等提出差分隐私逻辑回归：把噪声加入目标函数而非直接加在模型输出上，且正则化参数越大敏感度越低、所需噪声越小。')

H2('4.3  聚类算法')
P('SuLQ 框架下的差分 k-means 通过发布加噪的质心与计数满足隐私，但质心查询的敏感度是聚类最大直径，噪声过大。改进路线有三条：')
bullet('局部敏感度+抽样聚合框架（Nissim）：随机抽样 m 个子集分别求中间结果，聚合后按平滑敏感度加噪。')
bullet('核心集（Coreset，Feldman）：构造容量 O(kd^{−1}log n) 的加权点集替代原数据，输出满足 (1+α)-近似，实现差分 k-median/k-means。')
bullet('预算分配（Dwork）：迭代次数已知时每轮预算 ε/u；未知时首次 ε/2、之后逐轮减半。')

H2('4.4  频繁项集挖掘')
P('频繁项集本身及其频度可能泄露用户信息。代表方法：')
bullet('FIM（截断支持度）：先用传统算法得到 l-项集，定义截断支持度 ŝ=max(s, sₖ−γ) 避免遍历全部项集，再分别用指数机制（K 次调用选 K 个项集）或 Laplace 机制（对 ŝ 加噪取 Top-K）输出，最后对真实支持度加噪。K 较大（如 >100）时性能与准确性显著下降。')
bullet('PrivBasis：利用“频繁项集的子集必频繁”性质，用极大团构造 θ-基集覆盖所有高频项集，再统计 1-、2-项集支持度重构，避免遍历指数级候选，适合高维数据。')
bullet('TF（截断频率）与 SmartTrunc：TF 以概率 Pr(p)∝exp(εnf′(p)/4k) 从候选中选模式；SmartTrunc 通过截断长事务记录降低查询敏感性，但只对短事务占多数的数据集有效。')

H2('4.5  回归分析')
P('回归分析的目标是学习权重向量 w* = argmin Σ f(t_i, w)，直接发布 w* 会泄露数据。A4 比较了四类方法：')
bullet('LPLog：对 w* 直接加 Laplace 噪声，但因输入 x_i 与预测 ρ(x_i) 强关联，w* 敏感性计算代价高、预测精度低。')
bullet('ObjectivePerb：对目标函数均值加噪（f̃_D(w)=Σf/n + b^T w/n）再求 w*，噪声由 w* 敏感度决定，仅适用于强约束凸双可微目标函数，标准逻辑回归（布尔 y）下失效。')
bullet('FM 函数机制：把目标函数展开为多项式 Σλ_ji φ_j(w)，噪声量由目标函数自身敏感性决定（‖f_D−f_D′‖₁≤2max Σ‖λ_ji‖₁），绕过 w* 敏感性；逻辑回归目标函数用截断泰勒展开近似为多项式。')
bullet('DPME：用直方图技术生成加噪训练集再回归，绕开 w* 敏感性分析，但仅适合低维训练集。')

H2('4.6  支持向量机')
P('PrivateSVM 对法向量 w* 加 Laplace 噪声（w*+Lap(λ)），因 w* 敏感性过高导致噪声大、分类精度低；ObjectiveSVM 对风险函数 f_D(w) 加噪后再优化，精度更高，但要求目标函数凸且双可微。设计通用的扰动机制是开放问题。')

H2('4.7  原型系统')
add_table(
    ['系统', '架构特点', '状态攻击', '预算攻击', '定时攻击'],
    [
        ['PINQ（微软）', '最早交互式 DP 分析系统，LINQ API', '否', '否', '否'],
        ['GUPT（伯克利）', '抽样-聚合框架扩展，强制访问控制 MAC，按数据时效自动设预算', '是', '是', '是'],
        ['Airavat（德州大学）', 'MapReduce 框架，Mapper 用 MAC、Reducer 加噪', '否', '是', '否'],
        ['Pioneer（新加坡 ADS）', '查询优化器选择最省预算的执行计划', '否', '是', '否'],
    ],
    widths=[3.4, 7.2, 1.8, 1.8, 1.8],
)
caption('表 6  四种原型系统及其对三类攻击的防御（A4 表 9）')
P('A4 总结了差分隐私原型系统可能面临的三种攻击：状态攻击（修改系统静态变量）、隐私预算攻击（耗尽预算）、定时攻击（无限请求同一任务）。PINQ 由不可信分析者管理预算导致浪费；GUPT 在三种攻击下防御最好。')

H2('4.8  数据挖掘小结')
P('接口模式让数据挖掘者无需理解隐私，但查询次数与预算受限；完全访问模式灵活但要求算法改造。总体上，指数机制常优于直接 Laplace（预算利用率更高、噪声更小）；“扰动目标函数而非输出”是回归与 SVM 领域的主流改进思路；高维数据（大 K、长序列、多属性）仍是各类方法共同的性能瓶颈。')

# =========================================================
# 5 前沿案例精读
# =========================================================
H1('5  前沿案例精读：拆分学习+差分隐私的智能电网负荷预测')
P('A1（Iqbal 等，IEEE TDSC 2025）是四份材料中唯一的工程落地论文，完整示范了“差分隐私概念→算法设计→隐私量化→开销分析”的研究闭环，也是理解 A2–A4 理论如何被实际调用的最佳案例。')

H2('5.1  背景与问题')
P('智能电表高分辨率负荷数据可泄露用户隐私：何时在家、日常作息、特定电器乃至医疗设备使用情况均可被推断，受 GDPR 等法规约束。传统集中式训练需要把数据汇聚到中心服务器，通信开销大且隐私风险高；联邦学习（FL）让客户端训练整个模型，对计算能力有限的智能电表不现实；单客户端独立建模又因客户数量大而不经济。论文提出用拆分学习（SL）在“全局单模型”与“每客户端单模型”之间取得平衡。')

H2('5.2  方法：双层拆分学习框架')
P('系统模型为三层四实体：服务提供商（SP）→ 多个电网站点（GS，各管一个邻域）→ 智能电表客户端。网络采用 FEDformer（频域增强分解 Transformer）并双拆分：Split-1（数据嵌入、前两个编码器内部块与部分解码器）部署在 GS，Split-2（其余编码器块+完整编码器块+其余解码器块）部署在 SP。')
P('关键设计：智能电表只用自己的数据做 Split-1 前向传播、计算损失并发起反向传播；GS 负责 Split-1 的反向传播与参数更新；SP 维护 Split-2。这样既大幅降低客户端计算负担，又因 Split-1 保持输入-输出间复杂非线性关系而天然降低信息泄露（相比过度简化的客户端子网）。FEDformer 的频域增强块（FEB/FEA）用离散傅里叶变换+随机模式选择实现线性复杂度注意力。')
P('训练采用两种策略：SplitGlobal（每个 GS 有各自个性化 Split-1，SP 训练一个全局 Split-2）与 SplitPersonal（SP 为每个 GS 训练个性化 Split-2）。训练以 7 个步骤为一个回合：GS 下发 Split-1 权重→客户端前向并上传激活→GS 拼接激活发给 SP→SP 前向→输出经 GS 送回客户端算损失→梯度回传→SP/GS 更新参数（训练全程标签不离开客户端）。')

H2('5.3  实验设置与结果')
P('数据集：Electricity 数据集（320 个智能电表 2016-07 至 2019-07 小时级负荷），取前 17,566 条训练，归一化后用凝聚聚类分为 3 个邻域（54/201/65 个客户端）。设置：输入长度 L=96 小时、预测视界 O=96 小时，模型维度 D=512，每邻域 K=10 个客户端，评价指标 MAE/MSE/R²，训练 10 epoch+早停。')
add_table(
    ['模型', 'GS1 (MAE/MSE/R²)', 'GS2 (MAE/MSE/R²)', 'GS3 (MAE/MSE/R²)', '均值 (MAE/MSE/R²)'],
    [
        ['SplitGlobal', '0.451 / 0.387 / 0.396', '0.256 / 0.133 / 0.863', '0.372 / 0.256 / 0.663', '0.360 / 0.259 / 0.641'],
        ['SplitPersonal', '0.453 / 0.383 / 0.405', '0.253 / 0.130 / 0.868', '0.361 / 0.241 / 0.687', '0.356 / 0.251 / 0.653'],
        ['集中式 Central', '0.513 / 0.540 / 0.346', '0.246 / 0.133 / 0.855', '0.355 / 0.252 / 0.671', '0.371 / 0.308 / 0.624'],
        ['FedAvg', '0.535 / 0.554 / 0.342', '0.268 / 0.142 / 0.838', '0.359 / 0.273 / 0.641', '0.387 / 0.323 / 0.607'],
    ],
    widths=[2.6, 3.5, 3.3, 3.3, 3.3],
)
caption('表 7  各模型测试集分数（A1 表 II）')
P('结论：所提模型整体达到或超过集中式模型，SplitPersonal 均值最优；在负荷模式最多样、最难预测的邻域 1，集中式模型明显更差，说明“一个全局模型覆盖所有分布”不可行。集中式模型训练误差最低但验证/测试误差最高，存在过拟合。跨邻域预测中 SplitGlobal 泛化更好（其全局 Split-2 见过所有邻域数据）；对完全未见过的客户端，邻域 1 误差上升约 66%，追加 5 个 epoch 训练后降至约 35%；每轮随机选取客户端训练比固定客户端泛化更好。单 epoch 耗时：SL 模型约 20 分钟、集中式 40 分钟、FL 28 分钟。')

H2('5.4  隐私分析：互信息与差分隐私')
P('隐私泄露定量化：用互信息神经估计（MINE，3 层全连接网络 100/50/1 神经元、ELU 激活）估计输入与 Split-1 激活间的互信息。基线：输入与自身的 MI 约 9.12（理论上界），输入与纯 Laplace 噪声的 MI 约 1.83（下界）；干净激活的 MI 约 3.32，加噪激活约 2.6。结论：由于 Split-1 保持了输入-激活间的复杂非线性关系，即使不加噪声泄露也很有限，且客户端可实时估算 MI 决定是否安全上传当前批次激活。')
P('差分隐私实验：对 Split-1 激活施加 Laplace 机制（ε-DP）与高斯机制（(ε,δ)-DP）。ε=0.5 时 MI 降至 1.86（接近纯噪声水平），代价是 MAE 比无 DP 高约 36%；ε=2.5（中等隐私）时 MAE/MSE 仅高 16%/26%；ε≥5 时 MI 停滞在约 3.3、误差与无 DP 相当。固定 ε=1.0 改变 δ∈[0.001,0.1] 时，MI 与误差几乎不变，印证 δ 主要控制最坏情形而非平均隐私损失。')
P('抗攻击验证：在比诚实好奇更强的最坏假设下，用特征空间劫持攻击（FSHA）从激活重建原始数据。无 DP 时攻击收敛快、重建 MSE 低至 0.241（且框架本身仍有约 0.305 的重建下限，说明结构设计提供了天然防护）；施加 DP（ε=3.0, δ=0.005）后攻击显著变难，重建 MSE 升至 0.328。代价是训练时间从约 200 分钟增至约 454 分钟（约 2.54 倍，主要来自 Diffprivlib 噪声计算）。')

H2('5.5  计算与通信开销')
P('客户端每轮上传约 6.3 MB 激活、接收约 3 KB 预测输出，反向传播仅回传 3 KB 梯度；重负载的约 63 MB 梯度张量在 GS 与 SP 之间交换，客户端通信量仅占约 4.55%。计算方面，客户端只承担约 12.66% 的前向操作（完整更新轮中约 5%），反向传播与参数更新全部由 GS/SP 完成；延迟方面客户端约占端到端的约 21%。论文指出 SplitPersonal 相对集中式基线获得 18.5% 的性能提升，同时把能耗负担转移给电力公司侧基础设施，适合智能电表等资源受限设备。')

H2('5.6  论文启示')
bullet('DP 落地存在清晰权衡曲线：ε≥5 时隐私增强几乎不损失精度，ε=2.5 时以 16%~26% 的误差增幅换取显著隐私提升——工程上应结合合规要求选择预算。')
bullet('结构设计本身也是隐私机制：通过保持 Split-1 输入-输出非线性复杂度，框架在无噪声情况下已具备一定抗重建能力，DP 是叠加的“额外安全层”。')
bullet('量化工具（MINE 互信息）让隐私保护从“定性声明”变为“可在线决策的度量”，客户端可按需决定是否上传，是值得借鉴的隐私工程思路。')

# =========================================================
# 6 研究展望与开放问题
# =========================================================
H1('6  研究展望与开放问题')
P('综合 A4 的总结与 A1 的讨论，差分隐私当前与未来的主要挑战可归纳如下。')
bullet('动态数据发布：现有方法多针对静态快照；数据流（滑动窗口下的直方图发布）、增量更新（集值数据反复添加/删除记录）会导致噪声累积与预算耗尽，需要合理的更新边界与预算分配策略。')
bullet('图数据发布：社交网络等图数据稀疏，边-差分隐私下的 triangle、k-stars 等查询敏感度大、可用性低；邻接矩阵+双聚类挖掘密集区域是可行的探索方向。')
bullet('分布式差分隐私：云数据库等分布式环境中，站点可能与云服务商串通导致密钥泄露；需要 DP 与加密技术结合，防止聚集信息泄露，并支持分布式数据挖掘（如 Heavy hitter）。')
bullet('大数据分析：个性化推荐、社交网络分析等场景要求流式采样、直方图与图概要技术与 DP 结合，同时控制可预测性带来的隐私风险。')
bullet('理论难点：ε 的定量度量困难、普通用户难以选择合适的预算；DP 假设记录相互独立，对家庭、社交关系等关联数据保护不足。')
bullet('工程部署（A1）：实际电力系统部署面临通信延迟、设备异构、同步问题与安全数据交换等挑战，是规模化落地的关键障碍。')

# =========================================================
# 7 学习收获与下一步建议
# =========================================================
H1('7  学习收获与下一步建议')

H2('7.1  关键公式速查表')
add_table(
    ['名称', '公式', '说明'],
    [
        ['ε-差分隐私', 'Pr[M(D)∈S] ≤ e^ε·Pr[M(D′)∈S]', 'D、D′ 为邻近数据集（差一条记录）'],
        ['(ε,δ)-差分隐私', 'Pr[M(x)∈S] ≤ e^ε·Pr[M(y)∈S] + δ', 'δ 控制最坏情形违约概率'],
        ['全局敏感度', 'GS_f = max ‖f(D)−f(D′)‖₁', '决定噪声规模；计数查询为 1'],
        ['Laplace 机制', 'M(D)=f(D)+Lap(Δf/ε)', '数值型；方差 2(Δf)²/ε²'],
        ['指数机制', 'Pr[r] ∝ exp(εq(D,r)/(2Δq))', '非数值型；q 为可用性函数'],
        ['高斯机制', 'M(D)=f(D)+N(0, 2s²/ε²·log(2/δ))', '(ε,δ)-DP；支持高级组合'],
        ['序列组合', 'ε = Σε_i', '同一数据集多次调用'],
        ['并行组合', 'ε = max ε_i', '不相交数据集分别调用'],
        ['R²（A1）', 'R² = 1 − Σ(y_i−ŷ_i)²/Σ(y_i−E[y_i])²', '拟合优度，越大越好'],
        ['互信息（A1）', 'I(X;Y)=D_KL(P_XY‖P_X⊗P_Y)', '用神经网络估计（MINE）'],
    ],
    widths=[2.6, 6.6, 4.6],
)
caption('表 8  关键公式速查')

H2('7.2  四份材料的不可替代价值')
bullet('A2（Dwork–Roth）：唯一的“定义级”材料，解释 DP 为什么成立（随机化必要性）、理论边界在哪（下界、重建攻击），并延伸到机制设计、本地模型等高级主题，是后续理论阅读的锚点。')
bullet('A3（熊平等）：对基础概念（敏感度体系、组合性质）与两大应用线（PPDR/PPDM）的全景梳理最清晰，交互式机制对比表与接口/完全访问模式框架便于快速建立地图。')
bullet('A4（张啸剑等）：对发布技术（直方图两种策略、树/网格划分）与数据分析方法（回归、SVM、频繁项集）的分类最细致，含原型系统与三类攻击的分析，适合做方法索引。')
bullet('A1（Iqbal 2025）：唯一给出“隐私量化+参数权衡+攻击验证+开销分析”完整实证的材料，展示 DP 从理论到工程的全部环节，也提示了与 FL/SL 等分布式学习结合的活跃方向。')

H2('7.3  下一步学习路径建议')
bullet('理论深化：精读 A2 第 3 章（组合定理与稀疏向量技术）、第 4–6 章（线性查询发布与提升方法）、第 8 章（下界与重建攻击）、第 11–12 章（DP 机器学习与本地模型）。')
bullet('工具实践：动手使用 Diffprivlib（A1 所用）、Google Opacus、TensorFlow Privacy 等库，复现 Laplace/高斯机制与 DP-SGD，把 ε 与噪声的权衡做成可交互实验。')
bullet('应用方向：关注差分隐私与联邦学习/拆分学习的结合（A1 路线）、本地差分隐私（LDP，对应 A2 第 12 章）、以及推荐系统、医疗数据、智能电网等行业的合规落地案例。')
bullet('写作与复现：以 A1 为模板练习“方法→实验→隐私分析→开销分析”的论文结构；复现 MINE 互信息估计与 FSHA 攻击，理解隐私攻击与防御的对抗逻辑。')

# ---------- 参考文献 ----------
doc.add_page_break()
H1('参考文献')
refs = [
    '[1] C. Dwork, A. Roth. The Algorithmic Foundations of Differential Privacy. Foundations and Trends in Theoretical Computer Science, 2014, 9(3-4): 211-407.',
    '[2] 熊平, 朱天清, 王晓峰. 差分隐私保护及其应用. 计算机学报, 2014, 37(1): 101-113.',
    '[3] 张啸剑, 孟小峰. 面向数据发布和分析的差分隐私保护. 计算机学报, 2014, 37(4): 927-949.',
    '[4] A. Iqbal, P. Gope, B. Sikdar. Privacy-Preserving Collaborative Split Learning Framework for Smart Grid Load Forecasting. IEEE Transactions on Dependable and Secure Computing, 2025, 22(6): 6487-6502. DOI: 10.1109/TDSC.2025.3385297.',
    '[5] C. Dwork. Differential Privacy. ICALP 2006: 1-12.',
    '[6] C. Dwork, F. McSherry, K. Nissim, A. Smith. Calibrating Noise to Sensitivity in Private Data Analysis. TCC 2006: 363-385.',
    '[7] F. McSherry, K. Talwar. Mechanism Design via Differential Privacy. FOCS 2007: 94-103.',
    '[8] K. Nissim, S. Raskhodnikova, A. Smith. Smooth Sensitivity and Sampling in Private Data Analysis. STOC 2007: 75-84.',
    '[9] N. Holohan, S. Braghin, P. Mac Aonghusa, K. Levacher. Diffprivlib: The IBM Differential Privacy Library. arXiv:1907.02424, 2019.',
    '[10] M. Y. S. Uddin et al. MINE: Mutual Information Neural Estimation. arXiv:1801.04062, 2018.',
]
for ref in refs:
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.left_indent = Pt(0)
    pf.first_line_indent = Pt(0)
    pf.line_spacing = 1.5
    pf.space_after = Pt(3)
    # 悬挂缩进：首行 0，左缩进 2 字符
    pf.left_indent = Pt(24)
    pf.first_line_indent = Pt(-24)
    r = p.add_run(ref)
    set_run_font(r, size=10.5)

# ---------- 页脚页码 ----------
footer = sec.footer
fp = footer.paragraphs[0]
fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
fld1 = OxmlElement('w:fldChar'); fld1.set(qn('w:fldCharType'), 'begin')
instr = OxmlElement('w:instrText'); instr.set(qn('xml:space'), 'preserve'); instr.text = 'PAGE'
fld2 = OxmlElement('w:fldChar'); fld2.set(qn('w:fldCharType'), 'end')
run = fp.add_run()
run._r.append(fld1); run._r.append(instr); run._r.append(fld2)
set_run_font(run, size=10.5)

# ---------- 打开时自动更新域（目录） ----------
settings = doc.settings.element
upd = OxmlElement('w:updateFields')
upd.set(qn('w:val'), 'true')
settings.append(upd)

doc.save(OUT)
print('saved:', OUT)
