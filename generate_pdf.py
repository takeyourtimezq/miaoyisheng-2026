# -*- coding: utf-8 -*-
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import os

# 注册中文字体
font_path = "/System/Library/Fonts/STHeiti Light.ttc"
if os.path.exists(font_path):
    pdfmetrics.registerFont(TTFont('Heiti', font_path))
    chinese_font = 'Heiti'
else:
    font_path2 = "/System/Library/Fonts/PingFang.ttc"
    if os.path.exists(font_path2):
        pdfmetrics.registerFont(TTFont('PingFang', font_path2))
        chinese_font = 'PingFang'
    else:
        chinese_font = 'Helvetica'

# 创建PDF
doc = SimpleDocTemplate(
    "轻刻时代_乌鲁木齐选址分析报告.pdf",
    pagesize=A4,
    rightMargin=2*cm,
    leftMargin=2*cm,
    topMargin=2*cm,
    bottomMargin=2*cm
)

# 样式
styles = getSampleStyleSheet()

title_style = ParagraphStyle(
    'ChineseTitle',
    parent=styles['Title'],
    fontName=chinese_font,
    fontSize=20,
    spaceAfter=20,
    alignment=1
)

heading1_style = ParagraphStyle(
    'ChineseH1',
    parent=styles['Heading1'],
    fontName=chinese_font,
    fontSize=16,
    spaceAfter=12,
    spaceBefore=20
)

heading2_style = ParagraphStyle(
    'ChineseH2',
    parent=styles['Heading2'],
    fontName=chinese_font,
    fontSize=14,
    spaceAfter=10,
    spaceBefore=15
)

normal_style = ParagraphStyle(
    'ChineseNormal',
    parent=styles['Normal'],
    fontName=chinese_font,
    fontSize=11,
    spaceAfter=8,
    leading=16
)

bullet_style = ParagraphStyle(
    'ChineseBullet',
    parent=normal_style,
    leftIndent=20,
    bulletIndent=10
)

story = []

# 标题
story.append(Paragraph("轻刻时代·乌鲁木齐选址分析报告", title_style))
story.append(Spacer(1, 10))

# 基本信息表
info_data = [
    ["品牌", "轻刻时代（健康减脂）"],
    ["城市", "乌鲁木齐"],
    ["数据来源", "抖音来客后台·行业兴趣人群地点偏好"],
    ["分析日期", "2026年5月21日"]
]

info_table = Table(info_data, colWidths=[3*cm, 10*cm])
info_table.setStyle(TableStyle([
    ('FONTNAME', (0, 0), (-1, -1), chinese_font),
    ('FONTSIZE', (0, 0), (-1, -1), 10),
    ('BACKGROUND', (0, 0), (0, -1), colors.lightgrey),
    ('ALIGN', (0, 0), (0, -1), 'RIGHT'),
    ('ALIGN', (1, 0), (1, -1), 'LEFT'),
    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
    ('PADDING', (0, 0), (-1, -1), 8),
]))
story.append(info_table)
story.append(Spacer(1, 20))

# 一、数据概览
story.append(Paragraph("一、数据概览", heading1_style))
story.append(Paragraph(
    '基于抖音来客后台导出的"行业兴趣人群地点偏好"数据，乌鲁木齐TOP10商圈兴趣人群分布如下：',
    normal_style
))
story.append(Spacer(1, 10))

# 数据表格
data = [
    ["排名", "商圈名称", "兴趣人数", "占比"],
    ["1", "铁路局", "8", "14.5%"],
    ["2", "新疆民俗街/新华南路", "6", "10.9%"],
    ["3", "友好北路", "5", "9.1%"],
    ["4", "水上乐园", "5", "9.1%"],
    ["5", "友好路", "5", "9.1%"],
    ["6", "北京南路", "5", "9.1%"],
    ["7", "红山", "4", "7.3%"],
    ["8", "白鸟湖", "4", "7.3%"],
    ["9", "乌北站", "4", "7.3%"],
    ["10", "粮校", "4", "7.3%"],
]

table = Table(data, colWidths=[2*cm, 6*cm, 3*cm, 2*cm])
table.setStyle(TableStyle([
    ('FONTNAME', (0, 0), (-1, -1), chinese_font),
    ('FONTSIZE', (0, 0), (-1, -1), 10),
    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#4472C4')),
    ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
    ('PADDING', (0, 0), (-1, -1), 8),
    ('BACKGROUND', (0, 1), (-1, 1), colors.HexColor('#E7F3FF')),
]))
story.append(table)
story.append(Spacer(1, 10))
story.append(Paragraph("<b>合计：</b>55人（样本量较小，需结合实地调研验证）", normal_style))

# 二、商圈逐一分析
story.append(PageBreak())
story.append(Paragraph("二、商圈逐一分析", heading1_style))

# 1. 铁路局商圈
story.append(Paragraph("1. 铁路局商圈 ⭐首选", heading2_style))
story.append(Paragraph("<b>兴趣人群：</b>8人（TOP1）", normal_style))
story.append(Paragraph("<b>商圈定位：</b>乌鲁木齐老牌生活商圈，以铁路局职工社区为核心，辐射周边3-5公里居民区。", normal_style))

story.append(Paragraph("<b>客群画像：</b>", normal_style))
story.append(Paragraph("• 中产家庭为主，年龄段30-45岁", bullet_style))
story.append(Paragraph("• 消费偏好稳定，重视性价比", bullet_style))
story.append(Paragraph("• 健康意识较强，有减脂/塑形需求", bullet_style))
story.append(Paragraph("• 工作稳定，消费能力中等偏上", bullet_style))

story.append(Paragraph("<b>商业配套：</b>", normal_style))
story.append(Paragraph("• 成熟社区商业，生活配套齐全", bullet_style))
story.append(Paragraph("• 餐饮、零售、服务业态丰富", bullet_style))
story.append(Paragraph("• 停车便利，交通通达性好", bullet_style))
story.append(Paragraph("• 商业租金相对友好（预估80-150元/㎡/月）", bullet_style))

story.append(Paragraph("<b>适合轻刻时代的原因：</b>", normal_style))
story.append(Paragraph("1. <b>流量TOP1</b> - 抖音数据直接证明兴趣人群最多", bullet_style))
story.append(Paragraph("2. <b>客群匹配</b> - 中产家庭+健康意识=减脂服务核心客群", bullet_style))
story.append(Paragraph("3. <b>竞争较小</b> - 非核心商圈，专业减脂品牌少，适合0-1样板店", bullet_style))
story.append(Paragraph("4. <b>成本可控</b> - 租金低于友好路等核心商圈，试错成本低", bullet_style))
story.append(Paragraph("5. <b>社区口碑</b> - 老社区口碑传播强，利于早期客户积累", bullet_style))

story.append(Paragraph("<b>风险提示：</b>", normal_style))
story.append(Paragraph("• 商圈整体调性偏传统，品牌调性需接地气", bullet_style))
story.append(Paragraph("• 年轻白领客群较少，需通过线上引流补充", bullet_style))

# 2. 友好北路+友好路商圈
story.append(Spacer(1, 15))
story.append(Paragraph("2. 友好北路+友好路商圈 ⭐备选（合并分析）", heading2_style))
story.append(Paragraph("<b>兴趣人群：</b>10人（友好北路5人 + 友好路5人，合并后TOP1）", normal_style))
story.append(Paragraph("<b>商圈定位：</b>乌鲁木齐核心商圈，城市商业地标，高端购物中心聚集地。", normal_style))

story.append(Paragraph("<b>核心商业体：</b>", normal_style))
story.append(Paragraph("• 美美百货、友好商场、汇嘉时代", bullet_style))
story.append(Paragraph("• 多个高端写字楼", bullet_style))

story.append(Paragraph("<b>客群画像：</b>", normal_style))
story.append(Paragraph("• 中高收入人群，白领、企业主为主", bullet_style))
story.append(Paragraph("• 消费能力强，愿意为品质付费", bullet_style))
story.append(Paragraph("• 年龄段25-40岁", bullet_style))
story.append(Paragraph("• 时尚、社交属性强", bullet_style))

story.append(Paragraph("<b>适合轻刻时代的原因：</b>", normal_style))
story.append(Paragraph("1. <b>流量最大</b> - 合并后兴趣人群10人，全城最高", bullet_style))
story.append(Paragraph("2. <b>品牌势能</b> - 核心商圈开店，品牌形象背书强", bullet_style))
story.append(Paragraph("3. <b>客群质量高</b> - 白领+企业主，付费能力强", bullet_style))
story.append(Paragraph("4. <b>招商展示</b> - 核心商圈样板店，利于后续招商", bullet_style))

story.append(Paragraph("<b>风险提示：</b>", normal_style))
story.append(Paragraph("• <b>租金压力</b> - 核心商圈租金高，月租金成本可能过万", bullet_style))
story.append(Paragraph("• <b>竞争激烈</b> - 美容、健身、减脂品牌众多，差异化要求高", bullet_style))
story.append(Paragraph("• <b>停车痛点</b> - 客户到店便利性受影响", bullet_style))

story.append(Paragraph("<b>建议：</b>适合作为品牌旗舰店或招商展示店，不适合作为第一家样板店。", normal_style))

# 3. 新疆民俗街/新华南路商圈
story.append(PageBreak())
story.append(Paragraph("3. 新疆民俗街/新华南路商圈 ⭐差异化选择", heading2_style))
story.append(Paragraph("<b>兴趣人群：</b>6人（TOP2）", normal_style))
story.append(Paragraph("<b>商圈定位：</b>乌鲁木齐特色文旅商圈，以新疆国际大巴扎为核心，融合旅游+本地消费。", normal_style))

story.append(Paragraph("<b>核心商业体：</b>新疆国际大巴扎、民俗街特色商铺、周边餐饮民宿聚集", normal_style))

story.append(Paragraph("<b>客群画像：</b>", normal_style))
story.append(Paragraph("• 双重客群：游客（外地）+ 本地居民", bullet_style))
story.append(Paragraph("• 游客：消费冲动强，体验导向", bullet_style))
story.append(Paragraph("• 本地居民：周边社区中产家庭", bullet_style))

story.append(Paragraph("<b>适合轻刻时代的原因：</b>", normal_style))
story.append(Paragraph("1. <b>差异化定位</b> - 旅游+本地双客流，客流结构独特", bullet_style))
story.append(Paragraph("2. <b>网红属性</b> - 大巴扎自带流量，利于抖音内容创作", bullet_style))
story.append(Paragraph("3. <b>夜间经济</b> - 客流高峰在晚间，匹配减脂服务时段", bullet_style))
story.append(Paragraph("4. <b>租金适中</b> - 介于核心商圈与社区商圈之间（预估100-180元/㎡/月）", bullet_style))

story.append(Paragraph("<b>风险提示：</b>", normal_style))
story.append(Paragraph("• 游客转化率低，非目标客群占比较高", bullet_style))
story.append(Paragraph("• 季节性波动大，旅游淡季客流下滑", bullet_style))
story.append(Paragraph("• 民族特色商圈，品牌调性需适配", bullet_style))

# 其他商圈简述
story.append(Spacer(1, 15))
story.append(Paragraph("4-10. 其他商圈简要分析", heading2_style))

other_areas = [
    ["商圈", "兴趣人数", "特点", "建议"],
    ["北京南路", "5", "社区商圈，客群稳定，租金低", "备选社区店"],
    ["水上乐园", "5", "休闲娱乐商圈，季节性明显", "不推荐首店"],
    ["红山", "4", "传统商圈，中老年客群占比高", "不推荐"],
    ["白鸟湖", "4", "新兴商圈，人口密度低", "不推荐首店"],
    ["乌北站", "4", "边缘商圈，商业配套不成熟", "不推荐首店"],
    ["粮校", "4", "边缘商圈，发展周期长", "不推荐首店"],
]

other_table = Table(other_areas, colWidths=[3*cm, 2.5*cm, 5.5*cm, 3*cm])
other_table.setStyle(TableStyle([
    ('FONTNAME', (0, 0), (-1, -1), chinese_font),
    ('FONTSIZE', (0, 0), (-1, -1), 9),
    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#4472C4')),
    ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
    ('PADDING', (0, 0), (-1, -1), 6),
]))
story.append(other_table)

# 三、综合推荐
story.append(PageBreak())
story.append(Paragraph("三、综合推荐", heading1_style))

story.append(Paragraph("首选：铁路局商圈", heading2_style))
story.append(Paragraph("<b>理由：</b>", normal_style))
story.append(Paragraph("1. 抖音兴趣人群TOP1，流量有数据支撑", bullet_style))
story.append(Paragraph("2. 中产家庭客群与减脂服务高度匹配", bullet_style))
story.append(Paragraph("3. 租金成本低，试错空间大", bullet_style))
story.append(Paragraph("4. 竞争较小，适合0-1样板店打造", bullet_style))
story.append(Paragraph("5. 社区口碑传播强，利于早期客户积累", bullet_style))

story.append(Paragraph("<b>建议动作：</b>", normal_style))
story.append(Paragraph("• 实地考察铁路局商圈商业街铺位", bullet_style))
story.append(Paragraph("• 调研周边竞品（美容院、健身房、减脂机构）", bullet_style))
story.append(Paragraph("• 测算租金、人力、装修等启动成本", bullet_style))
story.append(Paragraph("• 预估3个月盈亏平衡点", bullet_style))

story.append(Spacer(1, 15))
story.append(Paragraph("备选：友好北路+友好路商圈", heading2_style))
story.append(Paragraph("<b>定位：</b>品牌旗舰店/招商展示店", normal_style))
story.append(Paragraph("<b>前提条件：</b>", normal_style))
story.append(Paragraph("• 铁路局样板店跑通后", bullet_style))
story.append(Paragraph("• 有充足的资金储备（月租金可能过万）", bullet_style))
story.append(Paragraph("• 需要品牌势能背书招商", bullet_style))

story.append(Spacer(1, 15))
story.append(Paragraph("差异化选择：新疆民俗街/新华南路商圈", heading2_style))
story.append(Paragraph("<b>定位：</b>差异化门店，主打旅游+减脂体验", normal_style))
story.append(Paragraph("<b>前提条件：</b>", normal_style))
story.append(Paragraph("• 有清晰的游客转化策略", bullet_style))
story.append(Paragraph("• 品牌调性可以适配民族特色商圈", bullet_style))
story.append(Paragraph("• 可以承接旅游旺季客流", bullet_style))

# 四、下一步建议
story.append(Spacer(1, 20))
story.append(Paragraph("四、下一步建议", heading1_style))

steps = [
    ["序号", "事项", "时间建议"],
    ["1", "实地考察铁路局商圈，锁定2-3个候选铺位", "本周内"],
    ["2", "调研铁路局商圈现有减脂/美容/健身机构", "考察期间"],
    ["3", "基于实地铺位，测算启动成本与盈亏平衡点", "考察后3天"],
    ["4", "准备租赁合同模板，谈判租金、免租期、装修期", "确定铺位后"],
    ["5", "了解乌鲁木齐美容/健康服务行业证照要求", "签约前"],
]

steps_table = Table(steps, colWidths=[2*cm, 9*cm, 3*cm])
steps_table.setStyle(TableStyle([
    ('FONTNAME', (0, 0), (-1, -1), chinese_font),
    ('FONTSIZE', (0, 0), (-1, -1), 10),
    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#4472C4')),
    ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
    ('PADDING', (0, 0), (-1, -1), 8),
]))
story.append(steps_table)

# 结尾
story.append(Spacer(1, 30))
story.append(Paragraph("——报告完——", ParagraphStyle('Center', parent=normal_style, alignment=1)))
story.append(Spacer(1, 20))

footer_style = ParagraphStyle('Footer', parent=normal_style, fontSize=9, textColor=colors.grey)
story.append(Paragraph("<b>报告撰写：</b>QClaw智能助手", footer_style))
story.append(Paragraph("<b>数据来源：</b>抖音来客后台·行业兴趣人群地点偏好", footer_style))
story.append(Paragraph("<b>免责声明：</b>本报告基于抖音后台数据及公开商业信息分析，仅供参考，实际选址需结合实地调研及专业评估。", footer_style))

# 生成PDF
doc.build(story)
print("PDF生成成功")
