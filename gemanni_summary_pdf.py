# -*- coding: utf-8 -*-
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4

font_path = '/System/Library/Fonts/STHeiti Light.ttc'
pdfmetrics.registerFont(TTFont('Hei', font_path, subfontIndex=0))

def create_pdf(output_path):
    c = canvas.Canvas(output_path, pagesize=A4)
    width, height = A4  # 595 x 842
    
    gold = '#B8860B'
    dark_gold = '#8B6914'
    black = '#1A1A1A'
    gray = '#666666'
    light_gray = '#999999'
    
    left = 40
    right = width - 40
    cw = right - left

    # Header
    y = height - 28
    c.setFont('Hei', 7)
    c.setFillColor(gray)
    c.drawRightString(right, y, "派特猫 × 歌曼丽 | 2026")
    y -= 15
    c.setFont('Hei', 15)
    c.setFillColor(black)
    c.drawCentredString(width/2, y, "歌曼丽抖音本地生活全案运营方案")
    y -= 11
    c.setFont('Hei', 9.5)
    c.setFillColor(gold)
    c.drawCentredString(width/2, y, "——让4100家门店成为你的流量护城河")
    y -= 6
    c.setStrokeColor(gold)
    c.setLineWidth(0.8)
    c.line(left, y, right, y)

    # === 一、三驾马车（全宽单列） ===
    y -= 16
    c.setFont('Hei', 10.5)
    c.setFillColor(dark_gold)
    c.drawString(left, y, "一、增长三驾马车 · 具体执行方法")
    y -= 13

    carriages = [
        ("【内容矩阵×达人生态】建立专业品牌认知", [
            "品牌总部号：品牌故事、创始人IP、技术实力展示、招商背书",
            "门店矩阵号：技师手法展示、蜕变案例分享、门店日常、本地活动引流",
            "职人IP号：技师个人IP、专业手法、蜕变故事、直播答疑",
            "达人合作：KOL品牌背书、KOC种草转化、素人达人UGC裂变",
            "自孵化达人：以培代招，6万+达人库，成本低于市场40-60%",
        ]),
        ("【话题营销×自造节】制造声量，引爆关注", [
            "话题定位：#草本减脂，把脂肪调出去 —— 反极端、反节食、反反弹",
            "四层传播：官号发起 → 门店扩散(4000+账号) → 职人互动(200+) → UGC裂变",
            "验证案例：#良咔桃花腰曝光3.2亿，#情动丽人塑造心动曝光3.2亿",
            "预期效果：目标曝光3亿+，参与账号4500+，参与用户100W+",
        ]),
        ("【精准投放×核销转化】流量变现，效果可验证", [
            "城市分层：高渗透地区降出价重转化，低渗透地区高出价抢新客",
            "团购分层：引流款99-199元 / 疗程款499-999元 / 管理款1999-3999元",
            "核销激励：到店核销送优惠、核销排行榜、邀请好友双奖励",
            "预期效果：核销率从35%提升至70%+，M2交易流量显著提升",
        ]),
    ]
    
    for title, items in carriages:
        c.setFont('Hei', 8.5)
        c.setFillColor(dark_gold)
        c.drawString(left, y, title)
        y -= 11
        c.setFont('Hei', 8)
        c.setFillColor(black)
        for item in items:
            c.drawString(left + 12, y, "• " + item)
            y -= 10
        y -= 4

    y -= 4

    # === 二、5/6/7月规划（全宽单列） ===
    c.setFont('Hei', 10.5)
    c.setFillColor(dark_gold)
    c.drawString(left, y, "二、5/6/7月三阶段增长规划")
    y -= 13

    months = [
        ("【5月·启动期】打基础、建体系", [
            "门店：筛选50家标杆门店，建立矩阵号，统一品牌视觉模板",
            "内容：建立脚本库、素材库，日均内容100+条，测试内容模型",
            "达人：接触洽谈KOL/KOC，确定首批合作达人名单",
            "话题：设计#草本减脂话题规则，官号发布预热内容",
            "重点：测试→复盘→优化，跑通单店模型再复制",
        ]),
        ("【6月·加速期】规模化、提声量", [
            "门店：激活150+门店，复制5月成功模型，批量培训门店运营",
            "内容：日均内容1000+条，话题#草本减脂热度爆发",
            "达人：KOL背书内容上线，KOC种草内容批量产出",
            "直播：6月底首场会销型直播，预期50W+观看，5000+成交",
            "重点：从测试转向规模化，数据驱动，快速迭代",
        ]),
        ("【7月·爆发期】造标杆、促招商", [
            "门店：激活250+门店，职人IP批量打造，技师变网红",
            "话题：#草本减脂曝光峰值，目标3亿+，参与用户100W+",
            "招商：标杆门店成功案例全网传播，招商线索转化加速",
            "沉淀：运营SOP沉淀，培训体系化，招商物料标准化",
            "重点：招商导向，线上声量→招商线索→成交转化",
        ]),
    ]
    
    for title, items in months:
        c.setFont('Hei', 8.5)
        c.setFillColor(dark_gold)
        c.drawString(left, y, title)
        y -= 11
        c.setFont('Hei', 8)
        c.setFillColor(black)
        for item in items:
            c.drawString(left + 12, y, "▸ " + item)
            y -= 10
        y -= 4

    y -= 4

    # === 三、线下招商会赋能 ===
    c.setFont('Hei', 10.5)
    c.setFillColor(dark_gold)
    c.drawString(left, y, "三、线下招商会赋能 · 一场会销，批量转化")
    y -= 13

    meeting = [
        "会前蓄客：线上话题预热+达人种草，锁定意向加盟商，预约到场率80%+",
        "现场造梦：标杆门店真实案例分享，GMV翻倍数据展示，技师IP变现故事",
        "政策释放：限时加盟优惠+首批扶持政策，现场签约享额外福利",
        "转化闭环：现场签约→次日打款→7天培训上岗，快速回本周期",
        "预期效果：单场到场200人，签约率30%+，单场成交60家+",
    ]
    c.setFont('Hei', 8)
    c.setFillColor(black)
    for item in meeting:
        c.drawString(left + 12, y, "• " + item)
        y -= 10

    y -= 8

    # === 四、六大引爆点 ===
    c.setFont('Hei', 10.5)
    c.setFillColor(dark_gold)
    c.drawString(left, y, "四、六大引爆点执行精要")
    y -= 13

    triggers = [
        ("1. 自营MCN生态闭环", "以培代招自孵化达人，覆盖全国各城市，成本降40-60%"),
        ("2. 标杆门店造梦", "先让1家店成功，打造神话，让所有门店羡慕并复制"),
        ("3. 自造节话题营销", "品牌自造节点，不靠平台施舍，预期曝光3亿+"),
        ("4. 本地推分层投放", "190+城市按渗透率分层，每一分钱都花在刀刃上"),
        ("5. 团购科学分层", "引流款→疗程款→管理款，层层转化，核销率70%+"),
        ("6. 会销型大场直播", "先锁客再开播，一场直播50W+观看，5000+成交"),
    ]
    
    for title, desc in triggers:
        c.setFont('Hei', 8.5)
        c.setFillColor(dark_gold)
        c.drawString(left + 12, y, title)
        c.setFont('Hei', 8)
        c.setFillColor(black)
        c.drawString(left + 130, y, desc)
        y -= 11

    y -= 6

    # === 五、十大心动理由 ===
    c.setFont('Hei', 10.5)
    c.setFillColor(dark_gold)
    c.drawString(left, y, "五、加盟歌曼丽 · 十大心动理由 —— 竞品给不了")
    y -= 13

    advantages = [
        ("1. 全国全渠道达人种草", "6万+自孵化达人矩阵，别人花大钱找达人，我们自带达人库"),
        ("2. 话题曝光3亿+", "自造节验证3.2亿曝光，门店在抖音被看见，加盟商自然找上门"),
        ("3. 标杆门店先成功", "不画饼，先让1家店成功给你看，真实GMV翻倍+收入数据"),
        ("4. 每周手把手培训", "从拍视频到开直播，从写脚本到做活动，小白也能变达人"),
        ("5. 城市分层精准投放", "4100+门店190+城市专属策略，不浪费一分钱广告费"),
        ("6. 核销率35%→70%", "到店核销才是真营收，科学分层+激励，客单价复购双提升"),
        ("7. 一场直播50万观看", "会销型大场直播，先锁客再开播，一场顶别人三个月"),
        ("8. 职人IP矩阵打造", "帮技师做个人IP，技师变网红，门店自带流量，招聘不发愁"),
        ("9. 草本减脂差异化", "中医智慧占位，竞品学不来，三四线城市天然信任感"),
        ("10. 派特猫实力背书", "2800+门店运营，6.35亿+总GMV，跟着靠谱的人做靠谱的事"),
    ]
    
    for title, desc in advantages:
        c.setFont('Hei', 8.5)
        c.setFillColor(dark_gold)
        c.drawString(left + 12, y, title)
        c.setFont('Hei', 8)
        c.setFillColor(black)
        c.drawString(left + 140, y, desc)
        y -= 11

    y -= 6

    # === 底部金句 ===
    c.setStrokeColor(gold)
    c.setLineWidth(0.5)
    c.line(left, y, right, y)
    y -= 14
    c.setFont('Hei', 11)
    c.setFillColor(dark_gold)
    c.drawCentredString(width/2, y, "线上声量 = 招商战力")
    y -= 12
    c.setFont('Hei', 8)
    c.setFillColor(gray)
    c.drawCentredString(width/2, y, "有线上声量：加盟决策1-2月  |  无线上声量：决策3-6月")
    y -= 10
    c.setFont('Hei', 7)
    c.setFillColor(light_gray)
    c.drawCentredString(width/2, y, "派特猫 · 让无限可能成为可能 | 2026年5月")
    
    c.save()

if __name__ == '__main__':
    output = '/Users/flyingspur/Downloads/歌曼丽全案价值摘要-招商优势.pdf'
    create_pdf(output)
    print(f'PDF已生成: {output}')
