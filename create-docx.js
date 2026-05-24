const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, HeadingLevel, AlignmentType, BorderStyle, WidthType, ShadingType, PageBreak } = require('docx');
const fs = require('fs');

const border = { style: BorderStyle.SINGLE, size: 8, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };

const doc = new Document({
  creator: "升单爆破训练营",
  title: "升单爆破训练营招商方案",
  subject: "美业连锁品牌升单培训",
  styles: {
    default: { document: { run: { font: "Arial", size: 22 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 36, bold: true, font: "Arial", color: "1F4E79" },
        paragraph: { spacing: { before: 360, after: 240 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, font: "Arial", color: "2E75B6" },
        paragraph: { spacing: { before: 240, after: 180 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 24, bold: true, font: "Arial" },
        paragraph: { spacing: { before: 180, after: 120 }, outlineLevel: 2 } },
    ]
  },
  numbering: {
    config: [
      { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT }] },
      { reference: "numbers", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT }] },
    ]
  },
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
    children: [
      // 封面
      new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ size: 48, bold: true, color: "1F4E79", children: ["升单爆破训练营"] })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ size: 36, bold: true, children: ["—— 美业连锁品牌升单特训 ——"] })] }),
      new Paragraph({ children: [new TextRun("\n\n\n\n")] }),
      new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ size: 24, children: ["招商方案"] })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ size: 20, color: "666666", children: ["线上培训 + 1个月陪跑"] })] }),
      new Paragraph({ children: [new TextRun("\n\n\n\n\n\n\n\n\n\n\n\n")] }),
      new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ size: 18, color: "999999", children: ["适用对象：美业连锁品牌代理商/门店"] })] }),
      
      new Paragraph({ children: [new PageBreak()] }),
      
      // 目录
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("目录")] }),
      new Paragraph({ children: [new TextRun("一、产品定位与价值\n二、课程体系（2天培训）\n三、陪跑机制（1个月SOP）\n四、课程案例：39.9→99/200→1000-2000\n五、价格策略\n六、品牌方交付清单\n七、话术手册摘要\n八、我们的优势")] }),
      
      new Paragraph({ children: [new TextRun("\n\n")] }),
      
      // 一、产品定位与价值
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("一、产品定位与价值")] }),
      new Paragraph({ children: [new TextRun("\n")] }),
      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("1.1 产品定位")] }),
      new Paragraph({ children: [new TextRun("产品名称："), new TextRun({ bold: true, children: ["升单爆破训练营"] })] }),
      new Paragraph({ children: [new TextRun("目标客群："), new TextRun({ bold: true, children: ["美业连锁品牌代理商/门店店长/美容师"] })] }),
      new Paragraph({ children: [new TextRun("交付形式："), new TextRun({ bold: true, children: ["线上培训 + 1个月线上陪跑 / 线下培训 + 1个月线上陪跑"] })] }),
      
      new Paragraph({ children: [new TextRun("\n")] }),
      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("1.2 核心价值")] }),
      new Paragraph({ numbering: { reference: "numbers", level: 0 }, children: [new TextRun("从\"不会推\"到\"主动推\"：解决升单心理障碍")] }),
      new Paragraph({ numbering: { reference: "numbers", level: 0 }, children: [new TextRun("从\"不敢推\"到\"轻松推\"：掌握标准化话术武器库")] }),
      new Paragraph({ numbering: { reference: "numbers", level: 0 }, children: [new TextRun("从\"偶尔推\"到\"持续推\"：1个月陪跑养成销售习惯")] }),
      
      // 二、课程体系
      new Paragraph({ children: [new TextRun("\n\n")] }),
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("二、课程体系（2天培训）")] }),
      new Paragraph({ children: [new TextRun("\n")] }),
      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("2.1 课程架构")] }),
      new Paragraph({ children: [new TextRun("培训周期：2天（每天4-5小时）\n陪跑周期：1个月线上陪跑")] }),
      
      new Paragraph({ children: [new TextRun("\n")] }),
      new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun("Day 1：开营 + 心法（3-4小时）")] }),
      new Paragraph({ children: [new TextRun("• 0-30min：开营仪式+训练营规则（30分钟）")] }),
      new Paragraph({ children: [new TextRun("• 30-90min：升单底层逻辑：流量成本→必须升单（60分钟）")] }),
      new Paragraph({ children: [new TextRun("• 90-150min：抖音客户消费心理：占便宜/试探/比较（60分钟）")] }),
      new Paragraph({ children: [new TextRun("• 150-210min：接待SOP与破冰话术（60分钟）")] }),
      new Paragraph({ children: [new TextRun("• 210-240min：小作业说明：竞品暗访（课后完成，30分钟）")] }),
      
      new Paragraph({ children: [new TextRun("\n")] }),
      new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun("Day 2：话术 + 实战（4-5小时）")] }),
      new Paragraph({ children: [new TextRun("• 0-60min：需求挖掘：望闻问切+FABE话术（60分钟）")] }),
      new Paragraph({ children: [new TextRun("• 60-120min：价格异议处理：\"太贵了\"/\"考虑一下\"（60分钟）")] }),
      new Paragraph({ children: [new TextRun("• 120-180min：升单时机：3个黄金节点（60分钟）")] }),
      new Paragraph({ children: [new TextRun("• 180-240min：组合销售与套餐设计模板（60分钟）")] }),
      new Paragraph({ children: [new TextRun("• 240-300min：场景演练通关：全流程模拟（60分钟）")] }),
      new Paragraph({ children: [new TextRun("• 300-330min：陪跑启动+目标制定（30分钟）")] }),
      
      // 三、陪跑机制
      new Paragraph({ children: [new TextRun("\n\n")] }),
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("三、陪跑机制（1个月SOP）")] }),
      new Paragraph({ children: [new TextRun("\n")] }),
      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("3.1 陪跑目标")] }),
      new Paragraph({ numbering: { reference: "numbers", level: 0 }, children: [new TextRun("巩固课程所学：将2天培训内容转化为日常销售习惯")] }),
      new Paragraph({ numbering: { reference: "numbers", level: 0 }, children: [new TextRun("实战出结果：每学员每天至少3次升单尝试，目标升单率提升50%")] }),
      new Paragraph({ numbering: { reference: "numbers", level: 0 }, children: [new TextRun("问题实时响应：每日答疑，快速解决实战中遇到的问题")] }),
      
      new Paragraph({ children: [new TextRun("\n")] }),
      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("3.2 每日节奏")] }),
      new Paragraph({ children: [new TextRun("• 7:30 晨间分享：优秀案例/话术示范/激励语音（1分钟）")] }),
      new Paragraph({ children: [new TextRun("• 12:00 今日任务下达：当日升单目标+学习任务")] }),
      new Paragraph({ children: [new TextRun("• 18:00 数据上报：当天升单次数/成功/金额")] }),
      new Paragraph({ children: [new TextRun("• 21:00 晚间复盘+答疑：当天问题汇总解答/优秀案例分享")] }),
      new Paragraph({ children: [new TextRun("• 22:00 明日预告：明日课程/任务预告")] }),
      
      new Paragraph({ children: [new TextRun("\n")] }),
      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("3.3 参与机制（消费者行为学设计）")] }),
      new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun("解锁机制")] }),
      new Paragraph({ children: [new TextRun("当天任务完成→解锁次日课程回顾+新任务\n作业未通过→打回重做→解锁下一任务")] }),
      
      new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun("损失厌恶")] }),
      new Paragraph({ children: [new TextRun("连续3天未打卡→小组公示+班主任私聊\n连续7天未登录→移出陪跑群（可申请重入，需补交\"违约金\"）")] }),
      
      new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun("社交证明+荣誉体系")] }),
      new Paragraph({ children: [new TextRun("每日TOP3升单榜公示\n每周冠军小组奖励\n完成全部任务获得\"升单高手\"认证徽章")] }),
      
      new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun("小组PK制")] }),
      new Paragraph({ children: [new TextRun("5人一组，互相监督\n周冠军小组奖励：奖金/荣誉\n周垫底小组：表演节目/发红包")] }),
      
      new Paragraph({ children: [new TextRun("\n")] }),
      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("3.4 周计划")] }),
      new Paragraph({ children: [new TextRun("第1周 破冰+需求挖掘：每天3次破冰尝试，记录客户反馈")] }),
      new Paragraph({ children: [new TextRun("第2周 二转+异议处理：39.9→99/200二转尝试，价格异议处理练习")] }),
      new Paragraph({ children: [new TextRun("第3周 利润品升单：99/200→1000+升单尝试，套餐组合练习")] }),
      new Paragraph({ children: [new TextRun("第4周 综合实战+复盘：全流程实战，数据复盘，结业证书")] }),
      
      // 四、课程案例
      new Paragraph({ children: [new TextRun("\n\n")] }),
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("四、课程案例：39.9→99/200→1000-2000")] }),
      new Paragraph({ children: [new TextRun("\n")] }),
      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("4.1 升单路径")] }),
      new Paragraph({ children: [new TextRun({ bold: true, color: "C00000", children: ["抖音团购引流 39.9元 "] })] }),
      new Paragraph({ children: [new TextRun(" ↓")] }),
      new Paragraph({ children: [new TextRun({ bold: true, color: "E06000", children: ["到店体验 + 二转 99/200元 "] })] }),
      new Paragraph({ children: [new TextRun(" ↓")] }),
      new Paragraph({ children: [new TextRun({ bold: true, color: "00B050", children: ["利润品升单 1000-2000元 "] })] }),
      
      new Paragraph({ children: [new TextRun("\n")] }),
      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("4.2 全流程话术示例")] }),
      new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun("第一步：39.9到店破冰（0-5分钟）")] }),
      new Paragraph({ children: [new TextRun("\"您好！请坐～我是今天为您服务的美容师XX\"\n\"您是通过抖音过来的吗？第一次来我们店吧\"\n\"我们店主要做XX项目，您今天是想体验哪个呢？\"")] }),
      
      new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun("第二步：需求挖掘（5-15分钟）")] }),
      new Paragraph({ children: [new TextRun("\"您平时皮肤/头发有什么困扰吗？\"\n\"您希望改善到什么程度？\"\n\"我帮您分析一下，您这个问题是因为...\"")] }),
      
      new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun("第三步：99元二转（15-30分钟）")] }),
      new Paragraph({ children: [new TextRun("\"您今天体验的39.9项目是我们引流项目，非常划算\"\n\"我们店还有一个XX项目，性价比很高，99元3次\"\n\"这个项目正好能解决您刚才说的XX问题，我帮您预约一下？\"")] }),
      
      new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun("第四步：1000-2000利润品升单（30-60分钟）")] }),
      new Paragraph({ children: [new TextRun("\"您皮肤/头发的问题，需要搭配XX产品在家护理效果才持久\"\n\"我们店有XX套盒，效果很好，很多客户都在用\"\n\"今天正好有活动，买套盒送XX项目，全套1980元，比单买划算500元\"")] }),
      
      // 五、价格策略
      new Paragraph({ children: [new TextRun("\n\n")] }),
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("五、价格策略")] }),
      new Paragraph({ children: [new TextRun("\n")] }),
      new Paragraph({ children: [new TextRun("注：以下价格为对品牌方收费（非对门店），包含品牌旗下所有门店学员")] }),
      
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [3120, 3120, 3120],
        rows: [
          new TableRow({
            children: [
              new TableCell({ width: { size: 3120, type: WidthType.DXA }, shading: { fill: "D5E8F0", type: ShadingType.CLEAR }, children: [new Paragraph({ children: [new TextRun({ bold: true, children: ["规模"] })] })] }),
              new TableCell({ width: { size: 3120, type: WidthType.DXA }, shading: { fill: "D5E8F0", type: ShadingType.CLEAR }, children: [new Paragraph({ children: [new TextRun({ bold: true, children: ["线上版"] })] })] }),
              new TableCell({ width: { size: 3120, type: WidthType.DXA }, shading: { fill: "D5E8F0", type: ShadingType.CLEAR }, children: [new Paragraph({ children: [new TextRun({ bold: true, children: ["线下版"] })] })] }),
            ]
          }),
          new TableRow({
            children: [
              new TableCell({ width: { size: 3120, type: WidthType.DXA }, children: [new Paragraph("30家门店")] }),
              new TableCell({ width: { size: 3120, type: WidthType.DXA }, children: [new Paragraph({ children: [new TextRun({ bold: true, children: ["9.8万"] })] })] }),
              new TableCell({ width: { size: 3120, type: WidthType.DXA }, children: [new Paragraph({ children: [new TextRun({ bold: true, children: ["19.8万"] })] })] }),
            ]
          }),
          new TableRow({
            children: [
              new TableCell({ width: { size: 3120, type: WidthType.DXA }, children: [new Paragraph("50家门店")] }),
              new TableCell({ width: { size: 3120, type: WidthType.DXA }, children: [new Paragraph({ children: [new TextRun({ bold: true, children: ["15.8万"] })] })] }),
              new TableCell({ width: { size: 3120, type: WidthType.DXA }, children: [new Paragraph({ children: [new TextRun({ bold: true, children: ["32.8万"] })] })] }),
            ]
          }),
          new TableRow({
            children: [
              new TableCell({ width: { size: 3120, type: WidthType.DXA }, children: [new Paragraph("100家门店")] }),
              new TableCell({ width: { size: 3120, type: WidthType.DXA }, children: [new Paragraph({ children: [new TextRun({ bold: true, children: ["28.8万"] })] })] }),
              new TableCell({ width: { size: 3120, type: WidthType.DXA }, children: [new Paragraph({ children: [new TextRun({ bold: true, children: ["58.8万"] })] })] }),
            ]
          }),
          new TableRow({
            children: [
              new TableCell({ width: { size: 3120, type: WidthType.DXA }, children: [new Paragraph("200家+")] }),
              new TableCell({ width: { size: 3120, type: WidthType.DXA }, children: [new Paragraph("定制方案")] }),
              new TableCell({ width: { size: 3120, type: WidthType.DXA }, children: [new Paragraph("定制方案")] }),
            ]
          }),
        ]
      }),
      
      // 六、品牌方交付清单
      new Paragraph({ children: [new TextRun("\n\n")] }),
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("六、品牌方交付清单")] }),
      
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [3120, 3120, 3120],
        rows: [
          new TableRow({
            children: [
              new TableCell({ width: { size: 3120, type: WidthType.DXA }, shading: { fill: "D5E8F0", type: ShadingType.CLEAR }, children: [new Paragraph({ children: [new TextRun({ bold: true, children: ["交付内容"] })] })] }),
              new TableCell({ width: { size: 3120, type: WidthType.DXA }, shading: { fill: "D5E8F0", type: ShadingType.CLEAR }, children: [new Paragraph({ children: [new TextRun({ bold: true, children: ["线上版"] })] })] }),
              new TableCell({ width: { size: 3120, type: WidthType.DXA }, shading: { fill: "D5E8F0", type: ShadingType.CLEAR }, children: [new Paragraph({ children: [new TextRun({ bold: true, children: ["线下版"] })] })] }),
            ]
          }),
          new TableRow({
            children: [
              new TableCell({ width: { size: 3120, type: WidthType.DXA }, children: [new Paragraph("课程培训")] }),
              new TableCell({ width: { size: 3120, type: WidthType.DXA }, children: [new Paragraph("2天线上录播+直播")] }),
              new TableCell({ width: { size: 3120, type: WidthType.DXA }, children: [new Paragraph("1天线下集中授课")] }),
            ]
          }),
          new TableRow({
            children: [
              new TableCell({ width: { size: 3120, type: WidthType.DXA }, children: [new Paragraph("线上陪跑")] }),
              new TableCell({ width: { size: 3120, type: WidthType.DXA }, children: [new Paragraph("1个月（30天）")] }),
              new TableCell({ width: { size: 3120, type: WidthType.DXA }, children: [new Paragraph("1个月（30天）")] }),
            ]
          }),
          new TableRow({
            children: [
              new TableCell({ width: { size: 3120, type: WidthType.DXA }, children: [new Paragraph("学员账号")] }),
              new TableCell({ width: { size: 3120, type: WidthType.DXA }, children: [new Paragraph("30个")] }),
              new TableCell({ width: { size: 3120, type: WidthType.DXA }, children: [new Paragraph("30个")] }),
            ]
          }),
          new TableRow({
            children: [
              new TableCell({ width: { size: 3120, type: WidthType.DXA }, children: [new Paragraph("督导名额")] }),
              new TableCell({ width: { size: 3120, type: WidthType.DXA }, children: [new Paragraph("1名线上督导")] }),
              new TableCell({ width: { size: 3120, type: WidthType.DXA }, children: [new Paragraph("1名线下督导+1名线上督导")] }),
            ]
          }),
          new TableRow({
            children: [
              new TableCell({ width: { size: 3120, type: WidthType.DXA }, children: [new Paragraph("物料")] }),
              new TableCell({ width: { size: 3120, type: WidthType.DXA }, children: [new Paragraph("话术手册（电子）"), new Paragraph("工具包（电子）")] }),
              new TableCell({ width: { size: 3120, type: WidthType.DXA }, children: [new Paragraph("话术手册（纸质+电子）"), new Paragraph("工具包（纸质+电子）")] }),
            ]
          }),
          new TableRow({
            children: [
              new TableCell({ width: { size: 3120, type: WidthType.DXA }, children: [new Paragraph("专属社群")] }),
              new TableCell({ width: { size: 3120, type: WidthType.DXA }, children: [new Paragraph("品牌专属陪跑群")] }),
              new TableCell({ width: { size: 3120, type: WidthType.DXA }, children: [new Paragraph("品牌专属陪跑群")] }),
            ]
          }),
          new TableRow({
            children: [
              new TableCell({ width: { size: 3120, type: WidthType.DXA }, children: [new Paragraph("成果报告")] }),
              new TableCell({ width: { size: 3120, type: WidthType.DXA }, children: [new Paragraph("数据周报+结业报告")] }),
              new TableCell({ width: { size: 3120, type: WidthType.DXA }, children: [new Paragraph("门店诊断书+数据报告")] }),
            ]
          }),
        ]
      }),
      
      // 七、话术手册摘要
      new Paragraph({ children: [new TextRun("\n\n")] }),
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("七、话术手册摘要")] }),
      new Paragraph({ children: [new TextRun("\n")] }),
      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("7.1 破冰话术（3句必杀）")] }),
      new Paragraph({ children: [new TextRun("1. \"您好！请坐～我是今天为您服务的美容师XX\"\n2. \"您是通过抖音过来的吗？第一次来我们店吧\"\n3. \"我们店主要做XX项目，您今天是想体验哪个呢？\"")] }),
      
      new Paragraph({ children: [new TextRun("\n")] }),
      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("7.2 需求挖掘话术（望闻问切）")] }),
      new Paragraph({ children: [new TextRun("望：观察客户皮肤/头发状态\n闻：听客户描述问题\n问：您平时...？您希望...？\n切：给专业建议")] }),
      
      new Paragraph({ children: [new TextRun("\n")] }),
      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("7.3 FABE话术模板")] }),
      new Paragraph({ children: [new TextRun("F（特性）：这个产品是...\nA（优势）：它可以...\nB（利益）：对您来说...\nE（证据）：很多客户用了都...（案例）")] }),
      
      new Paragraph({ children: [new TextRun("\n")] }),
      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("7.4 价格异议处理（5种应对）")] }),
      new Paragraph({ numbering: { reference: "numbers", level: 0 }, children: [new TextRun("太贵了 → 分解法：算单次成本/对比线下/活动优惠")] }),
      new Paragraph({ numbering: { reference: "numbers", level: 0 }, children: [new TextRun("考虑一下 → 紧迫感：活动限时/名额有限/下次涨价")] }),
      new Paragraph({ numbering: { reference: "numbers", level: 0 }, children: [new TextRun("我考虑一下 → 追问法：您顾虑的是？")] }),
      new Paragraph({ numbering: { reference: "numbers", level: 0 }, children: [new TextRun("别人更便宜 → 价值法：对比服务/产品/效果")] }),
      new Paragraph({ numbering: { reference: "numbers", level: 0 }, children: [new TextRun("我要问老公 → 决策分离：先体验/先买先享受")] }),
      
      new Paragraph({ children: [new TextRun("\n")] }),
      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("7.5 升单时机（3个黄金节点）")] }),
      new Paragraph({ numbering: { reference: "numbers", level: 0 }, children: [new TextRun("第一次项目做完（立即推）")] }),
      new Paragraph({ numbering: { reference: "numbers", level: 0 }, children: [new TextRun("客户说\"不错\"（乘胜追击）")] }),
      new Paragraph({ numbering: { reference: "numbers", level: 0 }, children: [new TextRun("客户问\"多少钱\"（组合销售）")] }),
      
      new Paragraph({ children: [new TextRun("\n")] }),
      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("7.6 套餐设计模板")] }),
      new Paragraph({ children: [new TextRun("引流款：39.9元（抖音团单）→ 到店\n利润款：99元3次 → 绑定复购\n升级款：1000-2000元 → 盈利\n组合：买套盒送项目 = 价值感+绑定")] }),
      
      // 八、我们的优势
      new Paragraph({ children: [new TextRun("\n\n")] }),
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("八、我们的优势")] }),
      new Paragraph({ children: [new TextRun("\n")] }),
      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("1. 专注美业，更懂美业")] }),
      new Paragraph({ children: [new TextRun("深耕美业连锁抖音本地生活，案例库丰富，实战经验丰富")] }),
      
      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("2. 理论+实战+陪跑三重保障")] }),
      new Paragraph({ children: [new TextRun("不仅讲理论，更注重实战；不仅培训，更陪跑到出结果")] }),
      
      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("3. 消费者行为学设计参与机制"] }),
      new Paragraph({ children: [new TextRun("课程解锁、损失厌恶、社交证明、荣誉体系，让学员主动学、持续练")] }),
      
      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("4. 可量化的结果承诺"] }),
      new Paragraph({ children: [new TextRun("陪跑结束提供数据报告，升单率提升效果可衡量")] }),
      
      // 联系
      new Paragraph({ children: [new TextRun("\n\n\n\n")] }),
      new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ bold: true, size: 24, color: "1F4E79", children: ["联系我们"] })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ size: 18, children: ["联系人：铖烈"] })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ size: 18, children: ["电话：待补充"] })] }),
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("/Users/flyingspur/.qclaw/workspace-agent-6f07984c/升单爆破训练营招商方案.docx", buffer);
  console.log("Document created: 升单爆破训练营招商方案.docx");
});
