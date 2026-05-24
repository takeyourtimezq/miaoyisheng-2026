const pptxgen = require("pptxgenjs");

// 创建演示文稿
let pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';
pres.author = '连锁帝国CGO';
pres.title = 'Mr.JUDY 抖音全案运营方案';
pres.subject = '参考TCJF云连锁打法，定制洗发养发品牌抖音线上运营策略';

// 蓝色主色调定义
const COLORS = {
  primary: "1E3A5F",      // 深蓝 - 主色
  secondary: "2E5A8C",    // 中蓝 - 辅助
  accent: "4A90D9",       // 亮蓝 - 强调
  light: "E8F4FC",        // 浅蓝 - 背景
  white: "FFFFFF",
  dark: "1A1A2E",         // 深色文字
  gray: "64748B",         // 灰色文字
  orange: "F59E0B",       // 橙色点缀
  green: "10B981",        // 绿色
  purple: "7C3AED"        // 紫色
};

// ========== 第1页：封面 ==========
let slide1 = pres.addSlide();
slide1.background = { color: COLORS.primary };

slide1.addShape(pres.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.15,
  fill: { color: COLORS.accent }
});

slide1.addText("Mr.JUDY", {
  x: 0.5, y: 1.8, w: 9, h: 1,
  fontSize: 54, fontFace: "Arial Black", color: COLORS.white,
  bold: true, align: "center"
});

slide1.addText("抖音全案运营方案", {
  x: 0.5, y: 2.8, w: 9, h: 0.8,
  fontSize: 36, fontFace: "Arial", color: COLORS.accent,
  bold: true, align: "center"
});

slide1.addText("参考TCJF云连锁打法，定制洗发养发品牌抖音线上运营策略", {
  x: 1, y: 4, w: 8, h: 0.6,
  fontSize: 16, fontFace: "Arial", color: COLORS.white,
  align: "center"
});

slide1.addShape(pres.shapes.RECTANGLE, {
  x: 0, y: 5.4, w: 10, h: 0.225,
  fill: { color: COLORS.accent }
});

slide1.addText("2026年5月", {
  x: 0.5, y: 4.8, w: 9, h: 0.4,
  fontSize: 14, fontFace: "Arial", color: COLORS.white,
  align: "center"
});

// ========== 第2页：目录 ==========
let slide2 = pres.addSlide();

slide2.addShape(pres.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 1.2,
  fill: { color: COLORS.primary }
});

slide2.addText("方案目录", {
  x: 0.5, y: 0.35, w: 9, h: 0.6,
  fontSize: 32, fontFace: "Arial Black", color: COLORS.white,
  bold: true
});

const tocItems = [
  "01  TCJF云连锁打法深度拆解（3页）",
  "02  Mr.JUDY品牌分析与差异化定位（3页）",
  "03  整体策略与内容规划（3页）",
  "04  团购套餐与转化设计（3页）",
  "05  达人推广体系（3页）",
  "06  商家矩阵账号体系（3页）",
  "07  执行落地与阶段目标（3页）"
];

let tocY = 1.6;
tocItems.forEach((item, index) => {
  slide2.addShape(pres.shapes.RECTANGLE, {
    x: 0.8, y: tocY, w: 0.5, h: 0.45,
    fill: { color: COLORS.accent }
  });
  
  slide2.addText(String(index + 1).padStart(2, '0'), {
    x: 0.8, y: tocY, w: 0.5, h: 0.45,
    fontSize: 18, fontFace: "Arial", color: COLORS.white,
    bold: true, align: "center", valign: "middle"
  });
  
  slide2.addText(item.substring(4), {
    x: 1.5, y: tocY, w: 7, h: 0.45,
    fontSize: 16, fontFace: "Arial", color: COLORS.dark,
    valign: "middle"
  });
  
  tocY += 0.6;
});

// ==========================================
// 板块一：TCJF云连锁打法深度拆解（3页）
// ==========================================

// 第3页：TCJF模式概述与核心数据
let slide3 = pres.addSlide();

slide3.addShape(pres.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 1,
  fill: { color: COLORS.primary }
});

slide3.addText("01  TCJF云连锁打法深度拆解", {
  x: 0.5, y: 0.3, w: 9, h: 0.5,
  fontSize: 28, fontFace: "Arial Black", color: COLORS.white,
  bold: true
});

slide3.addText("（1/3）模式概述与核心数据", {
  x: 0.5, y: 0.75, w: 9, h: 0.3,
  fontSize: 14, fontFace: "Arial", color: COLORS.accent
});

// 核心数据卡片
const tcjfStats = [
  { num: "2700+", label: "全国门店规模", desc: "云连锁模式串联不同品牌/法人门店" },
  { num: "29.9", label: "引流款起价", desc: "洗吹造型三次卡低价获客" },
  { num: "388", label: "利润款均价", desc: "总监烫染套餐主力盈利" },
  { num: "15天", label: "售后承诺期", desc: "不满意免费重做降低决策门槛" }
];

let statX = 0.5;
tcjfStats.forEach((stat) => {
  slide3.addShape(pres.shapes.RECTANGLE, {
    x: statX, y: 1.4, w: 2.2, h: 1.8,
    fill: { color: COLORS.light },
    shadow: { type: "outer", color: "000000", blur: 4, offset: 2, angle: 135, opacity: 0.1 }
  });
  
  slide3.addText(stat.num, {
    x: statX, y: 1.55, w: 2.2, h: 0.6,
    fontSize: 32, fontFace: "Arial Black", color: COLORS.accent,
    bold: true, align: "center"
  });
  
  slide3.addText(stat.label, {
    x: statX, y: 2.2, w: 2.2, h: 0.35,
    fontSize: 14, fontFace: "Arial", color: COLORS.primary,
    bold: true, align: "center"
  });
  
  slide3.addText(stat.desc, {
    x: statX + 0.1, y: 2.6, w: 2.0, h: 0.5,
    fontSize: 11, fontFace: "Arial", color: COLORS.gray,
    align: "center"
  });
  
  statX += 2.4;
});

// 云连锁模式定义
slide3.addShape(pres.shapes.RECTANGLE, {
  x: 0.5, y: 3.5, w: 9, h: 1.8,
  fill: { color: COLORS.white },
  shadow: { type: "outer", color: "000000", blur: 4, offset: 2, angle: 135, opacity: 0.1 }
});

slide3.addText("云连锁模式定义", {
  x: 0.7, y: 3.65, w: 8.6, h: 0.4,
  fontSize: 18, fontFace: "Arial", color: COLORS.primary,
  bold: true
});

slide3.addText("由品牌方/服务商牵头，将线下同类目可标准化产品的门店认领到一起，实现不同品牌、不同营业执照、不同法人、不同地区的门店通过主品牌串联，共同享受品牌流量辐射，形成品牌矩阵，集中火力打造中央直播间，增加团购套餐核销量。", {
  x: 0.7, y: 4.1, w: 8.6, h: 1.0,
  fontSize: 13, fontFace: "Arial", color: COLORS.dark
});

// 第4页：TCJF内容策略与流量打法
let slide4 = pres.addSlide();

slide4.addShape(pres.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 1,
  fill: { color: COLORS.primary }
});

slide4.addText("01  TCJF云连锁打法深度拆解", {
  x: 0.5, y: 0.3, w: 9, h: 0.5,
  fontSize: 28, fontFace: "Arial Black", color: COLORS.white,
  bold: true
});

slide4.addText("（2/3）内容策略与流量打法", {
  x: 0.5, y: 0.75, w: 9, h: 0.3,
  fontSize: 14, fontFace: "Arial", color: COLORS.accent
});

// 内容策略
slide4.addShape(pres.shapes.RECTANGLE, {
  x: 0.5, y: 1.3, w: 4.3, h: 2.3,
  fill: { color: COLORS.light }
});

slide4.addText("内容策略", {
  x: 0.7, y: 1.45, w: 4, h: 0.4,
  fontSize: 18, fontFace: "Arial", color: COLORS.primary,
  bold: true
});

const tcjfContent = [
  "真人出镜：发型师/总监真人IP打造",
  "服务过程：烫染过程全程记录",
  "前后对比：发型改造前后强烈对比",
  "场景植入：五一出游、约会前变美等",
  "痛点切入：15天不满意免费重做"
];

let tcjfY = 1.9;
tcjfContent.forEach((item) => {
  slide4.addText("• " + item, {
    x: 0.7, y: tcjfY, w: 3.9, h: 0.35,
    fontSize: 12, fontFace: "Arial", color: COLORS.dark
  });
  tcjfY += 0.38;
});

// 流量来源
slide4.addShape(pres.shapes.RECTANGLE, {
  x: 5.2, y: 1.3, w: 4.3, h: 2.3,
  fill: { color: COLORS.light }
});

slide4.addText("流量来源矩阵", {
  x: 5.4, y: 1.45, w: 4, h: 0.4,
  fontSize: 18, fontFace: "Arial", color: COLORS.primary,
  bold: true
});

const trafficSources = [
  "短视频矩阵：主号+门店号+员工号",
  "达人探店：本地生活达人批量合作",
  "中央直播间：总部统一直播带货",
  "付费投流：DOU+精准投放",
  "自然流量：爆款内容二次传播"
];

let trafficY = 1.9;
trafficSources.forEach((item) => {
  slide4.addText("• " + item, {
    x: 5.4, y: trafficY, w: 3.9, h: 0.35,
    fontSize: 12, fontFace: "Arial", color: COLORS.dark
  });
  trafficY += 0.38;
});

// 爆款内容公式
slide4.addShape(pres.shapes.RECTANGLE, {
  x: 0.5, y: 3.8, w: 9, h: 1.5,
  fill: { color: COLORS.secondary }
});

slide4.addText("TCJF爆款内容公式", {
  x: 0.7, y: 3.9, w: 8.6, h: 0.4,
  fontSize: 16, fontFace: "Arial", color: COLORS.white,
  bold: true
});

slide4.addText("3秒钩子（痛点/好奇）+ 30秒过程（服务展示）+ 10秒转化（团购引导+售后承诺）", {
  x: 0.7, y: 4.35, w: 8.6, h: 0.4,
  fontSize: 14, fontFace: "Arial", color: COLORS.white,
  align: "center"
});

slide4.addText("案例：\"五一出游前变美急救\"话题，总监烫染套餐388元，#TCJF #烫染新主张美发不重样", {
  x: 0.7, y: 4.8, w: 8.6, h: 0.4,
  fontSize: 12, fontFace: "Arial", color: COLORS.light,
  align: "center"
});

// 第5页：TCJF团购策略与成功要素
let slide5 = pres.addSlide();

slide5.addShape(pres.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 1,
  fill: { color: COLORS.primary }
});

slide5.addText("01  TCJF云连锁打法深度拆解", {
  x: 0.5, y: 0.3, w: 9, h: 0.5,
  fontSize: 28, fontFace: "Arial Black", color: COLORS.white,
  bold: true
});

slide5.addText("（3/3）团购策略与关键成功要素", {
  x: 0.5, y: 0.75, w: 9, h: 0.3,
  fontSize: 14, fontFace: "Arial", color: COLORS.accent
});

// 阶梯式团购设计
slide5.addText("阶梯式团购套餐设计", {
  x: 0.5, y: 1.3, w: 9, h: 0.4,
  fontSize: 18, fontFace: "Arial", color: COLORS.primary,
  bold: true
});

const tcjfPackages = [
  { type: "引流款", price: "29.9-59元", content: "洗吹造型三次卡", purpose: "获客到店", color: COLORS.green },
  { type: "利润款", price: "168-388元", content: "烫染套餐/总监设计", purpose: "主力盈利", color: COLORS.accent },
  { type: "形象款", price: "588-888元", content: "高端定制/明星同款", purpose: "品牌调性", color: COLORS.secondary }
];

let pkgX = 0.5;
tcjfPackages.forEach((pkg) => {
  slide5.addShape(pres.shapes.RECTANGLE, {
    x: pkgX, y: 1.8, w: 3, h: 1.8,
    fill: { color: COLORS.white },
    shadow: { type: "outer", color: "000000", blur: 4, offset: 2, angle: 135, opacity: 0.1 }
  });
  
  slide5.addShape(pres.shapes.RECTANGLE, {
    x: pkgX, y: 1.8, w: 3, h: 0.12,
    fill: { color: pkg.color }
  });
  
  slide5.addText(pkg.type, {
    x: pkgX, y: 2.05, w: 3, h: 0.4,
    fontSize: 18, fontFace: "Arial Black", color: pkg.color,
    bold: true, align: "center"
  });
  
  slide5.addText(pkg.price, {
    x: pkgX, y: 2.5, w: 3, h: 0.4,
    fontSize: 22, fontFace: "Arial Black", color: COLORS.dark,
    bold: true, align: "center"
  });
  
  slide5.addText(pkg.content, {
    x: pkgX, y: 2.95, w: 3, h: 0.3,
    fontSize: 12, fontFace: "Arial", color: COLORS.gray,
    align: "center"
  });
  
  slide5.addShape(pres.shapes.RECTANGLE, {
    x: pkgX + 0.5, y: 3.35, w: 2, h: 0.35,
    fill: { color: pkg.color }
  });
  
  slide5.addText(pkg.purpose, {
    x: pkgX + 0.5, y: 3.38, w: 2, h: 0.3,
    fontSize: 11, fontFace: "Arial", color: COLORS.white,
    bold: true, align: "center", valign: "middle"
  });
  
  pkgX += 3.2;
});

// 关键成功要素
slide5.addText("TCJF关键成功要素", {
  x: 0.5, y: 3.9, w: 9, h: 0.4,
  fontSize: 18, fontFace: "Arial", color: COLORS.primary,
  bold: true
});

const successFactors = [
  { title: "标准化服务流程", desc: "从接待到售后的全流程SOP，确保每家门店服务一致性" },
  { title: "强售后承诺", desc: "15天不满意免费重做，大幅降低用户决策门槛和试错成本" },
  { title: "云连锁模式", desc: "轻资产快速扩张，不同主体门店共享品牌流量和中央运营" },
  { title: "阶梯式团购", desc: "引流款获客、利润款赚钱、形象款立调性，三层架构完整" }
];

let factorY = 4.35;
successFactors.forEach((factor) => {
  slide5.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: factorY, w: 0.12, h: 0.12,
    fill: { color: COLORS.accent }
  });
  
  slide5.addText(factor.title + "：" + factor.desc, {
    x: 0.75, y: factorY - 0.02, w: 8.5, h: 0.3,
    fontSize: 12, fontFace: "Arial", color: COLORS.dark
  });
  
  factorY += 0.35;
});

// ==========================================
// 板块二：Mr.JUDY品牌分析与差异化定位（3页）
// ==========================================

// 第6页：Mr.JUDY品牌基础与核心优势
let slide6 = pres.addSlide();

slide6.addShape(pres.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 1,
  fill: { color: COLORS.primary }
});

slide6.addText("02  Mr.JUDY品牌分析与差异化定位", {
  x: 0.5, y: 0.3, w: 9, h: 0.5,
  fontSize: 28, fontFace: "Arial Black", color: COLORS.white,
  bold: true
});

slide6.addText("（1/3）品牌基础与核心优势", {
  x: 0.5, y: 0.75, w: 9, h: 0.3,
  fontSize: 14, fontFace: "Arial", color: COLORS.accent
});

// 品牌基础信息
slide6.addShape(pres.shapes.RECTANGLE, {
  x: 0.5, y: 1.3, w: 4.3, h: 3.8,
  fill: { color: COLORS.light }
});

slide6.addText("品牌基础信息", {
  x: 0.7, y: 1.45, w: 4, h: 0.4,
  fontSize: 18, fontFace: "Arial", color: COLORS.primary,
  bold: true
});

const brandBase = [
  { label: "品牌名称", value: "Mr.JUDY 洗个头发" },
  { label: "所属公司", value: "北京小悦科技有限公司" },
  { label: "成立时间", value: "2018年" },
  { label: "目标人群", value: "一线/新一线城市白领女性" },
  { label: "核心服务", value: "洗发造型、头皮护理、按摩护理" },
  { label: "服务模式", value: "沉浸式独立服务空间" },
  { label: "行业地位", value: "养发馆十大品牌之一" }
];

let baseY = 1.95;
brandBase.forEach((item) => {
  slide6.addText(item.label + "：", {
    x: 0.7, y: baseY, w: 1.5, h: 0.35,
    fontSize: 12, fontFace: "Arial", color: COLORS.primary,
    bold: true
  });
  
  slide6.addText(item.value, {
    x: 2.2, y: baseY, w: 2.4, h: 0.35,
    fontSize: 12, fontFace: "Arial", color: COLORS.dark
  });
  
  baseY += 0.45;
});

// 核心差异化优势
slide6.addShape(pres.shapes.RECTANGLE, {
  x: 5.2, y: 1.3, w: 4.3, h: 3.8,
  fill: { color: COLORS.white },
  shadow: { type: "outer", color: "000000", blur: 4, offset: 2, angle: 135, opacity: 0.1 }
});

slide6.addText("核心差异化优势", {
  x: 5.4, y: 1.45, w: 4, h: 0.4,
  fontSize: 18, fontFace: "Arial", color: COLORS.primary,
  bold: true
});

const advantages = [
  { icon: "★", title: "卡诗官方战略合作", desc: "使用卡诗专业产品线，品质背书强" },
  { icon: "★", title: "专利洗护按摩椅", desc: "多功能洗护按摩椅专利，体验差异化" },
  { icon: "★", title: "独立私密空间", desc: "沉浸式独立服务空间，白领女性痛点精准击中" },
  { icon: "★", title: "标准化流程", desc: "从头皮检测到护理全流程标准化" },
  { icon: "★", title: "专业定位", desc: "垂直洗发养发，区别于综合理发店" }
];

let advY = 1.95;
advantages.forEach((adv) => {
  slide6.addText(adv.icon, {
    x: 5.4, y: advY, w: 0.3, h: 0.35,
    fontSize: 14, fontFace: "Arial", color: COLORS.accent,
    bold: true
  });
  
  slide6.addText(adv.title, {
    x: 5.7, y: advY, w: 3.6, h: 0.35,
    fontSize: 13, fontFace: "Arial", color: COLORS.dark,
    bold: true
  });
  
  slide6.addText(adv.desc, {
    x: 5.7, y: advY + 0.32, w: 3.6, h: 0.3,
    fontSize: 11, fontFace: "Arial", color: COLORS.gray
  });
  
  advY += 0.7;
});

// 第7页：与TCJF的深度对比分析
let slide7 = pres.addSlide();

slide7.addShape(pres.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 1,
  fill: { color: COLORS.primary }
});

slide7.addText("02  Mr.JUDY品牌分析与差异化定位", {
  x: 0.5, y: 0.3, w: 9, h: 0.5,
  fontSize: 28, fontFace: "Arial Black", color: COLORS.white,
  bold: true
});

slide7.addText("（2/3）与TCJF的深度对比分析", {
  x: 0.5, y: 0.75, w: 9, h: 0.3,
  fontSize: 14, fontFace: "Arial", color: COLORS.accent
});

// 对比表格
const comparisons = [
  ["对比维度", "TCJF（理发）", "Mr.JUDY（养发）", "策略启示"],
  ["客单价", "中低（29.9-388元）", "中高（80-300元）", "不拼低价，拼体验价值"],
  ["消费频次", "高频（1-2月/次）", "中频（2-4周/次）", "更高复购潜力，需强化会员"],
  ["决策门槛", "低（剪发刚需）", "中（养发非刚需）", "需教育市场，强调头皮健康"],
  ["核心痛点", "发型效果、价格透明", "头皮健康、放松体验、私密", "主打放松+健康双重价值"],
  ["目标人群", "大众消费者", "白领女性、品质体验", "精准人群，高LTV潜力"],
  ["服务场景", "公开空间、快节奏", "独立空间、慢享受", "差异化体验是核心壁垒"],
  ["竞争壁垒", "价格、便利性", "品质、体验、专业", "建立专业养发专家形象"]
];

// 表头
let compX = 0.5;
const compWidths = [1.4, 2.3, 2.3, 3.5];
comparisons[0].forEach((header, idx) => {
  slide7.addShape(pres.shapes.RECTANGLE, {
    x: compX, y: 1.3, w: compWidths[idx], h: 0.5,
    fill: { color: COLORS.secondary }
  });
  
  slide7.addText(header, {
    x: compX, y: 1.38, w: compWidths[idx], h: 0.4,
    fontSize: 12, fontFace: "Arial", color: COLORS.white,
    bold: true, align: "center", valign: "middle"
  });
  
  compX += compWidths[idx];
});

// 表格内容
let compY = 1.85;
comparisons.slice(1).forEach((row, idx) => {
  const bgColor = idx % 2 === 0 ? COLORS.light : COLORS.white;
  
  let cellX = 0.5;
  row.forEach((cell, cellIdx) => {
    slide7.addShape(pres.shapes.RECTANGLE, {
      x: cellX, y: compY, w: compWidths[cellIdx], h: 0.48,
      fill: { color: bgColor }
    });
    
    const textColor = cellIdx === 0 ? COLORS.primary : (cellIdx === 3 ? COLORS.accent : COLORS.dark);
    const bold = cellIdx === 0 || cellIdx === 3;
    
    slide7.addText(cell, {
      x: cellX + 0.08, y: compY + 0.12, w: compWidths[cellIdx] - 0.16, h: 0.3,
      fontSize: 10, fontFace: "Arial", color: textColor,
      bold: bold, align: cellIdx === 0 ? "left" : "center"
    });
    
    cellX += compWidths[cellIdx];
  });
  
  compY += 0.5;
});

// 第8页：差异化定位与核心策略
let slide8 = pres.addSlide();

slide8.addShape(pres.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 1,
  fill: { color: COLORS.primary }
});

slide8.addText("02  Mr.JUDY品牌分析与差异化定位", {
  x: 0.5, y: 0.3, w: 9, h: 0.5,
  fontSize: 28, fontFace: "Arial Black", color: COLORS.white,
  bold: true
});

slide8.addText("（3/3）差异化定位与核心策略", {
  x: 0.5, y: 0.75, w: 9, h: 0.3,
  fontSize: 14, fontFace: "Arial", color: COLORS.accent
});

// 核心定位
slide8.addShape(pres.shapes.RECTANGLE, {
  x: 0.5, y: 1.3, w: 9, h: 1.0,
  fill: { color: COLORS.accent }
});

slide8.addText("核心定位：精致护理专家", {
  x: 0.5, y: 1.55, w: 9, h: 0.5,
  fontSize: 32, fontFace: "Arial Black", color: COLORS.white,
  bold: true, align: "center"
});

// 三大策略支柱
slide8.addText("三大策略支柱", {
  x: 0.5, y: 2.5, w: 9, h: 0.4,
  fontSize: 18, fontFace: "Arial", color: COLORS.primary,
  bold: true
});

const pillars = [
  { num: "01", title: "体验差异化", subtitle: "沉浸式独立空间", desc: "卡诗品质背书，专利设备加持，打造白领女性专属放松空间", color: COLORS.accent },
  { num: "02", title: "内容专业化", subtitle: "头皮健康知识", desc: "从头皮检测 to 护理方案，建立专业养发专家形象", color: COLORS.secondary },
  { num: "03", title: "转化阶梯化", subtitle: "体验价值驱动", desc: "低价引流体验 → 到店升级转化 → 储值锁客复购", color: COLORS.green }
];

let pillarX = 0.5;
pillars.forEach((pillar) => {
  slide8.addShape(pres.shapes.RECTANGLE, {
    x: pillarX, y: 3.0, w: 3, h: 2.5,
    fill: { color: COLORS.light }
  });
  
  slide8.addShape(pres.shapes.RECTANGLE, {
    x: pillarX, y: 3.0, w: 3, h: 0.6,
    fill: { color: pillar.color }
  });
  
  slide8.addText(pillar.num, {
    x: pillarX, y: 3.1, w: 3, h: 0.4,
    fontSize: 24, fontFace: "Arial Black", color: COLORS.white,
    bold: true, align: "center"
  });
  
  slide8.addText(pillar.title, {
    x: pillarX, y: 3.7, w: 3, h: 0.4,
    fontSize: 18, fontFace: "Arial Black", color: pillar.color,
    bold: true, align: "center"
  });
  
  slide8.addText(pillar.subtitle, {
    x: pillarX, y: 4.1, w: 3, h: 0.3,
    fontSize: 13, fontFace: "Arial", color: COLORS.dark,
    align: "center"
  });
  
  slide8.addText(pillar.desc, {
    x: pillarX + 0.15, y: 4.5, w: 2.7, h: 0.8,
    fontSize: 11, fontFace: "Arial", color: COLORS.gray,
    align: "center"
  });
  
  pillarX += 3.2;
});

// ==========================================
// 板块三：整体策略与内容规划（3页）
// ==========================================

// 第9页：内容策略总览与配比
let slide9 = pres.addSlide();

slide9.addShape(pres.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 1,
  fill: { color: COLORS.primary }
});

slide9.addText("03  整体策略与内容规划", {
  x: 0.5, y: 0.3, w: 9, h: 0.5,
  fontSize: 28, fontFace: "Arial Black", color: COLORS.white,
  bold: true
});

slide9.addText("（1/3）内容策略总览与类型配比", {
  x: 0.5, y: 0.75, w: 9, h: 0.3,
  fontSize: 14, fontFace: "Arial", color: COLORS.accent
});

// 内容公式
slide9.addShape(pres.shapes.RECTANGLE, {
  x: 0.5, y: 1.3, w: 9, h: 0.7,
  fill: { color: COLORS.secondary }
});

slide9.addText("内容公式：痛点共鸣 + 场景代入 + 效果可视化", {
  x: 0.5, y: 1.5, w: 9, h: 0.4,
  fontSize: 20, fontFace: "Arial", color: COLORS.white,
  bold: true, align: "center"
});

// 内容类型详细拆解
const contentTypes = [
  { type: "知识科普类", percent: "30%", items: ["头皮健康科普", "脱发原因解析", "养发误区纠正", "季节护理指南"], color: COLORS.accent },
  { type: "服务展示类", percent: "35%", items: ["沉浸式洗头过程", "头皮检测展示", "护理前后对比", "ASMR音效体验"], color: COLORS.secondary },
  { type: "用户证言类", percent: "20%", items: ["真实客户体验", "Vlog记录", "效果反馈", "复购故事"], color: COLORS.green },
  { type: "品牌调性类", percent: "15%", items: ["卡诗合作展示", "专利设备介绍", "独立空间环境", "白领生活方式"], color: COLORS.purple }
];

let ctX = 0.5;
contentTypes.forEach((ct) => {
  slide9.addShape(pres.shapes.RECTANGLE, {
    x: ctX, y: 2.2, w: 2.2, h: 3.2,
    fill: { color: COLORS.white },
    shadow: { type: "outer", color: "000000", blur: 4, offset: 2, angle: 135, opacity: 0.1 }
  });
  
  slide9.addShape(pres.shapes.RECTANGLE, {
    x: ctX, y: 2.2, w: 2.2, h: 0.7,
    fill: { color: ct.color }
  });
  
  slide9.addText(ct.percent, {
    x: ctX, y: 2.35, w: 2.2, h: 0.4,
    fontSize: 28, fontFace: "Arial Black", color: COLORS.white,
    bold: true, align: "center"
  });
  
  slide9.addText(ct.type, {
    x: ctX, y: 3.0, w: 2.2, h: 0.35,
    fontSize: 14, fontFace: "Arial", color: ct.color,
    bold: true, align: "center"
  });
  
  let itemY = 3.45;
  ct.items.forEach((item) => {
    slide9.addText("• " + item, {
      x: ctX + 0.15, y: itemY, w: 1.9, h: 0.35,
      fontSize: 11, fontFace: "Arial", color: COLORS.dark
    });
    itemY += 0.38;
  });
  
  ctX += 2.4;
});

// 第10页：爆款内容模板与脚本结构
let slide10 = pres.addSlide();

slide10.addShape(pres.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 1,
  fill: { color: COLORS.primary }
});

slide10.addText("03  整体策略与内容规划", {
  x: 0.5, y: 0.3, w: 9, h: 0.5,
  fontSize: 28, fontFace: "Arial Black", color: COLORS.white,
  bold: true
});

slide10.addText("（2/3）爆款内容模板与脚本结构", {
  x: 0.5, y: 0.75, w: 9, h: 0.3,
  fontSize: 14, fontFace: "Arial", color: COLORS.accent
});

// 3秒钩子
slide10.addShape(pres.shapes.RECTANGLE, {
  x: 0.5, y: 1.3, w: 2.8, h: 3.8,
  fill: { color: COLORS.light }
});

slide10.addShape(pres.shapes.RECTANGLE, {
  x: 0.5, y: 1.3, w: 2.8, h: 0.5,
  fill: { color: COLORS.accent }
});

slide10.addText("3秒钩子", {
  x: 0.5, y: 1.38, w: 2.8, h: 0.4,
  fontSize: 18, fontFace: "Arial Black", color: COLORS.white,
  bold: true, align: "center", valign: "middle"
});

slide10.addText("目标：抓住注意力", {
  x: 0.5, y: 1.9, w: 2.8, h: 0.3,
  fontSize: 12, fontFace: "Arial", color: COLORS.gray,
  align: "center"
});

const hooks = [
  "\"你知道为什么头发越洗越油吗？\"",
  "\"头皮比脸皮老6岁，你护理了吗？\"",
  "\"白领女生下班后的秘密基地...\"",
  "\"熬夜掉发的姐妹看过来！\"",
  "\"29.9元体验贵妇级头皮护理\""
];

let hookY = 2.3;
hooks.forEach((hook) => {
  slide10.addText(hook, {
    x: 0.65, y: hookY, w: 2.5, h: 0.5,
    fontSize: 11, fontFace: "Arial", color: COLORS.dark
  });
  hookY += 0.55;
});

// 30秒过程
slide10.addShape(pres.shapes.RECTANGLE, {
  x: 3.6, y: 1.3, w: 2.8, h: 3.8,
  fill: { color: COLORS.light }
});

slide10.addShape(pres.shapes.RECTANGLE, {
  x: 3.6, y: 1.3, w: 2.8, h: 0.5,
  fill: { color: COLORS.secondary }
});

slide10.addText("30秒过程", {
  x: 3.6, y: 1.38, w: 2.8, h: 0.4,
  fontSize: 18, fontFace: "Arial Black", color: COLORS.white,
  bold: true, align: "center", valign: "middle"
});

slide10.addText("目标：建立信任", {
  x: 3.6, y: 1.9, w: 2.8, h: 0.3,
  fontSize: 12, fontFace: "Arial", color: COLORS.gray,
  align: "center"
});

const processes = [
  "头皮检测特写（专业感）",
  "ASMR洗头音效（沉浸感）",
  "按摩手法展示（技术感）",
  "客户放松表情（效果感）",
  "卡诗产品展示（品质感）"
];

let procY = 2.3;
processes.forEach((proc) => {
  slide10.addText("• " + proc, {
    x: 3.75, y: procY, w: 2.5, h: 0.5,
    fontSize: 11, fontFace: "Arial", color: COLORS.dark
  });
  procY += 0.55;
});

// 10秒转化
slide10.addShape(pres.shapes.RECTANGLE, {
  x: 6.7, y: 1.3, w: 2.8, h: 3.8,
  fill: { color: COLORS.light }
});

slide10.addShape(pres.shapes.RECTANGLE, {
  x: 6.7, y: 1.3, w: 2.8, h: 0.5,
  fill: { color: COLORS.green }
});

slide10.addText("10秒转化", {
  x: 6.7, y: 1.38, w: 2.8, h: 0.4,
  fontSize: 18, fontFace: "Arial Black", color: COLORS.white,
  bold: true, align: "center", valign: "middle"
});

slide10.addText("目标：引导行动", {
  x: 6.7, y: 1.9, w: 2.8, h: 0.3,
  fontSize: 12, fontFace: "Arial", color: COLORS.gray,
  align: "center"
});

const ctas = [
  "\"洗完轻了2斤的感觉\"",
  "\"点击左下角，新客只要39.9\"",
  "\"每天仅限10个名额\"",
  "\"限时福利，手慢无\"",
  "\"评论区扣1预约\""
];

let ctaY = 2.3;
ctas.forEach((cta) => {
  slide10.addText(cta, {
    x: 6.85, y: ctaY, w: 2.5, h: 0.5,
    fontSize: 11, fontFace: "Arial", color: COLORS.dark
  });
  ctaY += 0.55;
});

// 第11页：内容日历与发布节奏
let slide11 = pres.addSlide();

slide11.addShape(pres.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 1,
  fill: { color: COLORS.primary }
});

slide11.addText("03  整体策略与内容规划", {
  x: 0.5, y: 0.3, w: 9, h: 0.5,
  fontSize: 28, fontFace: "Arial Black", color: COLORS.white,
  bold: true
});

slide11.addText("（3/3）内容日历与发布节奏", {
  x: 0.5, y: 0.75, w: 9, h: 0.3,
  fontSize: 14, fontFace: "Arial", color: COLORS.accent
});

// 门店号内容日历
slide11.addText("商家门店号内容日历（2-3条/天）", {
  x: 0.5, y: 1.25, w: 4.5, h: 0.35,
  fontSize: 14, fontFace: "Arial", color: COLORS.primary,
  bold: true
});

const calendar = [
  ["周一", "门店环境", "服务快剪", "好评截图"],
  ["周二", "知识科普", "员工日常", "福利预告"],
  ["周三", "达人转发", "检测过程", "预约提醒"],
  ["周四", "产品展示", "ASMR洗头", "活动预告"],
  ["周五", "周末福利", "客户Vlog", "预约提醒"],
  ["周六", "高峰实拍", "排队场景", "当日战报"],
  ["周日", "周总结", "UGC合集", "互动问答"]
];

// 日历表头
let calX = 0.5;
["星期", "上午", "下午", "晚间"].forEach((header, idx) => {
  const width = idx === 0 ? 0.9 : 1.2;
  slide11.addShape(pres.shapes.RECTANGLE, {
    x: calX, y: 1.7, w: width, h: 0.4,
    fill: { color: COLORS.secondary }
  });
  
  slide11.addText(header, {
    x: calX, y: 1.75, w: width, h: 0.3,
    fontSize: 11, fontFace: "Arial", color: COLORS.white,
    bold: true, align: "center", valign: "middle"
  });
  
  calX += width;
});

// 日历内容
let calY = 2.15;
calendar.forEach((row, idx) => {
  const bgColor = idx % 2 === 0 ? COLORS.light : COLORS.white;
  
  let cellX = 0.5;
  row.forEach((cell, cellIdx) => {
    const width = cellIdx === 0 ? 0.9 : 1.2;
    
    slide11.addShape(pres.shapes.RECTANGLE, {
      x: cellX, y: calY, w: width, h: 0.38,
      fill: { color: bgColor }
    });
    
    slide11.addText(cell, {
      x: cellX, y: calY + 0.08, w: width, h: 0.25,
      fontSize: 9, fontFace: "Arial", color: cellIdx === 0 ? COLORS.primary : COLORS.dark,
      bold: cellIdx === 0, align: "center", valign: "middle"
    });
    
    cellX += width;
  });
  
  calY += 0.42;
});

// 必拍镜头清单
slide11.addText("必拍8个标准镜头", {
  x: 5.5, y: 1.25, w: 4, h: 0.35,
  fontSize: 14, fontFace: "Arial", color: COLORS.primary,
  bold: true
});

const shots = [
  "门头招牌（3秒）- 品牌露出",
  "接待区（5秒）- 环境展示",
  "独立护理间（5秒）- 私密空间",
  "头皮检测（10秒）- 专业感",
  "洗护过程（15秒）- ASMR",
  "放松表情（5秒）- 效果感",
  "前后对比（5秒）- 可视化",
  "团购引导（5秒）- 转化点"
];

let shotY = 1.7;
shots.forEach((shot, idx) => {
  slide11.addShape(pres.shapes.RECTANGLE, {
    x: 5.5, y: shotY, w: 0.25, h: 0.25,
    fill: { color: COLORS.accent }
  });
  
  slide11.addText(String(idx + 1), {
    x: 5.5, y: shotY, w: 0.25, h: 0.25,
    fontSize: 10, fontFace: "Arial", color: COLORS.white,
    bold: true, align: "center", valign: "middle"
  });
  
  slide11.addText(shot, {
    x: 5.85, y: shotY, w: 3.5, h: 0.3,
    fontSize: 11, fontFace: "Arial", color: COLORS.dark
  });
  
  shotY += 0.42;
});

// ==========================================
// 板块四：团购套餐与转化设计（3页）
// ==========================================

// 第12页：团购套餐详细设计
let slide12 = pres.addSlide();

slide12.addShape(pres.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 1,
  fill: { color: COLORS.primary }
});

slide12.addText("04  团购套餐与转化设计", {
  x: 0.5, y: 0.3, w: 9, h: 0.5,
  fontSize: 28, fontFace: "Arial Black", color: COLORS.white,
  bold: true
});

slide12.addText("（1/3）团购套餐详细设计", {
  x: 0.5, y: 0.75, w: 9, h: 0.3,
  fontSize: 14, fontFace: "Arial", color: COLORS.accent
});

// 四档套餐详细卡片
const mrjudyPackages = [
  { 
    name: "引流款", 
    price: "39.9-59.9", 
    content: "精致洗发+基础头皮检测+造型",
    value: "原价128元",
    target: "首客获取",
    kpis: "点击率、到店率",
    color: COLORS.green
  },
  { 
    name: "利润款", 
    price: "168-268", 
    content: "深层头皮护理+卡诗产品+按摩",
    value: "原价398元",
    target: "主力盈利",
    kpis: "转化率、客单价",
    color: COLORS.accent
  },
  { 
    name: "形象款", 
    price: "398-598", 
    content: "全套头皮养护+私人定制方案+会员权益",
    value: "原价888元",
    target: "高端形象",
    kpis: "品牌调性、口碑",
    color: COLORS.secondary
  },
  { 
    name: "储值款", 
    price: "1000-3000", 
    content: "多次护理套餐+专属顾问+优先预约",
    value: "相当于7折",
    target: "锁客复购",
    kpis: "LTV、复购率",
    color: COLORS.purple
  }
];

let mpX = 0.4;
mrjudyPackages.forEach((pkg) => {
  slide12.addShape(pres.shapes.RECTANGLE, {
    x: mpX, y: 1.3, w: 2.35, h: 3.8,
    fill: { color: COLORS.white },
    shadow: { type: "outer", color: "000000", blur: 4, offset: 2, angle: 135, opacity: 0.15 }
  });
  
  slide12.addShape(pres.shapes.RECTANGLE, {
    x: mpX, y: 1.3, w: 2.35, h: 0.15,
    fill: { color: pkg.color }
  });
  
  slide12.addText(pkg.name, {
    x: mpX, y: 1.55, w: 2.35, h: 0.4,
    fontSize: 16, fontFace: "Arial Black", color: pkg.color,
    bold: true, align: "center"
  });
  
  slide12.addText("¥" + pkg.price, {
    x: mpX, y: 2.0, w: 2.35, h: 0.5,
    fontSize: 24, fontFace: "Arial Black", color: COLORS.dark,
    bold: true, align: "center"
  });
  
  slide12.addText(pkg.value, {
    x: mpX, y: 2.55, w: 2.35, h: 0.25,
    fontSize: 10, fontFace: "Arial", color: COLORS.gray,
    align: "center"
  });
  
  slide12.addText(pkg.content, {
    x: mpX + 0.1, y: 2.9, w: 2.15, h: 0.7,
    fontSize: 11, fontFace: "Arial", color: COLORS.dark,
    align: "center"
  });
  
  slide12.addShape(pres.shapes.RECTANGLE, {
    x: mpX + 0.3, y: 3.7, w: 1.75, h: 0.35,
    fill: { color: pkg.color }
  });
  
  slide12.addText(pkg.target, {
    x: mpX + 0.3, y: 3.73, w: 1.75, h: 0.3,
    fontSize: 11, fontFace: "Arial", color: COLORS.white,
    bold: true, align: "center", valign: "middle"
  });
  
  slide12.addText("KPI：" + pkg.kpis, {
    x: mpX + 0.1, y: 4.15, w: 2.15, h: 0.3,
    fontSize: 10, fontFace: "Arial", color: COLORS.gray,
    align: "center"
  });
  
  mpX += 2.45;
});

// 由于内容太长，继续生成剩余页面...
// 保存当前进度

pres.writeFile({ fileName: "/Users/flyingspur/.qclaw/workspace-agent-6f07984c/MrJUDY_抖音全案运营方案_v2_part1.pptx" })
  .then(() => {
    console.log("PPT Part 1 创建成功！");
  })
  .catch((err) => {
    console.error("创建失败:", err);
  });
