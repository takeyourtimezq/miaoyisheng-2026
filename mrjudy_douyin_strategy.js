const pptxgen = require("pptxgenjs");

// 创建演示文稿
let pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';
pres.author = '连锁帝国CGO';
pres.title = 'Mr.JUDY 抖音全案运营方案';
pres.subject = '参考TCJF云连锁打法，定制抖音线上运营策略';

// 蓝色主色调定义
const COLORS = {
  primary: "1E3A5F",      // 深蓝 - 主色
  secondary: "2E5A8C",    // 中蓝 - 辅助
  accent: "4A90D9",       // 亮蓝 - 强调
  light: "E8F4FC",        // 浅蓝 - 背景
  white: "FFFFFF",
  dark: "1A1A2E",         // 深色文字
  gray: "64748B",         // 灰色文字
  orange: "F59E0B"        // 橙色点缀
};

// 定义幻灯片母版
pres.defineSlideMaster({
  title: 'MASTER_SLIDE',
  background: { color: COLORS.white },
  objects: []
});

// ========== 第1页：封面 ==========
let slide1 = pres.addSlide();
slide1.background = { color: COLORS.primary };

// 装饰线条
slide1.addShape(pres.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.15,
  fill: { color: COLORS.accent }
});

// 主标题
slide1.addText("Mr.JUDY", {
  x: 0.5, y: 1.8, w: 9, h: 1,
  fontSize: 54, fontFace: "Arial Black", color: COLORS.white,
  bold: true, align: "center"
});

// 副标题
slide1.addText("抖音全案运营方案", {
  x: 0.5, y: 2.8, w: 9, h: 0.8,
  fontSize: 36, fontFace: "Arial", color: COLORS.accent,
  bold: true, align: "center"
});

// 说明文字
slide1.addText("参考TCJF云连锁打法，定制洗发养发品牌抖音线上运营策略", {
  x: 1, y: 4, w: 8, h: 0.6,
  fontSize: 16, fontFace: "Arial", color: COLORS.white,
  align: "center"
});

// 底部装饰
slide1.addShape(pres.shapes.RECTANGLE, {
  x: 0, y: 5.4, w: 10, h: 0.225,
  fill: { color: COLORS.accent }
});

// 日期
slide1.addText("2026年5月", {
  x: 0.5, y: 4.8, w: 9, h: 0.4,
  fontSize: 14, fontFace: "Arial", color: COLORS.white,
  align: "center"
});

// ========== 第2页：目录 ==========
let slide2 = pres.addSlide();

// 标题栏
slide2.addShape(pres.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 1.2,
  fill: { color: COLORS.primary }
});

slide2.addText("方案目录", {
  x: 0.5, y: 0.35, w: 9, h: 0.6,
  fontSize: 32, fontFace: "Arial Black", color: COLORS.white,
  bold: true
});

// 目录项
const tocItems = [
  "01  TCJF云连锁打法拆解",
  "02  Mr.JUDY品牌现状与差异分析",
  "03  整体策略定位",
  "04  内容策略规划",
  "05  团购套餐设计",
  "06  达人推广体系",
  "07  商家矩阵账号体系",
  "08  联动打法与阶段目标"
];

let tocY = 1.6;
tocItems.forEach((item, index) => {
  // 编号背景
  slide2.addShape(pres.shapes.RECTANGLE, {
    x: 0.8, y: tocY, w: 0.5, h: 0.45,
    fill: { color: COLORS.accent }
  });
  
  // 编号
  slide2.addText(String(index + 1).padStart(2, '0'), {
    x: 0.8, y: tocY, w: 0.5, h: 0.45,
    fontSize: 18, fontFace: "Arial", color: COLORS.white,
    bold: true, align: "center", valign: "middle"
  });
  
  // 文字
  slide2.addText(item.substring(4), {
    x: 1.5, y: tocY, w: 7, h: 0.45,
    fontSize: 18, fontFace: "Arial", color: COLORS.dark,
    valign: "middle"
  });
  
  tocY += 0.55;
});

// ========== 第3页：TCJF打法拆解 ==========
let slide3 = pres.addSlide();

// 标题栏
slide3.addShape(pres.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 1,
  fill: { color: COLORS.primary }
});

slide3.addText("01  TCJF云连锁打法拆解", {
  x: 0.5, y: 0.3, w: 9, h: 0.5,
  fontSize: 28, fontFace: "Arial Black", color: COLORS.white,
  bold: true
});

// 核心数据卡片
slide3.addShape(pres.shapes.RECTANGLE, {
  x: 0.5, y: 1.3, w: 2.8, h: 1.2,
  fill: { color: COLORS.light },
  shadow: { type: "outer", color: "000000", blur: 4, offset: 2, angle: 135, opacity: 0.1 }
});

slide3.addText("2700+", {
  x: 0.5, y: 1.45, w: 2.8, h: 0.5,
  fontSize: 36, fontFace: "Arial Black", color: COLORS.accent,
  bold: true, align: "center"
});

slide3.addText("全国门店规模", {
  x: 0.5, y: 1.9, w: 2.8, h: 0.4,
  fontSize: 14, fontFace: "Arial", color: COLORS.gray,
  align: "center"
});

// 第二个卡片
slide3.addShape(pres.shapes.RECTANGLE, {
  x: 3.6, y: 1.3, w: 2.8, h: 1.2,
  fill: { color: COLORS.light },
  shadow: { type: "outer", color: "000000", blur: 4, offset: 2, angle: 135, opacity: 0.1 }
});

slide3.addText("29.9元起", {
  x: 3.6, y: 1.45, w: 2.8, h: 0.5,
  fontSize: 32, fontFace: "Arial Black", color: COLORS.accent,
  bold: true, align: "center"
});

slide3.addText("低门槛引流价", {
  x: 3.6, y: 1.9, w: 2.8, h: 0.4,
  fontSize: 14, fontFace: "Arial", color: COLORS.gray,
  align: "center"
});

// 第三个卡片
slide3.addShape(pres.shapes.RECTANGLE, {
  x: 6.7, y: 1.3, w: 2.8, h: 1.2,
  fill: { color: COLORS.light },
  shadow: { type: "outer", color: "000000", blur: 4, offset: 2, angle: 135, opacity: 0.1 }
});

slide3.addText("15天", {
  x: 6.7, y: 1.45, w: 2.8, h: 0.5,
  fontSize: 36, fontFace: "Arial Black", color: COLORS.accent,
  bold: true, align: "center"
});

slide3.addText("不满意免费重做", {
  x: 6.7, y: 1.9, w: 2.8, h: 0.4,
  fontSize: 14, fontFace: "Arial", color: COLORS.gray,
  align: "center"
});

// 关键成功要素
slide3.addText("关键成功要素", {
  x: 0.5, y: 2.8, w: 9, h: 0.5,
  fontSize: 20, fontFace: "Arial", color: COLORS.primary,
  bold: true
});

const tcjfPoints = [
  "标准化服务流程 — 可复制的服务体验",
  "强售后承诺 — 降低用户试错成本",
  "云连锁模式 — 轻资产快速扩张，共享品牌流量",
  "阶梯式团购 — 引流款获客、利润款赚钱、形象款立调性"
];

let tcjfY = 3.3;
tcjfPoints.forEach((point) => {
  slide3.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: tcjfY, w: 0.12, h: 0.12,
    fill: { color: COLORS.accent }
  });
  
  slide3.addText(point, {
    x: 0.8, y: tcjfY - 0.05, w: 8.5, h: 0.4,
    fontSize: 14, fontFace: "Arial", color: COLORS.dark
  });
  
  tcjfY += 0.5;
});

// ========== 第4页：Mr.JUDY品牌分析 ==========
let slide4 = pres.addSlide();

// 标题栏
slide4.addShape(pres.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 1,
  fill: { color: COLORS.primary }
});

slide4.addText("02  Mr.JUDY品牌现状与差异分析", {
  x: 0.5, y: 0.3, w: 9, h: 0.5,
  fontSize: 28, fontFace: "Arial Black", color: COLORS.white,
  bold: true
});

// 品牌基础信息
slide4.addShape(pres.shapes.RECTANGLE, {
  x: 0.5, y: 1.3, w: 4.2, h: 3.8,
  fill: { color: COLORS.light }
});

slide4.addText("品牌基础", {
  x: 0.7, y: 1.5, w: 3.8, h: 0.4,
  fontSize: 18, fontFace: "Arial", color: COLORS.primary,
  bold: true
});

const brandInfo = [
  "定位：一线/新一线城市白领女性",
  "核心服务：洗发造型、头皮护理、按摩护理",
  "差异化：沉浸式独立服务空间",
  "战略合作：卡诗官方合作",
  "专利设备：洗护按摩椅专利",
  "行业地位：养发馆十大品牌之一"
];

let brandY = 2.0;
brandInfo.forEach((info) => {
  slide4.addText("• " + info, {
    x: 0.7, y: brandY, w: 3.8, h: 0.45,
    fontSize: 13, fontFace: "Arial", color: COLORS.dark
  });
  brandY += 0.5;
});

// 差异对比表
slide4.addShape(pres.shapes.RECTANGLE, {
  x: 5.0, y: 1.3, w: 4.5, h: 3.8,
  fill: { color: COLORS.white }
});

// 表头
slide4.addShape(pres.shapes.RECTANGLE, {
  x: 5.0, y: 1.3, w: 2.2, h: 0.5,
  fill: { color: COLORS.secondary }
});

slide4.addText("TCJF（理发）", {
  x: 5.0, y: 1.35, w: 2.2, h: 0.4,
  fontSize: 13, fontFace: "Arial", color: COLORS.white,
  bold: true, align: "center", valign: "middle"
});

slide4.addShape(pres.shapes.RECTANGLE, {
  x: 7.3, y: 1.3, w: 2.2, h: 0.5,
  fill: { color: COLORS.accent }
});

slide4.addText("Mr.JUDY（养发）", {
  x: 7.3, y: 1.35, w: 2.2, h: 0.4,
  fontSize: 13, fontFace: "Arial", color: COLORS.white,
  bold: true, align: "center", valign: "middle"
});

// 对比内容
const comparisons = [
  ["客单价", "中低（29.9-388）", "中高（80-300）"],
  ["消费频次", "高频（1-2月/次）", "中频（2-4周/次）"],
  ["决策门槛", "低（剪发刚需）", "中（养发非刚需）"],
  ["核心痛点", "发型效果", "头皮健康+放松体验"],
  ["目标人群", "大众消费者", "白领女性、品质体验"]
];

let compY = 1.9;
comparisons.forEach((row, idx) => {
  const bgColor = idx % 2 === 0 ? COLORS.light : COLORS.white;
  
  slide4.addShape(pres.shapes.RECTANGLE, {
    x: 5.0, y: compY, w: 4.5, h: 0.6,
    fill: { color: bgColor }
  });
  
  slide4.addText(row[0], {
    x: 5.1, y: compY + 0.15, w: 1.2, h: 0.3,
    fontSize: 12, fontFace: "Arial", color: COLORS.primary,
    bold: true
  });
  
  slide4.addText(row[1], {
    x: 6.3, y: compY + 0.15, w: 1.8, h: 0.3,
    fontSize: 11, fontFace: "Arial", color: COLORS.gray
  });
  
  slide4.addText(row[2], {
    x: 7.4, y: compY + 0.15, w: 1.9, h: 0.3,
    fontSize: 11, fontFace: "Arial", color: COLORS.accent,
    bold: true
  });
  
  compY += 0.6;
});

// ========== 第5页：整体策略定位 ==========
let slide5 = pres.addSlide();

// 标题栏
slide5.addShape(pres.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 1,
  fill: { color: COLORS.primary }
});

slide5.addText("03  整体策略定位", {
  x: 0.5, y: 0.3, w: 9, h: 0.5,
  fontSize: 28, fontFace: "Arial Black", color: COLORS.white,
  bold: true
});

// 核心定位大标题
slide5.addShape(pres.shapes.RECTANGLE, {
  x: 0.5, y: 1.4, w: 9, h: 1.2,
  fill: { color: COLORS.accent }
});

slide5.addText("「精致护理专家」定位 + 「体验驱动」转化", {
  x: 0.5, y: 1.7, w: 9, h: 0.6,
  fontSize: 28, fontFace: "Arial Black", color: COLORS.white,
  bold: true, align: "center"
});

// 策略支柱
slide5.addText("三大策略支柱", {
  x: 0.5, y: 3.0, w: 9, h: 0.5,
  fontSize: 20, fontFace: "Arial", color: COLORS.primary,
  bold: true
});

// 支柱1
slide5.addShape(pres.shapes.RECTANGLE, {
  x: 0.5, y: 3.6, w: 2.8, h: 1.6,
  fill: { color: COLORS.light }
});

slide5.addText("01", {
  x: 0.5, y: 3.7, w: 2.8, h: 0.5,
  fontSize: 32, fontFace: "Arial Black", color: COLORS.accent,
  bold: true, align: "center"
});

slide5.addText("体验差异化", {
  x: 0.5, y: 4.2, w: 2.8, h: 0.4,
  fontSize: 16, fontFace: "Arial", color: COLORS.primary,
  bold: true, align: "center"
});

slide5.addText("沉浸式独立空间\n卡诗品质背书", {
  x: 0.5, y: 4.6, w: 2.8, h: 0.5,
  fontSize: 12, fontFace: "Arial", color: COLORS.gray,
  align: "center"
});

// 支柱2
slide5.addShape(pres.shapes.RECTANGLE, {
  x: 3.6, y: 3.6, w: 2.8, h: 1.6,
  fill: { color: COLORS.light }
});

slide5.addText("02", {
  x: 3.6, y: 3.7, w: 2.8, h: 0.5,
  fontSize: 32, fontFace: "Arial Black", color: COLORS.accent,
  bold: true, align: "center"
});

slide5.addText("内容专业化", {
  x: 3.6, y: 4.2, w: 2.8, h: 0.4,
  fontSize: 16, fontFace: "Arial", color: COLORS.primary,
  bold: true, align: "center"
});

slide5.addText("头皮健康知识\n专业服务展示", {
  x: 3.6, y: 4.6, w: 2.8, h: 0.5,
  fontSize: 12, fontFace: "Arial", color: COLORS.gray,
  align: "center"
});

// 支柱3
slide5.addShape(pres.shapes.RECTANGLE, {
  x: 6.7, y: 3.6, w: 2.8, h: 1.6,
  fill: { color: COLORS.light }
});

slide5.addText("03", {
  x: 6.7, y: 3.7, w: 2.8, h: 0.5,
  fontSize: 32, fontFace: "Arial Black", color: COLORS.accent,
  bold: true, align: "center"
});

slide5.addText("转化阶梯化", {
  x: 6.7, y: 4.2, w: 2.8, h: 0.4,
  fontSize: 16, fontFace: "Arial", color: COLORS.primary,
  bold: true, align: "center"
});

slide5.addText("低价引流体验\n高价锁客复购", {
  x: 6.7, y: 4.6, w: 2.8, h: 0.5,
  fontSize: 12, fontFace: "Arial", color: COLORS.gray,
  align: "center"
});

// ========== 第6页：内容策略 ==========
let slide6 = pres.addSlide();

// 标题栏
slide6.addShape(pres.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 1,
  fill: { color: COLORS.primary }
});

slide6.addText("04  内容策略规划", {
  x: 0.5, y: 0.3, w: 9, h: 0.5,
  fontSize: 28, fontFace: "Arial Black", color: COLORS.white,
  bold: true
});

// 内容公式
slide6.addShape(pres.shapes.RECTANGLE, {
  x: 0.5, y: 1.3, w: 9, h: 0.8,
  fill: { color: COLORS.secondary }
});

slide6.addText("内容公式：痛点共鸣 + 场景代入 + 效果可视化", {
  x: 0.5, y: 1.5, w: 9, h: 0.4,
  fontSize: 20, fontFace: "Arial", color: COLORS.white,
  bold: true, align: "center"
});

// 内容类型占比
const contentTypes = [
  { type: "知识科普类", percent: "30%", desc: "头皮健康、脱发原因、养发误区、季节护理", color: COLORS.accent },
  { type: "服务展示类", percent: "35%", desc: "沉浸式洗头过程、头皮检测、护理前后对比", color: COLORS.secondary },
  { type: "用户证言类", percent: "20%", desc: "真实客户体验、Vlog记录、效果反馈", color: "059669" },
  { type: "品牌调性类", percent: "15%", desc: "卡诗合作、专利设备、独立空间、白领生活方式", color: "7C3AED" }
];

let contentY = 2.3;
contentTypes.forEach((item) => {
  // 占比数字
  slide6.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: contentY, w: 1.2, h: 0.5,
    fill: { color: item.color }
  });
  
  slide6.addText(item.percent, {
    x: 0.5, y: contentY + 0.08, w: 1.2, h: 0.35,
    fontSize: 18, fontFace: "Arial Black", color: COLORS.white,
    bold: true, align: "center"
  });
  
  // 类型名称
  slide6.addText(item.type, {
    x: 1.9, y: contentY + 0.08, w: 2, h: 0.35,
    fontSize: 16, fontFace: "Arial", color: COLORS.primary,
    bold: true
  });
  
  // 描述
  slide6.addText(item.desc, {
    x: 4.0, y: contentY + 0.1, w: 5.5, h: 0.35,
    fontSize: 13, fontFace: "Arial", color: COLORS.gray
  });
  
  contentY += 0.65;
});

// 爆款模板
slide6.addText("爆款内容模板", {
  x: 0.5, y: 5.0, w: 9, h: 0.4,
  fontSize: 16, fontFace: "Arial", color: COLORS.primary,
  bold: true
});

slide6.addText("3秒钩子 → 沉浸式体验 → 效果+行动", {
  x: 0.5, y: 5.4, w: 9, h: 0.3,
  fontSize: 14, fontFace: "Arial", color: COLORS.dark
});

// ========== 第7页：团购套餐设计 ==========
let slide7 = pres.addSlide();

// 标题栏
slide7.addShape(pres.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 1,
  fill: { color: COLORS.primary }
});

slide7.addText("05  团购套餐设计", {
  x: 0.5, y: 0.3, w: 9, h: 0.5,
  fontSize: 28, fontFace: "Arial Black", color: COLORS.white,
  bold: true
});

// 套餐卡片
const packages = [
  { name: "引流款", price: "39.9-59.9", content: "精致洗发+头皮检测+造型", purpose: "首客获取", color: "10B981" },
  { name: "利润款", price: "168-268", content: "深层护理+卡诗产品+按摩", purpose: "主力盈利", color: COLORS.accent },
  { name: "形象款", price: "398-598", content: "全套养护+定制方案+会员", purpose: "高端形象", color: COLORS.secondary },
  { name: "储值款", price: "1000-3000", content: "多次套餐+专属顾问+优先", purpose: "锁客复购", color: "7C3AED" }
];

let pkgX = 0.5;
packages.forEach((pkg) => {
  // 卡片背景
  slide7.addShape(pres.shapes.RECTANGLE, {
    x: pkgX, y: 1.4, w: 2.2, h: 3.2,
    fill: { color: COLORS.white },
    shadow: { type: "outer", color: "000000", blur: 6, offset: 2, angle: 135, opacity: 0.15 }
  });
  
  // 顶部色条
  slide7.addShape(pres.shapes.RECTANGLE, {
    x: pkgX, y: 1.4, w: 2.2, h: 0.15,
    fill: { color: pkg.color }
  });
  
  // 套餐名称
  slide7.addText(pkg.name, {
    x: pkgX, y: 1.7, w: 2.2, h: 0.4,
    fontSize: 18, fontFace: "Arial Black", color: pkg.color,
    bold: true, align: "center"
  });
  
  // 价格
  slide7.addText("¥" + pkg.price, {
    x: pkgX, y: 2.2, w: 2.2, h: 0.5,
    fontSize: 24, fontFace: "Arial Black", color: COLORS.dark,
    bold: true, align: "center"
  });
  
  // 内容
  slide7.addText(pkg.content, {
    x: pkgX + 0.1, y: 2.9, w: 2.0, h: 0.8,
    fontSize: 12, fontFace: "Arial", color: COLORS.gray,
    align: "center"
  });
  
  // 目的标签
  slide7.addShape(pres.shapes.RECTANGLE, {
    x: pkgX + 0.3, y: 4.0, w: 1.6, h: 0.4,
    fill: { color: pkg.color }
  });
  
  slide7.addText(pkg.purpose, {
    x: pkgX + 0.3, y: 4.05, w: 1.6, h: 0.3,
    fontSize: 12, fontFace: "Arial", color: COLORS.white,
    bold: true, align: "center", valign: "middle"
  });
  
  pkgX += 2.4;
});

// 转化路径
slide7.addShape(pres.shapes.RECTANGLE, {
  x: 0.5, y: 4.9, w: 9, h: 0.6,
  fill: { color: COLORS.light }
});

slide7.addText("转化路径：低价体验 → 到店升级 → 储值锁客", {
  x: 0.5, y: 5.0, w: 9, h: 0.4,
  fontSize: 16, fontFace: "Arial", color: COLORS.primary,
  bold: true, align: "center"
});

// ========== 第8页：达人推广体系 ==========
let slide8 = pres.addSlide();

// 标题栏
slide8.addShape(pres.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 1,
  fill: { color: COLORS.primary }
});

slide8.addText("06  达人推广体系", {
  x: 0.5, y: 0.3, w: 9, h: 0.5,
  fontSize: 28, fontFace: "Arial Black", color: COLORS.white,
  bold: true
});

// 达人生态分层表头
slide8.addShape(pres.shapes.RECTANGLE, {
  x: 0.5, y: 1.3, w: 9, h: 0.5,
  fill: { color: COLORS.secondary }
});

const tableHeaders = ["层级", "粉丝量", "合作数量", "核心目的"];
let headerX = 0.5;
const colWidths = [1.5, 1.8, 2.0, 3.7];

tableHeaders.forEach((header, idx) => {
  slide8.addText(header, {
    x: headerX, y: 1.35, w: colWidths[idx], h: 0.4,
    fontSize: 14, fontFace: "Arial", color: COLORS.white,
    bold: true, align: "center", valign: "middle"
  });
  headerX += colWidths[idx];
});

// 表格内容
const influencerTiers = [
  ["头部标杆", "100万+", "2-3位/季度", "品牌背书、破圈曝光"],
  ["腰部种草", "10-100万", "10-15位/月", "精准种草、信任建立"],
  ["本地生活", "1-10万", "30-50位/月", "到店转化、GMV拉动"],
  ["素人体验官", "0-1万", "100+位/月", "口碑沉淀、真实感营造"]
];

let tierY = 1.85;
influencerTiers.forEach((tier, idx) => {
  const bgColor = idx % 2 === 0 ? COLORS.light : COLORS.white;
  
  slide8.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: tierY, w: 9, h: 0.55,
    fill: { color: bgColor }
  });
  
  let cellX = 0.5;
  tier.forEach((cell, cellIdx) => {
    slide8.addText(cell, {
      x: cellX + 0.1, y: tierY + 0.15, w: colWidths[cellIdx] - 0.2, h: 0.3,
      fontSize: 12, fontFace: "Arial", color: cellIdx === 0 ? COLORS.primary : COLORS.dark,
      bold: cellIdx === 0, align: cellIdx === 0 ? "left" : "center"
    });
    cellX += colWidths[cellIdx];
  });
  
  tierY += 0.6;
});

// 达人类型
slide8.addText("达人类型与内容角度", {
  x: 0.5, y: 4.5, w: 9, h: 0.4,
  fontSize: 16, fontFace: "Arial", color: COLORS.primary,
  bold: true
});

const influencerTypes = [
  "美妆护肤博主 — 头皮护理是护肤的延伸",
  "职场/生活方式 — 白领下班放松场景",
  "本地探店达人 — 同城好店推荐",
  "健康养生博主 — 头皮健康科普"
];

let typeY = 4.95;
influencerTypes.forEach((type) => {
  slide8.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: typeY, w: 0.1, h: 0.1,
    fill: { color: COLORS.accent }
  });
  
  slide8.addText(type, {
    x: 0.7, y: typeY - 0.02, w: 8.5, h: 0.3,
    fontSize: 12, fontFace: "Arial", color: COLORS.dark
  });
  
  typeY += 0.35;
});

// ========== 第9页：商家矩阵账号体系 ==========
let slide9 = pres.addSlide();

// 标题栏
slide9.addShape(pres.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 1,
  fill: { color: COLORS.primary }
});

slide9.addText("07  商家矩阵账号体系", {
  x: 0.5, y: 0.3, w: 9, h: 0.5,
  fontSize: 28, fontFace: "Arial Black", color: COLORS.white,
  bold: true
});

// 架构图
slide9.addShape(pres.shapes.RECTANGLE, {
  x: 3.5, y: 1.3, w: 3, h: 0.7,
  fill: { color: COLORS.secondary }
});

slide9.addText("品牌主号", {
  x: 3.5, y: 1.4, w: 3, h: 0.5,
  fontSize: 16, fontFace: "Arial", color: COLORS.white,
  bold: true, align: "center", valign: "middle"
});

slide9.addText("官方形象、活动发布", {
  x: 3.5, y: 1.9, w: 3, h: 0.3,
  fontSize: 11, fontFace: "Arial", color: COLORS.gray,
  align: "center"
});

// 连接线
slide9.addShape(pres.shapes.LINE, {
  x: 5, y: 2.0, w: 0, h: 0.4,
  line: { color: COLORS.accent, width: 2 }
});

// 两个分支
slide9.addShape(pres.shapes.RECTANGLE, {
  x: 0.8, y: 2.5, w: 3.5, h: 0.7,
  fill: { color: COLORS.accent }
});

slide9.addText("商家门店号（1店1号）", {
  x: 0.8, y: 2.6, w: 3.5, h: 0.5,
  fontSize: 16, fontFace: "Arial", color: COLORS.white,
  bold: true, align: "center", valign: "middle"
});

slide9.addShape(pres.shapes.RECTANGLE, {
  x: 5.7, y: 2.5, w: 3.5, h: 0.7,
  fill: { color: COLORS.accent }
});

slide9.addText("职人IP号（1店2-3人）", {
  x: 5.7, y: 2.6, w: 3.5, h: 0.5,
  fontSize: 16, fontFace: "Arial", color: COLORS.white,
  bold: true, align: "center", valign: "middle"
});

// 门店号详情
slide9.addShape(pres.shapes.RECTANGLE, {
  x: 0.5, y: 3.4, w: 4.3, h: 2.2,
  fill: { color: COLORS.light }
});

slide9.addText("商家门店号", {
  x: 0.7, y: 3.55, w: 4, h: 0.4,
  fontSize: 16, fontFace: "Arial", color: COLORS.primary,
  bold: true
});

const storeDetails = [
  "定位：同城引流主力",
  "频率：2-3条/天",
  "内容：门店环境、服务过程、客户反馈",
  "KPI：同城曝光、团购点击、到店核销"
];

let storeY = 4.0;
storeDetails.forEach((detail) => {
  slide9.addText("• " + detail, {
    x: 0.7, y: storeY, w: 3.9, h: 0.35,
    fontSize: 12, fontFace: "Arial", color: COLORS.dark
  });
  storeY += 0.4;
});

// 职人IP详情
slide9.addShape(pres.shapes.RECTANGLE, {
  x: 5.2, y: 3.4, w: 4.3, h: 2.2,
  fill: { color: COLORS.light }
});

slide9.addText("职人IP号", {
  x: 5.4, y: 3.55, w: 4, h: 0.4,
  fontSize: 16, fontFace: "Arial", color: COLORS.primary,
  bold: true
});

const ipDetails = [
  "定位：人格化连接",
  "频率：1-2条/天",
  "内容：养发知识、工作日常、客户故事",
  "KPI：粉丝粘性、咨询转化、私域沉淀"
];

let ipY = 4.0;
ipDetails.forEach((detail) => {
  slide9.addText("• " + detail, {
    x: 5.4, y: ipY, w: 3.9, h: 0.35,
    fontSize: 12, fontFace: "Arial", color: COLORS.dark
  });
  ipY += 0.4;
});

// ========== 第10页：阶段目标 ==========
let slide10 = pres.addSlide();

// 标题栏
slide10.addShape(pres.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 1,
  fill: { color: COLORS.primary }
});

slide10.addText("08  阶段目标与关键差异", {
  x: 0.5, y: 0.3, w: 9, h: 0.5,
  fontSize: 28, fontFace: "Arial Black", color: COLORS.white,
  bold: true
});

// 三阶段目标
const stages = [
  { name: "冷启动期", time: "1-2月", goals: "账号搭建、内容测试\n粉丝破万、团购上线", color: "10B981" },
  { name: "成长期", time: "3-6月", goals: "爆款内容产出\n达人合作铺开\n月销破百万", color: COLORS.accent },
  { name: "规模化期", time: "6-12月", goals: "门店号矩阵成型\n职人IP孵化\n品牌声量提升", color: COLORS.secondary }
];

let stageX = 0.5;
stages.forEach((stage) => {
  // 卡片
  slide10.addShape(pres.shapes.RECTANGLE, {
    x: stageX, y: 1.4, w: 3, h: 2.8,
    fill: { color: COLORS.white },
    shadow: { type: "outer", color: "000000", blur: 6, offset: 2, angle: 135, opacity: 0.15 }
  });
  
  // 顶部色条
  slide10.addShape(pres.shapes.RECTANGLE, {
    x: stageX, y: 1.4, w: 3, h: 0.12,
    fill: { color: stage.color }
  });
  
  // 阶段名称
  slide10.addText(stage.name, {
    x: stageX, y: 1.65, w: 3, h: 0.4,
    fontSize: 20, fontFace: "Arial Black", color: stage.color,
    bold: true, align: "center"
  });
  
  // 时间
  slide10.addText(stage.time, {
    x: stageX, y: 2.1, w: 3, h: 0.3,
    fontSize: 14, fontFace: "Arial", color: COLORS.gray,
    align: "center"
  });
  
  // 目标
  slide10.addText(stage.goals, {
    x: stageX + 0.2, y: 2.5, w: 2.6, h: 1.5,
    fontSize: 13, fontFace: "Arial", color: COLORS.dark,
    align: "center"
  });
  
  stageX += 3.2;
});

// 与TCJF差异
slide10.addText("与TCJF的关键差异", {
  x: 0.5, y: 4.5, w: 9, h: 0.4,
  fontSize: 18, fontFace: "Arial", color: COLORS.primary,
  bold: true
});

const differences = [
  "TCJF：云连锁模式为主  →  Mr.JUDY：先单店模型跑通，再云连锁扩张",
  "TCJF：低价高频引流  →  Mr.JUDY：体验价值驱动，中高端定位",
  "TCJF：大众市场  →  Mr.JUDY：精准白领女性人群"
];

let diffY = 4.95;
differences.forEach((diff) => {
  slide10.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: diffY, w: 0.12, h: 0.12,
    fill: { color: COLORS.accent }
  });
  
  slide10.addText(diff, {
    x: 0.75, y: diffY - 0.02, w: 8.5, h: 0.3,
    fontSize: 12, fontFace: "Arial", color: COLORS.dark
  });
  
  diffY += 0.4;
});

// ========== 第11页：结束页 ==========
let slide11 = pres.addSlide();
slide11.background = { color: COLORS.primary };

// 装饰线条
slide11.addShape(pres.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.15,
  fill: { color: COLORS.accent }
});

// 感谢文字
slide11.addText("感谢观看", {
  x: 0.5, y: 2.2, w: 9, h: 1,
  fontSize: 48, fontFace: "Arial Black", color: COLORS.white,
  bold: true, align: "center"
});

// 副标题
slide11.addText("Mr.JUDY 抖音全案运营方案", {
  x: 0.5, y: 3.3, w: 9, h: 0.6,
  fontSize: 24, fontFace: "Arial", color: COLORS.accent,
  align: "center"
});

// 底部装饰
slide11.addShape(pres.shapes.RECTANGLE, {
  x: 0, y: 5.4, w: 10, h: 0.225,
  fill: { color: COLORS.accent }
});

// 保存文件
pres.writeFile({ fileName: "/Users/flyingspur/.qclaw/workspace-agent-6f07984c/MrJUDY_抖音全案运营方案.pptx" })
  .then(() => {
    console.log("PPT创建成功！");
  })
  .catch((err) => {
    console.error("创建失败:", err);
  });
