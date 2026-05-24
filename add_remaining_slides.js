const pptxgen = require("pptxgenjs");

let pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';

const COLORS = {
  primary: "1E3A5F",
  secondary: "2E5A8C",
  accent: "4A90D9",
  light: "E8F4FC",
  white: "FFFFFF",
  dark: "1A1A2E",
  gray: "64748B",
  green: "10B981",
  purple: "7C3AED"
};

function addTitle(slide, t, st) {
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 1.1, fill: { color: COLORS.primary } });
  slide.addText(t, { x: 0.5, y: 0.25, w: 9, h: 0.5, fontSize: 26, fontFace: "Arial Black", color: COLORS.white, bold: true });
  if (st) slide.addText(st, { x: 0.5, y: 0.72, w: 9, h: 0.3, fontSize: 13, fontFace: "Arial", color: COLORS.accent });
}

// ========== 板块四：团购套餐（3页）==========

// 团购-1
let p12 = pres.addSlide();
addTitle(p12, "04  团购套餐与转化设计", "（1/3）团购套餐详细设计");
let px12 = 0.4;
[
  { n: "引流款", p: "39.9-59.9", v: "原价128元", c: "精致洗发+头皮检测+造型", t: "首客获取", k: "点击/到店率", cl: COLORS.green },
  { n: "利润款", p: "168-268", v: "原价398元", c: "深层护理+卡诗+按摩", t: "主力盈利", k: "转化/客单价", cl: COLORS.accent },
  { n: "形象款", p: "398-598", v: "原价888元", c: "全套养护+定制+会员", t: "高端形象", k: "品牌调性", cl: COLORS.secondary },
  { n: "储值款", p: "1000-3000", v: "相当于7折", c: "多次套餐+专属顾问", t: "锁客复购", k: "LTV/复购率", cl: COLORS.purple }
].forEach(pkg => {
  p12.addShape(pres.shapes.RECTANGLE, { x: px12, y: 1.3, w: 2.35, h: 3.8, fill: { color: COLORS.white } });
  p12.addShape(pres.shapes.RECTANGLE, { x: px12, y: 1.3, w: 2.35, h: 0.15, fill: { color: pkg.cl } });
  p12.addText(pkg.n, { x: px12, y: 1.55, w: 2.35, h: 0.4, fontSize: 16, fontFace: "Arial Black", color: pkg.cl, bold: true, align: "center" });
  p12.addText("¥" + pkg.p, { x: px12, y: 2.05, w: 2.35, h: 0.45, fontSize: 24, fontFace: "Arial Black", color: COLORS.dark, bold: true, align: "center" });
  p12.addText(pkg.v, { x: px12, y: 2.55, w: 2.35, h: 0.25, fontSize: 10, fontFace: "Arial", color: COLORS.gray, align: "center" });
  p12.addText(pkg.c, { x: px12 + 0.1, y: 2.9, w: 2.15, h: 0.7, fontSize: 11, fontFace: "Arial", color: COLORS.dark, align: "center" });
  p12.addShape(pres.shapes.RECTANGLE, { x: px12 + 0.3, y: 3.7, w: 1.75, h: 0.35, fill: { color: pkg.cl } });
  p12.addText(pkg.t, { x: px12 + 0.3, y: 3.73, w: 1.75, h: 0.3, fontSize: 11, fontFace: "Arial", color: COLORS.white, bold: true, align: "center", valign: "middle" });
  p12.addText("KPI：" + pkg.k, { x: px12 + 0.1, y: 4.15, w: 2.15, h: 0.3, fontSize: 10, fontFace: "Arial", color: COLORS.gray, align: "center" });
  px12 += 2.45;
});
p12.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.9, w: 9, h: 0.6, fill: { color: COLORS.light } });
p12.addText("转化路径：低价体验 → 到店升级 → 储值锁客", { x: 0.5, y: 5.0, w: 9, h: 0.4, fontSize: 16, fontFace: "Arial", color: COLORS.primary, bold: true, align: "center" });

// 团购-2
let p13 = pres.addSlide();
addTitle(p13, "04  团购套餐与转化设计", "（2/3）到店转化与升级策略");
p13.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.3, w: 4.3, h: 2.5, fill: { color: COLORS.light } });
p13.addText("到店体验流程", { x: 0.7, y: 1.45, w: 4, h: 0.4, fontSize: 18, fontFace: "Arial", color: COLORS.primary, bold: true });
[
  "1. 接待确认团购信息",
  "2. 头皮检测专业分析",
  "3. 服务过程体验展示",
  "4. 效果对比可视化",
  "5. 升级方案推荐",
  "6. 储值会员转化"
].forEach((s, i) => {
  p13.addShape(pres.shapes.RECTANGLE, { x: 0.7, y: 1.95 + i * 0.38, w: 0.2, h: 0.2, fill: { color: COLORS.accent } });
  p13.addText(s, { x: 1.0, y: 1.9 + i * 0.38, w: 3.6, h: 0.35, fontSize: 13, fontFace: "Arial", color: COLORS.dark });
});
p13.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.3, w: 4.3, h: 2.5, fill: { color: COLORS.light } });
p13.addText("升级话术要点", { x: 5.4, y: 1.45, w: 4, h: 0.4, fontSize: 18, fontFace: "Arial", color: COLORS.primary, bold: true });
[
  "\"您的头皮属于油性敏感...\"",
  "\"建议加做深层清洁护理...\"",
  "\"今天办理可享新客优惠...\"",
  "\"储值1000送200更划算...\"",
  "\"预约下次专属顾问服务...\""
].forEach((s, i) => {
  p13.addText("• " + s, { x: 5.4, y: 1.95 + i * 0.45, w: 3.9, h: 0.4, fontSize: 12, fontFace: "Arial", color: COLORS.dark });
});
p13.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.0, w: 9, h: 1.6, fill: { color: COLORS.white } });
p13.addText("私域沉淀路径", { x: 0.7, y: 4.15, w: 8.6, h: 0.4, fontSize: 18, fontFace: "Arial", color: COLORS.primary, bold: true });
p13.addText("抖音曝光 → 点击团购 → 到店体验 → 添加企微 → 会员体系 → 复购+转介绍", { x: 0.7, y: 4.6, w: 8.6, h: 0.4, fontSize: 14, fontFace: "Arial", color: COLORS.dark, align: "center" });
p13.addText("↑___________________________________________|", { x: 0.7, y: 5.0, w: 8.6, h: 0.3, fontSize: 12, fontFace: "Arial", color: COLORS.gray, align: "center" });
p13.addText("（社群运营、专属福利、预约提醒）", { x: 0.7, y: 5.3, w: 8.6, h: 0.3, fontSize: 11, fontFace: "Arial", color: COLORS.gray, align: "center" });

// 团购-3
let p14 = pres.addSlide();
addTitle(p14, "04  团购套餐与转化设计", "（3/3）与TCJF团购策略对比");
let cx14 = 0.5;
const cw14 = [1.5, 2.5, 2.5, 2.8];
["对比维度", "TCJF策略", "Mr.JUDY策略", "差异化价值"].forEach((h, i) => {
  p14.addShape(pres.shapes.RECTANGLE, { x: cx14, y: 1.3, w: cw14[i], h: 0.5, fill: { color: COLORS.secondary } });
  p14.addText(h, { x: cx14, y: 1.38, w: cw14[i], h: 0.4, fontSize: 12, fontFace: "Arial", color: COLORS.white, bold: true, align: "center", valign: "middle" });
  cx14 += cw14[i];
});
let cy14 = 1.85;
[
  ["引流款定价", "29.9元起", "39.9元起", "保持品质感不拼低价"],
  ["核心卖点", "15天不满意重做", "卡诗+独立空间", "体验价值差异化"],
  ["升级路径", "烫染套餐", "头皮护理套餐", "健康理念教育"],
  ["储值策略", "剪发卡", "护理会员", "高频复购锁定"],
  ["目标LTV", "500元/年", "2000元/年", "高净值人群价值"],
  ["售后承诺", "15天重做", "7天无理由", "降低决策门槛"]
].forEach((r, i) => {
  cx14 = 0.5;
  r.forEach((c, j) => {
    p14.addShape(pres.shapes.RECTANGLE, { x: cx14, y: cy14, w: cw14[j], h: 0.55, fill: { color: i % 2 === 0 ? COLORS.light : COLORS.white } });
    p14.addText(c, { x: cx14 + 0.1, y: cy14 + 0.15, w: cw14[j] - 0.2, h: 0.3, fontSize: 11, fontFace: "Arial", color: j === 0 ? COLORS.primary : (j === 3 ? COLORS.accent : COLORS.dark), bold: j === 0 || j === 3, align: j === 0 ? "left" : "center" });
    cx14 += cw14[j];
  });
  cy14 += 0.6;
});
p14.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.6, w: 9, h: 0.9, fill: { color: COLORS.accent } });
p14.addText("核心差异：TCJF拼低价+便利性，Mr.JUDY拼体验+专业度，目标人群LTV提升4倍", { x: 0.5, y: 4.85, w: 9, h: 0.4, fontSize: 14, fontFace: "Arial", color: COLORS.white, bold: true, align: "center" });

// ========== 板块五：达人推广（3页）==========

// 达人-1
let p15 = pres.addSlide();
addTitle(p15, "05  达人推广体系", "（1/3）达人生态分层策略");
let cx15 = 0.5;
const cw15 = [1.4, 1.4, 2.0, 2.6, 2.2];
["层级", "粉丝量", "合作数量", "合作方式", "核心目的"].forEach((h, i) => {
  p15.addShape(pres.shapes.RECTANGLE, { x: cx15, y: 1.3, w: cw15[i], h: 0.5, fill: { color: COLORS.secondary } });
  p15.addText(h, { x: cx15, y: 1.38, w: cw15[i], h: 0.4, fontSize: 12, fontFace: "Arial", color: COLORS.white, bold: true, align: "center", valign: "middle" });
  cx15 += cw15[i];
});
let cy15 = 1.85;
[
  ["头部标杆", "100万+", "2-3位/季", "品牌代言+深度合作", "品牌背书、破圈"],
  ["腰部种草", "10-100万", "10-15位/月", "探店视频+知识科普", "精准种草信任"],
  ["本地生活", "1-10万", "30-50位/月", "团购挂载+同城引流", "到店转化GMV"],
  ["素人体验官", "0-1万", "100+位/月", "免费体验换UGC", "口碑沉淀真实感"]
].forEach((r, i) => {
  cx15 = 0.5;
  r.forEach((c, j) => {
    p15.addShape(pres.shapes.RECTANGLE, { x: cx15, y: cy15, w: cw15[j], h: 0.65, fill: { color: i % 2 === 0 ? COLORS.light : COLORS.white } });
    p15.addText(c, { x: cx15 + 0.05, y: cy15 + 0.2, w: cw15[j] - 0.1, h: 0.3, fontSize: 11, fontFace: "Arial", color: j === 0 ? COLORS.primary : COLORS.dark, bold: j === 0, align: j === 0 ? "left" : "center" });
    cx15 += cw15[j];
  });
  cy15 += 0.7;
});
p15.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.7, w: 9, h: 1.0, fill: { color: COLORS.white } });
p15.addText("月度达人合作预算参考", { x: 0.7, y: 4.8, w: 8.6, h: 0.35, fontSize: 16, fontFace: "Arial", color: COLORS.primary, bold: true });
p15.addText("头部：5-10万/位  |  腰部：5000-2万/位  |  本地：500-2000/位  |  素人：免费置换", { x: 0.7, y: 5.2, w: 8.6, h: 0.35, fontSize: 13, fontFace: "Arial", color: COLORS.dark, align: "center" });

// 达人-2
let p16 = pres.addSlide();
addTitle(p16, "05  达人推广体系", "（2/3）达人类型与内容角度");
let px16 = 0.5;
[
  { t: "美妆护肤博主", a: "头皮护理是护肤延伸", s: "\"头皮比脸皮老6岁\"", c: COLORS.accent },
  { t: "职场生活方式", a: "白领下班放松场景", s: "\"打工人的秘密基地\"", c: COLORS.secondary },
  { t: "本地探店达人", a: "同城好店推荐", s: "\"XX城市最值得体验\"", c: COLORS.green },
  { t: "健康养生博主", a: "头皮健康科普", s: "\"脱发头油解决方案\"", c: COLORS.purple }
].forEach((d, i) => {
  p16.addShape(pres.shapes.RECTANGLE, { x: px16, y: 1.4, w: 2.2, h: 2.8, fill: { color: COLORS.white } });
  p16.addShape(pres.shapes.RECTANGLE, { x: px16, y: 1.4, w: 2.2, h: 0.5, fill: { color: d.c } });
  p16.addText(d.t, { x: px16, y: 1.48, w: 2.2, h: 0.4, fontSize: 14, fontFace: "Arial Black", color: COLORS.white, bold: true, align: "center", valign: "middle" });
  p16.addText("内容角度：", { x: px16 + 0.1, y: 2.0, w: 2.0, h: 0.3, fontSize: 11, fontFace: "Arial", color: COLORS.gray });
  p16.addText(d.a, { x: px16 + 0.1, y: 2.3, w: 2.0, h: 0.5, fontSize: 12, fontFace: "Arial", color: COLORS.dark, bold: true });
  p16.addText("核心话术：", { x: px16 + 0.1, y: 2.9, w: 2.0, h: 0.3, fontSize: 11, fontFace: "Arial", color: COLORS.gray });
  p16.addText(d.s, { x: px16 + 0.1, y: 3.2, w: 2.0, h: 0.8, fontSize: 12, fontFace: "Arial", color: d.c, bold: true });
  px16 += 2.4;
});
p16.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.4, w: 9, h: 1.3, fill: { color: COLORS.light } });
p16.addText("达人合作SOP", { x: 0.7, y: 4.5, w: 8.6, h: 0.35, fontSize: 16, fontFace: "Arial", color: COLORS.primary, bold: true });
p16.addText("筛选 → 触达 → 内容共创 → 投放追踪 → 效果复盘", { x: 0.7, y: 4.9, w: 8.6, h: 0.35, fontSize: 14, fontFace: "Arial", color: COLORS.dark, align: "center" });
p16.addText("关键指标：播放量、互动率、团购点击、到店核销、单达人ROI", { x: 0.7, y: 5.25, w: 8.6, h: 0.35, fontSize: 12, fontFace: "Arial", color: COLORS.gray, align: "center" });

// 达人-3
let p17 = pres.addSlide();
addTitle(p17, "05  达人推广体系", "（3/3）达人Brief与内容规范");
p17.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.3, w: 4.3, h: 3.8, fill: { color: COLORS.light } });
p17.addText("达人Brief必拍内容", { x: 0.7, y: 1.45, w: 4, h: 0.4, fontSize: 18, fontFace: "Arial", color: COLORS.primary, bold: true });
[
  "1. 门头/环境（突出私密空间）",
  "2. 头皮检测过程（专业感）",
  "3. 洗护服务过程（ASMR）",
  "4. 前后对比/反馈（效果）",
  "5. 卡诗产品展示（品质）",
  "6. 团购引导口播（转化）"
].forEach((s, i) => {
  p17.addText(s, { x: 0.7, y: 1.95 + i * 0.55, w: 3.9, h: 0.5, fontSize: 12, fontFace: "Arial", color: COLORS.dark });
});
p17.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.3, w: 4.3, h: 1.7, fill: { color: COLORS.white } });
p17.addText("必提卖点（至少2个）", { x: 5.4, y: 1.45, w: 4, h: 0.4, fontSize: 18, fontFace: "Arial", color: COLORS.primary, bold: true });
[
  "\"独立私密空间，安心享受\"",
  "\"卡诗官方合作，大牌产品\"",
  "\"专利按摩椅，躺着护理\"",
  "\"头皮检测，了解真实状态\""
].forEach((s, i) => {
  p17.addText("• " + s, { x: 5.4, y: 1.9 + i * 0.35, w: 3.9, h: 0.35, fontSize: 11, fontFace: "Arial", color: COLORS.dark });
});
p17.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 3.2, w: 4.3, h: 1.9, fill: { color: "FEE2E2" } });
p17.addText("内容禁忌", { x: 5.4, y: 3.35, w: 4, h: 0.4, fontSize: 18, fontFace: "Arial", color: "DC2626", bold: true });
[
  "❌ 不要提\"理发\"\"剪发\"",
  "❌ 不要过度承诺效果",
  "❌ 不要出现竞品对比",
  "❌ 不要泄露客户隐私"
].forEach((s, i) => {
  p17.addText(s, { x: 5.4, y: 3.8 + i * 0.38, w: 3.9, h: 0.35, fontSize: 12, fontFace: "Arial", color: "DC2626" });
});

// ========== 板块六：矩阵账号（3页）==========

// 矩阵-1
let p18 = pres.addSlide();
addTitle(p18, "06  商家矩阵账号体系", "（1/3）矩阵架构与分工");
p18.addShape(pres.shapes.RECTANGLE, { x: 3.5, y: 1.3, w: 3, h: 0.7, fill: { color: COLORS.secondary } });
p18.addText("品牌主号", { x: 3.5, y: 1.4, w: 3, h: 0.5, fontSize: 16, fontFace: "Arial", color: COLORS.white, bold: true, align: "center", valign: "middle" });
p18.addText("官方形象、活动发布", { x: 3.5, y: 2.05, w: 3, h: 0.25, fontSize: 11, fontFace: "Arial", color: COLORS.gray, align: "center" });
p18.addShape(pres.shapes.LINE, { x: 5, y: 2.35, w: 0, h: 0.3, line: { color: COLORS.accent, width: 2 } });

p18.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 2.7, w: 3.5, h: 0.7, fill: { color: COLORS.accent } });
p18.addText("商家门店号（1店1号）", { x: 0.8, y: 2.8, w: 3.5, h: 0.5, fontSize: 16, fontFace: "Arial", color: COLORS.white, bold: true, align: "center", valign: "middle" });
p18.addText("同城引流、团购转化、2-3条/天", { x: 0.8, y: 3.45, w: 3.5, h: 0.25, fontSize: 11, fontFace: "Arial", color: COLORS.gray, align: "center" });

p18.addShape(pres.shapes.RECTANGLE, { x: 5.7, y: 2.7, w: 3.5, h: 0.7, fill: { color: COLORS.accent } });
p18.addText("职人IP号（1店2-3人）", { x: 5.7, y: 2.8, w: 3.5, h: 0.5, fontSize: 16, fontFace: "Arial", color: COLORS.white, bold: true, align: "center", valign: "middle" });
p18.addText("信任建立、专业背书、1-2条/天", { x: 5.7, y: 3.45, w: 3.5, h: 0.25, fontSize: 11, fontFace: "Arial", color: COLORS.gray, align: "center" });

p18.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.9, w: 4.3, h: 1.8, fill: { color: COLORS.light } });
p18.addText("门店号核心职责", { x: 0.7, y: 4.0, w: 4, h: 0.35, fontSize: 16, fontFace: "Arial", color: COLORS.primary, bold: true });
["发布门店环境、服务展示", "挂载团购链接、活动预告", "转发达人内容、客户反馈", "KPI：到店核销数、团购GMV"].forEach((s, i) => {
  p18.addText("• " + s, { x: 0.7, y: 4.4 + i * 0.38, w: 3.9, h: 0.35, fontSize: 12, fontFace: "Arial", color: COLORS.dark });
});

p18.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 3.9, w: 4.3, h: 1.8, fill: { color: COLORS.light } });
p18.addText("职人IP核心职责", { x: 5.4, y: 4.0, w: 4, h: 0.35, fontSize: 16, fontFace: "Arial", color: COLORS.primary, bold: true });
["发布养发知识、专业科普", "分享工作日常、客户故事", "承接咨询、建立信任转私域", "KPI：咨询转化、私域沉淀"].forEach((s, i) => {
  p18.addText("• " + s, { x: 5.4, y: 4.4 + i * 0.38, w: 3.9, h: 0.35, fontSize: 12, fontFace: "Arial", color: COLORS.dark });
});

// 矩阵-2
let p19 = pres.addSlide();
addTitle(p19, "06  商家矩阵账号体系", "（2/3）职人IP打造指南");
p19.addText("职人IP人设定位方向", { x: 0.5, y: 1.25, w: 9, h: 0.4, fontSize: 18, fontFace: "Arial", color: COLORS.primary, bold: true });
let px19 = 0.5;
[
  { t: "专业顾问型", s: "资深养发师", c: "头皮诊断+方案定制", ex: "@头皮医生小王", cl: COLORS.accent },
  { t: "温暖陪伴型", s: "前台/客服", c: "客户故事+情感连接", ex: "@JUDY小助理", cl: COLORS.secondary },
  { t: "知识科普型", s: "培训师", c: "养发知识+产品测评", ex: "@养发研究所", cl: COLORS.green },
  { t: "颜值担当型", s: "形象好员工", c: "工作日常+生活方式", ex: "@JUDY洗头哥", cl: COLORS.purple }
].forEach(p => {
  p19.addShape(pres.shapes.RECTANGLE, { x: px19, y: 1.75, w: 2.2, h: 2.8, fill: { color: COLORS.white } });
  p19.addShape(pres.shapes.RECTANGLE, { x: px19, y: 1.75, w: 2.2, h: 0.5, fill: { color: p.cl } });
  p19.addText(p.t, { x: px19, y: 1.83, w: 2.2, h: 0.4, fontSize: 14, fontFace: "Arial Black", color: COLORS.white, bold: true, align: "center", valign: "middle" });
  p19.addText(p.s, { x: px19, y: 2.35, w: 2.2, h: 0.3, fontSize: 11, fontFace: "Arial", color: COLORS.gray, align: "center" });
  p19.addText(p.c, { x: px19 + 0.1, y: 2.75, w: 2.0, h: 0.6, fontSize: 12, fontFace: "Arial", color: COLORS.dark, align: "center" });
  p19.addText(p.ex, { x: px19, y: 3.5, w: 2.2, h: 0.3, fontSize: 11, fontFace: "Arial", color: p.cl, bold: true, align: "center" });
  px19 += 2.4;
});
p19.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.75, w: 9, h: 1.0, fill: { color: COLORS.light } });
p19.addText("职人IP选拔标准", { x: 0.7, y: 4.85, w: 8.6, h: 0.35, fontSize: 16, fontFace: "Arial", color: COLORS.primary, bold: true });
p19.addText("形象气质佳 | 表达能力好 | 专业技能扎实 | 愿意出镜 | 有社交活跃度优先", { x: 0.7, y: 5.25, w: 8.6, h: 0.35, fontSize: 13, fontFace: "Arial", color: COLORS.dark, align: "center" });

// 矩阵-3
let p20 = pres.addSlide();
addTitle(p20, "06  商家矩阵账号体系", "（3/3）协作机制与考核");
p20.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.3, w: 4.3, h: 2.6, fill: { color: COLORS.light } });
p20.addText("门店号与职人IP协作", { x: 0.7, y: 1.45, w: 4, h: 0.4, fontSize: 18, fontFace: "Arial", color: COLORS.primary, bold: true });
[
  "达人探店后：门店号转发挂载团购，职人IP知识承接",
  "新品上线：门店号官方发布，职人IP测评分享",
  "客户咨询：门店号引导预约，职人IP专业解答",
  "负面反馈：门店号官方回应，职人IP情感安抚"
].forEach((s, i) => {
  p20.addText("• " + s, { x: 0.7, y: 1.95 + i * 0.55, w: 3.9, h: 0.5, fontSize: 12, fontFace: "Arial", color: COLORS.dark });
});
p20.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.3, w: 4.3, h: 2.6, fill: { color: COLORS.light } });
p20.addText("考核指标", { x: 5.4, y: 1.45, w: 4, h: 0.4, fontSize: 18, fontFace: "Arial", color: COLORS.primary, bold: true });
p20.addText("商家门店号", { x: 5.4, y: 1.95, w: 4, h: 0.35, fontSize: 14, fontFace: "Arial", color: COLORS.accent, bold: true });
p20.addText("核心：到店核销数、团购GMV", { x: 5.4, y: 2.35, w: 4, h: 0.3, fontSize: 12, fontFace: "Arial", color: COLORS.dark });
p20.addText("次要：同城曝光、粉丝增长", { x: 5.4, y: 2.7, w: 4, h: 0.3, fontSize: 12, fontFace: "Arial", color: COLORS.gray });
p20.addText("职人IP号", { x: 5.4, y: 3.15, w: 4, h: 0.35, fontSize: 14, fontFace: "Arial", color: COLORS.accent, bold: true });
p20.addText("核心：咨询转化率、私域沉淀数", { x: 5.4, y: 3.55, w: 4, h: 0.3, fontSize: 12, fontFace: "Arial", color: COLORS.dark });
p20.addText("次要：粉丝量、互动率", { x: 5.4, y: 3.9, w: 4, h: 0.3, fontSize: 12, fontFace: "Arial", color: COLORS.gray });

p20.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.1, w: 9, h: 1.6, fill: { color: COLORS.white } });
p20.addText("组织架构", { x: 0.7, y: 4.2, w: 8.6, h: 0.35, fontSize: 18, fontFace: "Arial", color: COLORS.primary, bold: true });
p20.addText("总部运营中心", { x: 0.7, y: 4.65, w: 2.5, h: 0.3, fontSize: 14, fontFace: "Arial", color: COLORS.accent, bold: true });
p20.addText("品牌主号运营 + 达人合作组 + 区域运营经理", { x: 3.3, y: 4.65, w: 5.5, h: 0.3, fontSize: 12, fontFace: "Arial", color: COLORS.dark });
p20.addText("门店执行层", { x: 0.7, y: 5.05, w: 2.5, h: 0.3, fontSize: 14, fontFace: "Arial", color: COLORS.accent, bold: true });
p20.addText("门店号运营（员工兼职）+ 职人IP孵化（总部培训支持）", { x: 3.3, y: 5.05, w: 5.5, h: 0.3, fontSize: 12, fontFace: "Arial", color: COLORS.dark });

// ========== 板块七：执行落地（3页）==========

// 执行-1
let p21 = pres.addSlide();
addTitle(p21, "07  执行落地与阶段目标", "（1/3）三阶段目标规划");
let px21 = 0.5;
[
  { n: "01", t: "冷启动期", tm: "1-2月", g: ["账号搭建完成", "内容测试跑通", "粉丝破万", "团购上线"], cl: COLORS.green },
  { n: "02", t: "成长期", tm: "3-6月", g: ["爆款内容产出", "达人合作铺开", "月销破百万", "门店矩阵成型"], cl: COLORS.accent },
  { n: "03", t: "规模化期", tm: "6-12月", g: ["职人IP孵化", "品牌声量提升", "云连锁扩张", "行业标杆打造"], cl: COLORS.secondary }
].forEach(st => {
  p21.addShape(pres.shapes.RECTANGLE, { x: px21, y: 1.4, w: 3, h: 3.5, fill: { color: COLORS.white } });
  p21.addShape(pres.shapes.RECTANGLE, { x: px21, y: 1.4, w: 3, h: 0.8, fill: { color: st.cl } });
  p21.addText(st.n, { x: px21, y: 1.55, w: 3, h: 0.5, fontSize: 32, fontFace: "Arial Black", color: COLORS.white, bold: true, align: "center" });
  p21.addText(st.t, { x: px21, y: 2.3, w: 3, h: 0.4, fontSize: 20, fontFace: "Arial Black", color: st.cl, bold: true, align: "center" });
  p21.addText(st.tm, { x: px21, y: 2.75, w: 3, h: 0.3, fontSize: 13, fontFace: "Arial", color: COLORS.gray, align: "center" });
  st.g.forEach((gl, i) => {
    p21.addText("✓ " + gl, { x: px21 + 0.2, y: 3.15 + i * 0.4, w: 2.6, h: 0.35, fontSize: 12, fontFace: "Arial", color: COLORS.dark });
  });
  px21 += 3.2;
});
p21.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 5.05, w: 9, h: 0.7, fill: { color: COLORS.light } });
p21.addText("关键里程碑：2月团购上线 → 4月月销破百万 → 12月千店矩阵", { x: 0.5, y: 5.15, w: 9, h: 0.45, fontSize: 15, fontFace: "Arial", color: COLORS.primary, bold: true, align: "center" });

// 执行-2
let p22 = pres.addSlide();
addTitle(p22, "07  执行落地与阶段目标", "（2/3）月度执行节奏");
let cx22 = 0.5;
const cw22 = [1.2, 2.2, 2.2, 3.7];
["阶段", "达人动作", "矩阵动作", "关键目标"].forEach((h, i) => {
  p22.addShape(pres.shapes.RECTANGLE, { x: cx22, y: 1.3, w: cw22[i], h: 0.5, fill: { color: COLORS.secondary } });
  p22.addText(h, { x: cx22, y: 1.38, w: cw22[i], h: 0.4, fontSize: 12, fontFace: "Arial", color: COLORS.white, bold: true, align: "center", valign: "middle" });
  cx22 += cw22[i];
});
let cy22 = 1.85;
[
  ["W1", "头部达人合作发布", "品牌主号宣发+门店/职人转发", "品牌声量爆发"],
  ["W2", "腰部达人集中投放", "门店号跟进+本地达人二探", "到店转化提升"],
  ["W3", "素人体验官UGC收集", "职人IP知识科普承接流量", "口碑沉淀"],
  ["W4", "数据复盘优化Brief", "月度爆款盘点+SOP沉淀", "持续优化"]
].forEach((r, i) => {
  cx22 = 0.5;
  r.forEach((c, j) => {
    p22.addShape(pres.shapes.RECTANGLE, { x: cx22, y: cy22, w: cw22[j], h: 0.65, fill: { color: i % 2 === 0 ? COLORS.light : COLORS.white } });
    p22.addText(c, { x: cx22 + 0.1, y: cy22 + 0.2, w: cw22[j] - 0.2, h: 0.3, fontSize: 11, fontFace: "Arial", color: j === 0 ? COLORS.primary : (j === 3 ? COLORS.accent : COLORS.dark), bold: j === 0 || j === 3, align: j === 0 ? "center" : "left" });
    cx22 += cw22[j];
  });
  cy22 += 0.7;
});
p22.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.7, w: 9, h: 1.1, fill: { color: COLORS.white } });
p22.addText("达人+矩阵联动打法", { x: 0.7, y: 4.8, w: 8.6, h: 0.35, fontSize: 16, fontFace: "Arial", color: COLORS.primary, bold: true });
p22.addText("达人产出爆款 → 品牌主号转发 → 门店号挂载团购 → 职人IP二次分发 → DOU+追投 → 沉淀素材库", { x: 0.7, y: 5.2, w: 8.6, h: 0.4, fontSize: 12, fontFace: "Arial", color: COLORS.dark, align: "center" });

// 执行-3
let p23 = pres.addSlide();
addTitle(p23, "07  执行落地与阶段目标", "（3/3）与TCJF差异总结");
p23.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.3, w: 9, h: 0.8, fill: { color: COLORS.accent } });
p23.addText("Mr.JUDY vs TCJF 核心差异", { x: 0.5, y: 1.5, w: 9, h: 0.45, fontSize: 24, fontFace: "Arial Black", color: COLORS.white, bold: true, align: "center" });

let dy23 = 2.3;
[
  ["扩张模式", "云连锁为主", "先单店跑通再扩张", "稳扎稳打"],
  ["定价策略", "低价高频引流", "体验价值中高端", "品质定位"],
  ["效果展示", "发型即时可见", "过程+长期效果", "专业可信"],
  ["目标人群", "大众市场", "精准白领女性", "高LTV人群"],
  ["核心壁垒", "价格+便利", "体验+专业", "差异化护城河"]
].forEach((r, i) => {
  p23.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: dy23, w: 2.0, h: 0.6, fill: { color: COLORS.secondary } });
  p23.addText(r[0], { x: 0.5, y: dy23 + 0.15, w: 2.0, h: 0.35, fontSize: 13, fontFace: "Arial", color: COLORS.white, bold: true, align: "center", valign: "middle" });
  p23.addShape(pres.shapes.RECTANGLE, { x: 2.6, y: dy23, w: 3.2, h: 0.6, fill: { color: COLORS.light } });
  p23.addText(r[1], { x: 2.6, y: dy23 + 0.18, w: 3.2, h: 0.3, fontSize: 12, fontFace: "Arial", color: COLORS.gray, align: "center" });
  p23.addShape(pres.shapes.RECTANGLE, { x: 5.9, y: dy23, w: 3.2, h: 0.6, fill: { color: COLORS.light } });
  p23.addText(r[2], { x: 5.9, y: dy23 + 0.18, w: 3.2, h: 0.3, fontSize: 12, fontFace: "Arial", color: COLORS.accent, bold: true, align: "center" });
  dy23 += 0.7;
});

p23.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 5.0, w: 9, h: 0.9, fill: { color: COLORS.primary } });
p23.addText("核心结论：借鉴TCJF云连锁模式，但走差异化路线——以体验价值驱动，打造精致护理专家定位", { x: 0.5, y: 5.25, w: 9, h: 0.45, fontSize: 14, fontFace: "Arial", color: COLORS.white, bold: true, align: "center" });

// ========== 结束页 ==========
let p24 = pres.addSlide();
p24.background = { color: COLORS.primary };
p24.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.15, fill: { color: COLORS.accent } });
p24.addText("感谢观看", { x: 0.5, y: 2.2, w: 9, h: 1, fontSize: 48, fontFace: "Arial Black", color: COLORS.white, bold: true, align: "center" });
p24.addText("Mr.JUDY 抖音全案运营方案", { x: 0.5, y: 3.3, w: 9, h: 0.6, fontSize: 24, fontFace: "Arial", color: COLORS.accent, align: "center" });
p24.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.4, w: 10, h: 0.225, fill: { color: COLORS.accent } });

// 保存
pres.writeFile({ fileName: "/Users/flyingspur/.qclaw/workspace-agent-6f07984c/MrJUDY_抖音全案运营方案_补充页.pptx" });
console.log("补充页已生成（13页）：板块四到七 + 结束页");
