const pptxgen = require("pptxgenjs");

const C = {
  primary:   "1B7F5E",
  secondary: "3ECFA0",
  accent:    "F4A261",
  dark:      "1A3C34",
  light:     "F0FAF5",
  white:     "FFFFFF",
  gray:      "6B7C7A",
  lightGray: "E8F4EF",
  red:       "E63946",
  gold:      "E9C46A",
};

function shadow() {
  return { type: "outer", color: "000000", blur: 8, offset: 2, angle: 135, opacity: 0.12 };
}
function card(slide, x, y, w, h, fillColor, lineColor) {
  slide.addShape("rect", { x, y, w, h, fill: { color: fillColor || C.white }, line: { color: lineColor || C.lightGray, width: 1 }, shadow: shadow() });
}
function leftAccent(slide, x, y, h, color) {
  slide.addShape("rect", { x, y, w: 0.08, h, fill: { color }, line: { width: 0 } });
}
function slideHeader(slide, title, subtitle) {
  slide.addShape("rect", { x: 0, y: 0, w: 10, h: 1.1, fill: { color: C.dark }, line: { color: C.dark, width: 0 } });
  slide.addText(title, { x: 0.45, y: 0.18, w: 9, h: 0.55, fontSize: 24, bold: true, color: C.white, margin: 0 });
  if (subtitle) slide.addText(subtitle, { x: 0.45, y: 0.68, w: 9, h: 0.35, fontSize: 13, color: C.secondary, margin: 0 });
}
function footer(slide, page) {
  slide.addShape("rect", { x: 0, y: 5.38, w: 10, h: 0.25, fill: { color: C.lightGray }, line: { width: 0 } });
  slide.addText("荟宝团单升级培训 | 内部资料", { x: 0.4, y: 5.4, w: 6, h: 0.2, fontSize: 9, color: C.gray, margin: 0 });
  slide.addText(String(page), { x: 9.2, y: 5.4, w: 0.5, h: 0.2, fontSize: 9, color: C.gray, align: "right", margin: 0 });
}
function tag(slide, text, x, y, bg) {
  slide.addShape("roundRect", { x, y, w: 1.6, h: 0.28, fill: { color: bg }, line: { width: 0 }, rectRadius: 0.05 });
  slide.addText(text, { x, y: y + 0.02, w: 1.6, h: 0.24, fontSize: 9, bold: true, color: C.white, align: "center", margin: 0 });
}

// ═══════════════════════════════════════════════════════════════
// BUILD PRESENTATION
// ═══════════════════════════════════════════════════════════════
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "荟宝团单升级解读";
pres.author = "荟宝品牌";

// ── Slide 1: 封面 ──────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape("ellipse", { x: -2.5, y: 2.5, w: 7, h: 7, fill: { color: C.primary, transparency: 30 }, line: { width: 0 } });
  s.addShape("ellipse", { x: 7.5, y: -1.5, w: 4, h: 4, fill: { color: C.secondary, transparency: 20 }, line: { width: 0 } });
  s.addText("荟宝团单升级解读", { x: 0.6, y: 1.6, w: 9, h: 1.2, fontSize: 48, bold: true, color: C.white, margin: 0 });
  s.addText("服务标准化 × 转化再升级", { x: 0.6, y: 2.75, w: 6, h: 0.5, fontSize: 22, color: C.secondary, margin: 0 });
  s.addShape("rect", { x: 0.6, y: 3.35, w: 2, h: 0.06, fill: { color: C.accent }, line: { width: 0 } });
  s.addText("门店线上培训", { x: 0.6, y: 3.55, w: 4, h: 0.4, fontSize: 16, color: C.white, margin: 0 });
  s.addText("美业连锁品牌增长方案", { x: 0.6, y: 4.6, w: 4, h: 0.35, fontSize: 12, color: "8FBAA8", margin: 0 });
  s.addText("内部培训资料 | 请勿外传", { x: 6.5, y: 4.6, w: 3, h: 0.35, fontSize: 11, color: "5E8A7C", align: "right", margin: 0 });
}

// ── Slide 2: 目录 ──────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.light };
  slideHeader(s, "今日内容概览", "4大模块 · 系统掌握新团单");
  const modules = [
    { num: "01", title: "行业现状", sub: "为什么转化低、客户体验差？" },
    { num: "02", title: "销售心理", sub: "大品牌引流怎么做客户体验？" },
    { num: "03", title: "话术体系", sub: "邀约·到店·转化·跟进全链路" },
    { num: "04", title: "团单详解", sub: "新团单产品+佣金+操作SOP" },
  ];
  modules.forEach((m, i) => {
    const x = 0.5 + (i % 2) * 4.7;
    const y = 1.4 + Math.floor(i / 2) * 1.95;
    card(s, x, y, 4.3, 1.7);
    s.addText(m.num, { x: x + 0.2, y: y + 0.2, w: 0.8, h: 0.8, fontSize: 32, bold: true, color: C.secondary, margin: 0 });
    s.addText(m.title, { x: x + 1.1, y: y + 0.25, w: 3, h: 0.45, fontSize: 18, bold: true, color: C.dark, margin: 0 });
    s.addText(m.sub, { x: x + 1.1, y: y + 0.7, w: 3, h: 0.4, fontSize: 11, color: C.gray, margin: 0 });
    s.addShape("rect", { x, y: y + 1.45, w: 4.3, h: 0.25, fill: { color: C.lightGray }, line: { width: 0 } });
    s.addShape("rect", { x, y: y + 1.45, w: 4.3 * ((i + 1) / 4), h: 0.25, fill: { color: C.secondary }, line: { width: 0 } });
  });
  footer(s, 2);
}

// ── Slide 3: 美业现状痛点 ──────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.light };
  slideHeader(s, "模块一：行业现状", "转化低、体验差、客户少——根本原因是什么？");
  s.addText("行业\n病了", { x: 0.4, y: 1.4, w: 2.4, h: 2, fontSize: 48, bold: true, color: C.red, align: "center", margin: 0 });
  const painPoints = [
    { icon: "⚠", title: "过度推销", desc: "进店像进了菜市场，顾问轮番轰炸，客户只想逃" },
    { icon: "⏰", title: "时间陷阱", desc: "项目说30分钟，实际拖到90分钟，客户对时间失去信任" },
    { icon: "💸", title: "价格套路", desc: "低价引流到店，升单话术连环套，客户感觉被宰" },
    { icon: "😤", title: "敷衍服务", desc: "引流客被当二等公民，服务缩水，客户感受不到价值" },
    { icon: "🔄", title: "只管拉新", desc: "不维护、不跟进、不建信任，复购率极低" },
    { icon: "📉", title: "品牌受损", desc: "差评传播快，抖音、小红书一条差评毁所有" },
  ];
  painPoints.forEach((p, i) => {
    const x = 0.5 + (i % 3) * 3.1;
    const y = 1.4 + Math.floor(i / 3) * 2.0;
    card(s, x, y, 2.9, 1.8);
    s.addShape("ellipse", { x: x + 1.1, y: y + 0.15, w: 0.65, h: 0.65, fill: { color: "FEE2E2" }, line: { width: 0 } });
    s.addText(p.icon, { x: x + 1.1, y: y + 0.17, w: 0.65, h: 0.65, fontSize: 22, align: "center", valign: "middle", margin: 0 });
    s.addText(p.title, { x: x + 0.15, y: y + 0.9, w: 2.6, h: 0.35, fontSize: 14, bold: true, color: C.dark, margin: 0 });
    s.addText(p.desc, { x: x + 0.15, y: y + 1.22, w: 2.6, h: 0.55, fontSize: 10, color: C.gray, margin: 0 });
  });
  footer(s, 3);
}

// ── Slide 4: 客户体验差在哪 ────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  slideHeader(s, "模块一：客户体验差的真相", "客户不是不想变美，是不想被「推销」");
  s.addText("客户的真实心理", { x: 0.4, y: 1.25, w: 4, h: 0.4, fontSize: 15, bold: true, color: C.dark, margin: 0 });
  const customerMind = [
    "「我知道你们要推卡，但别太明显」",
    "「我时间很贵，别浪费我90分钟」",
    "「你越热情，我越觉得有套路」",
    "「我不想当场做决定，给我时间考虑」",
    "「如果体验真的好，我会回来的」",
    "「你推荐的是我需要的，不是我最贵的」",
  ];
  customerMind.forEach((t, i) => {
    const y = 1.7 + i * 0.58;
    s.addShape("rect", { x: 0.4, y, w: 4.4, h: 0.48, fill: { color: i % 2 === 0 ? C.lightGray : C.white }, line: { width: 0 } });
    s.addText(t, { x: 0.55, y: y + 0.06, w: 4.1, h: 0.38, fontSize: 11, color: C.dark, margin: 0 });
  });
  s.addText("销售常见的「自杀式」行为", { x: 5.1, y: 1.25, w: 4.5, h: 0.4, fontSize: 15, bold: true, color: C.red, margin: 0 });
  const wrongBehaviors = [
    { b: "刚躺下就开始推卡", r: "信任还没建立，客户直接关闭耳朵" },
    { b: "用「今天特价」制造紧迫感", r: "客户感觉被操控，更加不信任" },
    { b: "反复强调「错过就没了」", r: "越催促，客户越想逃离现场" },
    { b: "同时上两个顾问轮流说", r: "压力升级，客户内心抵触翻倍" },
    { b: "升单目标写在脸上", r: "客户感受不到服务，只看到销售" },
    { b: "服务过程不断接打电话", r: "客户觉得被打扰，价值感降低" },
  ];
  wrongBehaviors.forEach((item, i) => {
    const y = 1.7 + i * 0.58;
    card(s, 5.1, y, 4.4, 0.52, "FEF2F2");
    leftAccent(s, 5.1, y, 0.52, C.red);
    s.addText(item.b, { x: 5.28, y: y + 0.04, w: 2.3, h: 0.25, fontSize: 11, bold: true, color: C.dark, margin: 0 });
    s.addText(item.r, { x: 5.28, y: y + 0.26, w: 4.1, h: 0.22, fontSize: 9.5, color: C.red, margin: 0 });
  });
  footer(s, 4);
}

// ── Slide 5: 大品牌引流案例 ────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  slideHeader(s, "模块二：销售心理 · 大品牌引流怎么做", "让客户感觉「占了便宜」，而不是「被薅了羊毛」");
  const brands = [
    { name: "屈臣氏", approach: "试用装体验→BA只说一句话→客户自主选择", psychology: "降低决策门槛，不施压，让客户主动说「我要买」", color: "4A90D9" },
    { name: "海底捞", approach: "等位免费→小菜无限续→服务感动→主动好评", psychology: "超预期服务在前，客户回报在后，从不主动提「好评」", color: "E67E22" },
    { name: "Apple", approach: "产品体验区→员工不推销→问题引导→自然成交", psychology: "「天才吧」员工不背KPI，专注解决问题，销售自然发生", color: "27AE60" },
    { name: "7-11", approach: "试吃→陈列有门道→店员热情招呼→复购", psychology: "「欢迎光临」每天说200遍，温度感建立信任感", color: "E74C3C" },
  ];
  brands.forEach((b, i) => {
    const x = 0.4 + i * 2.4;
    card(s, x, 1.25, 2.2, 3.9);
    s.addShape("rect", { x, y: 1.25, w: 2.2, h: 0.55, fill: { color: b.color }, line: { width: 0 } });
    s.addText(b.name, { x, y: 1.32, w: 2.2, h: 0.42, fontSize: 16, bold: true, color: C.white, align: "center", margin: 0 });
    s.addText("方法", { x: x + 0.15, y: 1.95, w: 0.7, h: 0.26, fontSize: 9, bold: true, color: b.color, margin: 0 });
    s.addText(b.approach, { x: x + 0.15, y: 2.2, w: 1.95, h: 1.1, fontSize: 10, color: C.dark, margin: 0 });
    s.addText("底层逻辑", { x: x + 0.15, y: 3.35, w: 1.5, h: 0.26, fontSize: 9, bold: true, color: C.accent, margin: 0 });
    s.addText(b.psychology, { x: x + 0.15, y: 3.6, w: 1.95, h: 1.1, fontSize: 10, color: C.gray, margin: 0 });
  });
  s.addShape("rect", { x: 0.4, y: 5.1, w: 9.2, h: 0.25, fill: { color: C.dark }, line: { width: 0 } });
  s.addText("核心原则：先让客户感受到价值，再谈价格。先建立信任，再自然转化。", { x: 0.5, y: 5.12, w: 9, h: 0.22, fontSize: 10.5, bold: true, color: C.white, margin: 0 });
  footer(s, 5);
}

// ── Slide 6: 销售心理模型 ─────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.light };
  slideHeader(s, "模块二：销售心理模型", "销售不是「说服」，是「引导」——让客户自己说服自己");
  const levels = [
    { label: "信任建立", color: C.primary, y: 1.35, items: ["真诚的服务态度", "专业的问题诊断", "用产品效果说话"] },
    { label: "需求共鸣", color: "2E86AB", y: 2.15, items: ["问对问题（肤质/习惯/期待）", "复述确认（让客户感到被理解）", "共情表达（「我理解您的顾虑」）"] },
    { label: "价值呈现", color: C.secondary, y: 2.95, items: ["用对比展示效果（Before/After）", "用案例证明（其他客户的改变）", "用专业解释原理（成分/技术）"] },
    { label: "选择引导", color: C.accent, y: 3.75, items: ["给2-3个选项，不要只给一个", "帮客户分析利弊，不替客户选", "用「我们推荐」而非「你必须」"] },
    { label: "承诺闭环", color: C.gold, y: 4.55, items: ["当场小承诺（下次预约时间）", "书面确认（项目卡/次数卡）", "跟进承诺（效果跟踪/复诊提醒）"] },
  ];
  levels.forEach((l, i) => {
    const w = 3.8 - i * 0.3;
    const x = 0.5 + i * 0.15;
    s.addShape("rect", { x, y: l.y, w, h: 0.68, fill: { color: l.color, transparency: 15 }, line: { color: l.color, width: 1 } });
    s.addText(l.label, { x: x + 0.1, y: l.y + 0.08, w: 1.2, h: 0.28, fontSize: 12, bold: true, color: l.color, margin: 0 });
    s.addText(l.items[0], { x: x + 1.3, y: l.y + 0.12, w: w - 1.4, h: 0.5, fontSize: 10, color: C.dark, margin: 0 });
  });
  s.addText("底层逻辑", { x: 5.1, y: 1.25, w: 4.5, h: 0.4, fontSize: 15, bold: true, color: C.dark, margin: 0 });
  const principles = [
    { title: "客户买的不是产品，是更好的自己", desc: "要卖愿景，不卖成分表" },
    { title: "价格由价值决定，不由成本决定", desc: "「贵」是因为你没说清楚为什么值" },
    { title: "拒绝不是「不要」，是「还没准备好」", desc: "给时间，不给压力，7天内必回访" },
    { title: "最好的销售是不像销售的销售", desc: "像朋友推荐，不像顾问在卖货" },
    { title: "转介绍才是最高质量的获客", desc: "服务好一个人，可能带来3个新客" },
  ];
  principles.forEach((p, i) => {
    const y = 1.72 + i * 0.73;
    card(s, 5.1, y, 4.4, 0.66);
    s.addShape("ellipse", { x: 5.22, y: y + 0.12, w: 0.42, h: 0.42, fill: { color: C.secondary, transparency: 20 }, line: { width: 0 } });
    s.addText(String(i + 1), { x: 5.22, y: y + 0.14, w: 0.42, h: 0.38, fontSize: 14, bold: true, color: C.primary, align: "center", margin: 0 });
    s.addText(p.title, { x: 5.74, y: y + 0.06, w: 3.65, h: 0.28, fontSize: 12, bold: true, color: C.dark, margin: 0 });
    s.addText(p.desc, { x: 5.74, y: y + 0.34, w: 3.65, h: 0.26, fontSize: 10, color: C.gray, margin: 0 });
  });
  footer(s, 6);
}

// ── Slide 7: 邀约话术 ─────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  slideHeader(s, "模块三：话术体系 · 邀约阶段", "邀约的核心：给客户一个「无法拒绝」的理由");
  s.addText("邀约表达逻辑模型", { x: 0.4, y: 1.25, w: 4.3, h: 0.38, fontSize: 14, bold: true, color: C.primary, margin: 0 });
  const formula = [
    { step: "① 身份确认", content: "您好，请问是XX女士吗？我是荟宝的美容顾问小X，您上次在我们这里做了XX项目，还记得吗？" },
    { step: "② 价值告知", content: "这次联系您，是因为我们刚到了一批适合您肤质的王牌灌肤套盒，很多客户用了反馈特别好" },
    { step: "③ 稀缺营造", content: "这次团单活动仅限XX日前，名额有限，我已经帮您预留了一个" },
    { step: "④ 行动触发", content: "您看这周六还是周日下午方便？我帮您预约一个专属时段，不用排队等候" },
  ];
  formula.forEach((f, i) => {
    const y = 1.82 + i * 0.88;
    card(s, 0.55, y, 4.1, 0.78);
    leftAccent(s, 0.55, y, 0.78, C.primary);
    s.addText(f.step, { x: 0.72, y: y + 0.06, w: 1.6, h: 0.26, fontSize: 10, bold: true, color: C.primary, margin: 0 });
    s.addText(f.content, { x: 0.72, y: y + 0.3, w: 3.8, h: 0.46, fontSize: 9.5, color: C.dark, margin: 0 });
  });
  s.addText("不同场景话术示例", { x: 5.1, y: 1.25, w: 4.5, h: 0.38, fontSize: 14, bold: true, color: C.dark, margin: 0 });
  const scenes = [
    { title: "场景：新客首邀", color: C.secondary, lines: [
      "「您好，是XX小姐吗？我是荟宝的美容顾问小王。看到您在平台上购买了我们29.9元的新客体验套餐，首先感谢您的信任！」",
      "「我想提前了解一下您的肤质情况，这样到店后我们可以为您节省时间，做更针对性的服务。请问您平时皮肤最困扰的问题是什么？」",
      "「好的，我了解了。我们这次体验包含芦荟灌肤补水，是非常温和适合您肤质的项目。我帮您预约周六上午10点可以吗？这个时段我们有独立的护理房间，不用等候。」"
    ]},
    { title: "场景：老客复邀", color: C.accent, lines: [
      "「您好陈姐！我是荟宝小王，还记得我吗？您上次来做的是深层清洁项目，效果还满意吧？」",
      "「这次联系您是因为我们上线了一个王牌灌肤3次卡，3次护理总共169元，折算下来单次不到57元。很多老客户都在囤，您要不要也来体验一下？」",
      "「您看这周有时间吗？我看您上次做的清洁效果很好，这次可以安排一个深层补水+灌肤，给皮肤做个完整的养护。」"
    ]},
  ];
  scenes.forEach((scene, i) => {
    const y = 1.72 + i * 1.75;
    card(s, 5.1, y, 4.4, 1.62);
    s.addShape("rect", { x: 5.1, y, w: 4.4, h: 0.32, fill: { color: scene.color }, line: { width: 0 } });
    s.addText(scene.title, { x: 5.2, y: y + 0.04, w: 4.2, h: 0.26, fontSize: 11, bold: true, color: C.white, margin: 0 });
    scene.lines.forEach((line, j) => {
      s.addText((j + 1) + ". " + line, { x: 5.2, y: y + 0.38 + j * 0.4, w: 4.2, h: 0.38, fontSize: 9, color: C.dark, margin: 0 });
    });
  });
  footer(s, 7);
}

// ── Slide 8: 到店接待话术 ─────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.light };
  slideHeader(s, "模块三：话术体系 · 到店接待阶段", "前5分钟决定客户会不会信任你");
  s.addText("接待表达逻辑模型", { x: 0.4, y: 1.25, w: 4.3, h: 0.38, fontSize: 14, bold: true, color: C.primary, margin: 0 });
  const reception = [
    { step: "① 热情迎宾（0-2分钟）", content: "客户进店→微笑迎接→引导就座→递上茶水→询问等待感受", note: "核心：让客户感到被重视，不是被接待" },
    { step: "② 需求了解（2-5分钟）", content: "「您好，欢迎来到荟宝！我是今天为您服务的美容顾问，先给您倒杯温水」", note: "核心：问诊式了解肤质，不推产品，先建立信任" },
    { step: "③ 方案呈现（5-10分钟）", content: "根据诊断结果，推荐适合的团单项目，解释价值和预期效果", note: "核心：说清楚「为什么适合」，不是「为什么贵」" },
    { step: "④ 疑虑处理（10-15分钟）", content: "客户提出顾虑→共情回应→提供证明→给选择权", note: "核心：每一个拒绝背后都有一个未被满足的需求" },
  ];
  reception.forEach((r, i) => {
    const y = 1.72 + i * 0.9;
    card(s, 0.4, y, 4.4, 0.82);
    leftAccent(s, 0.4, y, 0.82, C.secondary);
    s.addText(r.step, { x: 0.57, y: y + 0.06, w: 4.1, h: 0.26, fontSize: 10, bold: true, color: C.dark, margin: 0 });
    s.addText(r.content, { x: 0.57, y: y + 0.32, w: 4.1, h: 0.25, fontSize: 9.5, color: C.primary, margin: 0 });
    s.addText(r.note, { x: 0.57, y: y + 0.56, w: 4.1, h: 0.22, fontSize: 9, color: C.gray, italic: true, margin: 0 });
  });
  s.addText("接待关键话术", { x: 5.1, y: 1.25, w: 4.5, h: 0.38, fontSize: 14, bold: true, color: C.dark, margin: 0 });
  const tips = [
    { q: "客户说「我只是来体验的」", a: "「完全理解！我们的体验项目就是专门为新客户设计的，先感受再做决定。您今天皮肤状态怎么样？我先帮您做个诊断，让服务更有针对性。」", color: C.secondary },
    { q: "客户问「会不会一直推销」", a: "「我非常理解您的顾虑。我们店的理念是：先让您感受到效果，您自己觉得好再考虑其他。今天我的任务就是让您满意，不是让您买卡。」", color: C.accent },
    { q: "客户说「我再考虑一下」", a: "「没问题，考虑一下很正常。今天的项目您先好好体验，效果会说话。下周我会给您发一条护理小贴士，到时候您有任何问题随时联系我。」", color: C.primary },
    { q: "客户对价格犹豫", a: "「您看重的是效果对吧？今天这个29.9元项目包含了鲜芦荟灌肤和济州岛面膜，市场上单次芦荟灌肤都要80以上，这个价格是品牌补贴体验价的。您先体验，感受好了再说。」", color: "8B5CF6" },
  ];
  tips.forEach((t, i) => {
    const y = 1.72 + i * 0.88;
    card(s, 5.1, y, 4.4, 0.8);
    s.addShape("rect", { x: 5.1, y, w: 4.4, h: 0.28, fill: { color: t.color, transparency: 85 }, line: { width: 0 } });
    s.addText("Q：" + t.q, { x: 5.2, y: y + 0.04, w: 4.2, h: 0.22, fontSize: 9.5, bold: true, color: C.dark, margin: 0 });
    s.addText("A：" + t.a, { x: 5.2, y: y + 0.3, w: 4.2, h: 0.48, fontSize: 9, color: C.gray, margin: 0 });
  });
  footer(s, 8);
}

// ── Slide 9: 转化升单话术 ──────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  slideHeader(s, "模块三：话术体系 · 转化升单阶段", "服务结束时，才是销售的开始");
  s.addText("转化表达逻辑模型", { x: 0.4, y: 1.25, w: 4.3, h: 0.38, fontSize: 14, bold: true, color: C.primary, margin: 0 });
  const convert = [
    { step: "① 效果回顾（服务结束前5分钟）", content: "「今天您的皮肤整体偏干，T区油脂分泌旺盛，我用了荟宝芦荟鲜汁帮您做了深层补水，您感觉怎么样？」", logic: "用服务结果说话，让客户自己感受到价值" },
    { step: "② 问题延伸", content: "「您平时有没有觉得皮肤在换季的时候特别敏感，或者熬夜后肤色暗沉的情况？」", logic: "发现问题，而不是制造焦虑；延伸需求，而不是强行升单" },
    { step: "③ 方案推荐", content: "「根据您今天的情况，我建议您可以尝试我们的王牌灌肤3次卡，169元3次，平均每次不到57元，配合您的深层清洁效果会更好。」", logic: "用「平均每次」降低感知价格；用「效果更好」提升感知价值" },
    { step: "④ 限时锁定", content: "「这个团单这周活动价，下周就恢复了。您看要不要今天先锁定一次？反正可以分3次使用，不着急用完。」", logic: "给理由，不给压力；给选择权，不给紧迫感" },
  ];
  convert.forEach((c, i) => {
    const y = 1.72 + i * 0.9;
    card(s, 0.4, y, 4.4, 0.82);
    leftAccent(s, 0.4, y, 0.82, C.accent);
    s.addText(c.step, { x: 0.57, y: y + 0.06, w: 4.1, h: 0.26, fontSize: 10, bold: true, color: C.dark, margin: 0 });
    s.addText(c.content, { x: 0.57, y: y + 0.32, w: 4.1, h: 0.26, fontSize: 9, color: C.primary, margin: 0 });
    s.addText("◆ " + c.logic, { x: 0.57, y: y + 0.57, w: 4.1, h: 0.22, fontSize: 8.5, color: C.gray, italic: true, margin: 0 });
  });
  s.addText("升单话术示例", { x: 5.1, y: 1.25, w: 4.5, h: 0.38, fontSize: 14, bold: true, color: C.dark, margin: 0 });
  const upSell = [
    { from: "29.9元引流 → 58元2次卡", script: "「您今天体验的芦荟灌肤效果非常好，皮肤明显透亮了。我们有一个王牌清洁+灌肤的2次卡，58元2次，第一次做深层清洁，第二次做赋活灌肤，配合起来效果可以维持更久。您要不要试试？」", reason: "从单次体验升级到多次卡，提升复购和到店频次" },
    { from: "58元2次卡 → 169元3次卡", script: "「您这两次护理下来效果真的很不错！皮肤水润度上来了。我们有一个王牌灌肤3次卡，169元3次，折算下来比单次购买便宜30%，而且3次护理可以分一个月做完，循序渐进效果更稳定。」", reason: "用价格对比凸显价值；用时间维度建立持续改善的期待" },
    { from: "任意团单 → 299元综合卡", script: "「您今天的护理做得挺全面的。我们有一个综合清洁+补水+眼护+屏障修复的299元套餐，原本单独做这些项目要500多。您之前有没有做过眼部和屏障修复的项目？这两项对整体肤质改善特别重要。」", reason: "用「套餐总价值」对比「单购总价」，让优惠感更强" },
  ];
  upSell.forEach((u, i) => {
    const y = 1.72 + i * 1.15;
    card(s, 5.1, y, 4.4, 1.05);
    s.addShape("rect", { x: 5.1, y, w: 4.4, h: 0.26, fill: { color: C.gold }, line: { width: 0 } });
    s.addText(u.from, { x: 5.2, y: y + 0.04, w: 4.2, h: 0.2, fontSize: 10, bold: true, color: C.dark, margin: 0 });
    s.addText(u.script, { x: 5.2, y: y + 0.3, w: 4.2, h: 0.42, fontSize: 8.5, color: C.dark, margin: 0 });
    s.addText("→ " + u.reason, { x: 5.2, y: y + 0.74, w: 4.2, h: 0.28, fontSize: 8.5, color: C.primary, italic: true, margin: 0 });
  });
  footer(s, 9);
}

// ── Slide 10: 跟进话术 ─────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.light };
  slideHeader(s, "模块三：话术体系 · 跟进维护阶段", "80%的门店死在「成交后不跟进」这一步");
  s.addText("跟进表达逻辑模型", { x: 0.4, y: 1.25, w: 4.3, h: 0.38, fontSize: 14, bold: true, color: C.primary, margin: 0 });
  const followup = [
    { time: "当天离店后 1小时内", type: "满意度回访", script: "「您好陈姐，今天护理结束后感觉怎么样？皮肤有没有觉得舒服一些？有什么不适感随时告诉我，另外荟宝的面膜今晚可以正常使用，早晚各一次效果更好哦！」", goal: "确认满意度，防止差评，表达持续关注" },
    { time: "第3天", type: "护理提醒", script: "「陈姐好，这几天天气比较干燥，我发一个护理小贴士给您：洁面后先用荟宝精华水轻拍3遍，再敷补水面膜，效果会更好。如果有晒后泛红的情况，芦荟胶可以厚敷急救。」", goal: "建立专业形象，不推产品，提供价值内容" },
    { time: "第7天", type: "效果跟进", script: "「陈姐，下周您的第二次护理时间到了。上次做完深层清洁后皮肤状态应该不错，这次我们做赋活灌肤，效果会进一步提升。我帮您预约同一个时段，还是上次的小美为您服务。」", goal: "自然引导复诊，锁定下次到店时间" },
    { time: "第30天", type: "复购激活", script: "「陈姐，您的3次卡还剩1次，下周是最佳使用时间。您的皮肤经过这两次护理，屏障功能已经改善很多，这次做完可以让效果更稳定。我帮您约好？」", goal: "激活沉默客户，推动最后1次使用完成" },
  ];
  followup.forEach((f, i) => {
    const y = 1.72 + i * 0.88;
    card(s, 0.4, y, 4.4, 0.8);
    s.addShape("rect", { x: 0.4, y, w: 4.4, h: 0.28, fill: { color: C.primary }, line: { width: 0 } });
    s.addText(f.time + "  |  " + f.type, { x: 0.5, y: y + 0.04, w: 4.2, h: 0.22, fontSize: 9.5, bold: true, color: C.white, margin: 0 });
    s.addText(f.script, { x: 0.5, y: y + 0.3, w: 4.2, h: 0.3, fontSize: 8.5, color: C.dark, margin: 0 });
    s.addText("目标：" + f.goal, { x: 0.5, y: y + 0.6, w: 4.2, h: 0.18, fontSize: 8.5, color: C.secondary, italic: true, margin: 0 });
  });
  s.addText("跟进注意事项", { x: 5.1, y: 1.25, w: 4.5, h: 0.38, fontSize: 14, bold: true, color: C.dark, margin: 0 });
  const cautions = [
    { icon: "❌", text: "不要在跟进信息里直接提「升单」或「续卡」", color: C.red },
    { icon: "❌", text: "不要在晚上10点后发消息，打扰客户休息", color: C.red },
    { icon: "❌", text: "不要发长语音，客户时间宝贵", color: C.red },
    { icon: "✅", text: "内容要有价值：护理知识/产品使用/天气提醒", color: C.secondary },
    { icon: "✅", text: "时间要合适：工作日10:00-11:30 / 15:00-17:00", color: C.secondary },
    { icon: "✅", text: "语气要像朋友关心，不像是销售在追单", color: C.secondary },
    { icon: "✅", text: "跟进频率：成交后3次（第1天/第3天/第7天）", color: C.secondary },
    { icon: "✅", text: "备注客户偏好，下次精准服务", color: C.secondary },
  ];
  cautions.forEach((c, i) => {
    const y = 1.72 + i * 0.44;
    s.addShape("rect", { x: 5.1, y, w: 4.4, h: 0.38, fill: { color: c.color === C.red ? "FEF2F2" : "F0FDF4" }, line: { width: 0 } });
    s.addText(c.icon, { x: 5.2, y: y + 0.05, w: 0.4, h: 0.3, fontSize: 13, margin: 0 });
    s.addText(c.text, { x: 5.58, y: y + 0.07, w: 3.85, h: 0.28, fontSize: 9.5, color: C.dark, margin: 0 });
  });
  footer(s, 10);
}

// ── Slide 11: 新团单过渡页 ─────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.dark };
  s.addText("模块四", { x: 0.4, y: 1.8, w: 9.2, h: 0.5, fontSize: 16, color: C.secondary, margin: 0 });
  s.addText("新团单矩阵全解读", { x: 0.4, y: 2.3, w: 9.2, h: 1, fontSize: 44, bold: true, color: C.white, margin: 0 });
  s.addShape("rect", { x: 0.4, y: 3.4, w: 2, h: 0.06, fill: { color: C.accent }, line: { width: 0 } });
  s.addText("从引流到利润，12个团单覆盖客户全生命周期", { x: 0.4, y: 3.6, w: 9, h: 0.5, fontSize: 16, color: "A8D5C2", margin: 0 });
  const layers = [
    { label: "引流层", price: "29.9元", color: C.secondary },
    { label: "锁客层", price: "58元2次卡", color: C.accent },
    { label: "利润层", price: "69-499元", color: C.gold },
  ];
  layers.forEach((l, i) => {
    const x = 0.5 + i * 3.1;
    s.addShape("rect", { x, y: 4.3, w: 2.8, h: 0.75, fill: { color: l.color, transparency: 70 }, line: { color: l.color, width: 1 } });
    s.addText(l.label + "  " + l.price, { x, y: 4.38, w: 2.8, h: 0.6, fontSize: 13, bold: true, color: C.white, align: "center", margin: 0 });
  });
  footer(s, 11);
}

// ── Slide 12: 引流团单详解 ─────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  slideHeader(s, "引流团单：29.9元芦荟灌肤（新客专享）", "使命：把人请进来，用30分钟证明价值");
  card(s, 0.4, 1.25, 4.4, 3.7);
  leftAccent(s, 0.4, 1.25, 3.7, C.secondary);
  s.addText("商品信息", { x: 0.57, y: 1.35, w: 4, h: 0.35, fontSize: 13, bold: true, color: C.primary, margin: 0 });
  const infoRows = [
    ["商品名称", "【春季焕新】新鲜芦荟灌肤补水护理30分钟"],
    ["商品类型", "新客专享·引流团单"],
    ["商品售价", "29.9元"],
    ["门店佣金率", "73%"],
    ["服务产品", "鲜芦荟+济州岛面膜+芦荟胶一支"],
    ["服务时长", "30分钟（标准化）"],
    ["备注", "广东区域不上线"],
  ];
  infoRows.forEach((row, i) => {
    const y = 1.75 + i * 0.48;
    s.addText(row[0], { x: 0.57, y, w: 1.5, h: 0.42, fontSize: 9.5, bold: true, color: C.gray, margin: 0 });
    s.addText(row[1], { x: 2.1, y, w: 2.6, h: 0.42, fontSize: 9.5, color: C.dark, margin: 0 });
  });
  s.addText("30分钟标准化SOP", { x: 5.1, y: 1.25, w: 4.5, h: 0.35, fontSize: 13, bold: true, color: C.primary, margin: 0 });
  const sop = [
    { step: "1", time: "5min", name: "面部诊断", desc: "了解肤质，记录问题点" },
    { step: "2", time: "3min", name: "卸妆洁面", desc: "荟宝洁面产品" },
    { step: "3", time: "1min", name: "面部爽肤", desc: "荟宝爽肤水" },
    { step: "4", time: "7min", name: "鲜汁补水", desc: "新鲜芦荟鲜汁直接作用于面部" },
    { step: "5", time: "15min", name: "面膜补水", desc: "荟宝济州岛面膜+冷喷" },
    { step: "6", time: "1min", name: "产品赠送", desc: "芦荟胶试用装一支带回家" },
  ];
  sop.forEach((item, i) => {
    const y = 1.68 + i * 0.6;
    s.addShape("rect", { x: 5.1, y, w: 0.65, h: 0.5, fill: { color: C.secondary }, line: { width: 0 } });
    s.addText(item.time, { x: 5.1, y: y + 0.12, w: 0.65, h: 0.28, fontSize: 9.5, bold: true, color: C.white, align: "center", margin: 0 });
    s.addShape("ellipse", { x: 5.85, y: y + 0.08, w: 0.35, h: 0.35, fill: { color: C.dark }, line: { width: 0 } });
    s.addText(item.step, { x: 5.85, y: y + 0.1, w: 0.35, h: 0.3, fontSize: 12, bold: true, color: C.white, align: "center", margin: 0 });
    s.addText(item.name, { x: 6.3, y: y + 0.04, w: 1.6, h: 0.26, fontSize: 11, bold: true, color: C.dark, margin: 0 });
    s.addText(item.desc, { x: 6.3, y: y + 0.27, w: 3.1, h: 0.22, fontSize: 9, color: C.gray, margin: 0 });
  });
  s.addShape("rect", { x: 0.4, y: 5.05, w: 9.2, h: 0.28, fill: { color: C.gold, transparency: 70 }, line: { width: 0 } });
  s.addText("⚠ 关键：引流团单不做强推，用专业诊断和效果赢得信任，让客户主动问你「下次怎么做」", { x: 0.5, y: 5.08, w: 9, h: 0.22, fontSize: 10, bold: true, color: C.dark, margin: 0 });
  footer(s, 12);
}

// ── Slide 13: 二转锁客团单 ─────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  slideHeader(s, "锁客团单：58元2次卡【双重焕肌体验】", "使命：从陌生客变成信任客，用两次服务完成转化");
  card(s, 0.4, 1.25, 4.4, 1.8);
  leftAccent(s, 0.4, 1.25, 1.8, C.accent);
  s.addText("商品信息", { x: 0.57, y: 1.35, w: 4, h: 0.32, fontSize: 12, bold: true, color: C.primary, margin: 0 });
  s.addText([
    { text: "商品名称：", options: { bold: true } }, { text: "王牌清洁泥膜+芦荟灌肤+唤肤精油按摩（可用2次75分钟）", options: { breakLine: true } },
    { text: "商品售价：", options: { bold: true } }, { text: "58元（次卡）", options: { breakLine: true } },
    { text: "门店佣金率：", options: { bold: true } }, { text: "75%", options: { breakLine: true } },
    { text: "适用区域：", options: { bold: true } }, { text: "湖南/湖北/江西/广西/福建/陕西", options: {} },
  ], { x: 0.57, y: 1.72, w: 4, h: 1.2, fontSize: 10, color: C.dark, margin: 0 });
  s.addText("两次服务内容对比", { x: 5.1, y: 1.25, w: 4.5, h: 0.35, fontSize: 12, bold: true, color: C.dark, margin: 0 });
  const twoTimes = [
    { title: "护理一（35min）", items: ["深层清洁+芦荟灌肤+面膜补水", "适合首次到店体验客户", "建立服务标准认知"] },
    { title: "护理二（40min）", items: ["赋活灌肤+芦荟灌肤+面膜补水", "加入精华油开穴+9线刮排+拨筋", "深度放松，效果更显著"] },
  ];
  twoTimes.forEach((t, i) => {
    const y = 1.68 + i * 1.2;
    card(s, 5.1, y, 4.4, 1.05);
    s.addShape("rect", { x: 5.1, y, w: 4.4, h: 0.28, fill: { color: i === 0 ? C.secondary : C.accent }, line: { width: 0 } });
    s.addText(t.title, { x: 5.2, y: y + 0.04, w: 4.2, h: 0.22, fontSize: 10, bold: true, color: C.white, margin: 0 });
    t.items.forEach((item, j) => {
      s.addText("• " + item, { x: 5.2, y: y + 0.32 + j * 0.22, w: 4.2, h: 0.2, fontSize: 9, color: C.dark, margin: 0 });
    });
  });
  s.addText("竞品对比：芭芭多69元/2次卡", { x: 0.4, y: 3.15, w: 4.4, h: 0.35, fontSize: 12, bold: true, color: C.dark, margin: 0 });
  s.addShape("rect", { x: 0.4, y: 3.55, w: 4.4, h: 1.35, fill: { color: C.lightGray }, line: { width: 0 } });
  const comp = [
    ["项目", "荟宝", "芭芭多"],
    ["价格", "58元", "69元"],
    ["时长", "75分钟", "70分钟"],
    ["项目数", "3个", "2个"],
  ];
  comp.forEach((row, i) => {
    const y = 3.6 + i * 0.32;
    const tc = i === 0 ? C.white : C.dark;
    const bg = i === 0 ? C.primary : (i % 2 === 0 ? C.lightGray : C.white);
    row.forEach((cell, j) => {
      s.addText(cell, { x: 0.5 + j * 1.45, y, w: 1.45, h: 0.3, fontSize: 9, bold: i === 0, color: tc, align: "center", margin: 0 });
    });
    if (i < comp.length - 1) {
      s.addShape("rect", { x: 1.95, y: 3.6, w: 0.01, h: 1.25, fill: { color: C.gray }, line: { width: 0 } });
      s.addShape("rect", { x: 3.4, y: 3.6, w: 0.01, h: 1.25, fill: { color: C.gray }, line: { width: 0 } });
    }
  });
  s.addShape("rect", { x: 5.1, y: 3.95, w: 4.4, h: 1.4, fill: { color: "FFF7ED" }, line: { width: 0 } });
  leftAccent(s, 5.1, 3.95, 1.4, C.accent);
  s.addText("转化话术", { x: 5.27, y: 4.03, w: 4, h: 0.28, fontSize: 10, bold: true, color: C.accent, margin: 0 });
  s.addText("「您今天体验的深层清洁效果非常好，皮肤清爽了很多。我们有一个58元2次卡，包含清洁+灌肤+面膜，一共75分钟。第二次我们会加入精油按摩，效果会比第一次更明显。卡的有效期是90天，您什么时候方便来做第二次？」", { x: 5.27, y: 4.32, w: 4.1, h: 0.95, fontSize: 9, color: C.dark, margin: 0 });
  footer(s, 13);
}

// ── Slide 14: 利润团单矩阵 ─────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.light };
  slideHeader(s, "利润团单矩阵（69-299元区间）", "使命：让每一张进店的客户都有机会成为高价值客户");
  const profits = [
    { name: "69元净肌清洁", sub: "女士专享·深层清洁", commission: "83%", time: "35分钟", products: "清洁套盒+济州岛面膜", color: C.secondary, best: "日常到店引流" },
    { name: "99元男士清洁", sub: "男士专享·深层清洁+T区", commission: "83%", time: "40分钟", products: "清洁套盒+济州岛面膜", color: "3498DB", best: "开拓男性市场" },
    { name: "199元清洁补水3次卡", sub: "换季福利·清洁+补水3次", commission: "95%", time: "35分钟×3次", products: "清洁套盒+荟宝面膜", color: C.primary, best: "高复购锁客" },
    { name: "299元综合护理", sub: "清洁+补水+眼护+屏障修复", commission: "93%", time: "全套约90分钟", products: "清洁+灌肤+眼套+修护套+面膜", color: C.gold, best: "高价值升单" },
  ];
  profits.forEach((p, i) => {
    const x = 0.4 + i * 2.4;
    card(s, x, 1.25, 2.2, 3.7);
    s.addShape("rect", { x, y: 1.25, w: 2.2, h: 0.7, fill: { color: p.color }, line: { width: 0 } });
    s.addText(p.name.split("元")[0] + "元", { x, y: 1.3, w: 2.2, h: 0.6, fontSize: 22, bold: true, color: C.white, align: "center", margin: 0 });
    s.addText(p.sub, { x, y: 2.02, w: 2.2, h: 0.3, fontSize: 10, color: C.dark, align: "center", margin: 0 });
    const rows = [
      ["佣金率", p.commission],
      ["服务时长", p.time],
      ["配套产品", p.products],
      ["最佳用途", p.best],
    ];
    rows.forEach((r, j) => {
      const y = 2.4 + j * 0.6;
      s.addText(r[0], { x: x + 0.12, y, w: 0.8, h: 0.25, fontSize: 9, bold: true, color: C.gray, margin: 0 });
      s.addText(r[1], { x: x + 0.12, y: y + 0.25, w: 2, h: 0.32, fontSize: 9, color: C.dark, margin: 0 });
    });
  });
  s.addShape("rect", { x: 0.4, y: 5.05, w: 9.2, h: 0.28, fill: { color: C.dark }, line: { width: 0 } });
  s.addText("佣金率从73%提升至83%-95%，品牌从抽佣90%变为帮你赚更多——这才是新团单的核心价值", { x: 0.5, y: 5.08, w: 9, h: 0.22, fontSize: 10, bold: true, color: C.secondary, margin: 0 });
  footer(s, 14);
}

// ── Slide 15: 升单品矩阵 ───────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  slideHeader(s, "升单品 · 产品利润品矩阵", "使命：从服务客户到销售客户，从一次性到持续复购");
  const upsells = [
    { name: "499元焕肤体验", sub: "10项芦荟霜按摩+鲜汁补水+面膜", commission: "93%", products: "王牌灌肤套", target: "高消费力客户", color: C.gold },
    { name: "169元灌肤3次卡", sub: "芦荟霜按摩+鲜汁补水+面膜SPA×3次", commission: "95%", products: "王牌灌肤套", target: "已体验客户复购锁客", color: C.primary },
    { name: "99元芦荟胶三件套", sub: "悠活芦荟胶护肤三件套体验装", commission: "93%", products: "三件套产品", target: "产品利润品·伴手礼", color: C.secondary },
    { name: "100元代金券", sub: "全场通用，可叠加使用", commission: "93%", products: "无实物", target: "综合利润·灵活促销", color: C.accent },
    { name: "29.9元刷单团单", sub: "【码上特惠】基础清洁", commission: "93%", products: "无", target: "做评价·拉数据", color: "95A3B3" },
  ];
  upsells.forEach((u, i) => {
    const x = 0.4 + i * 1.9;
    card(s, x, 1.25, 1.7, 3.7);
    s.addShape("rect", { x, y: 1.25, w: 1.7, h: 0.65, fill: { color: u.color }, line: { width: 0 } });
    const pricePart = u.name.split("元")[0] + "元";
    s.addText(pricePart, { x, y: 1.3, w: 1.7, h: 0.3, fontSize: 16, bold: true, color: C.white, align: "center", margin: 0 });
    s.addText(u.name.split("元")[1] || "", { x, y: 1.58, w: 1.7, h: 0.28, fontSize: 9, color: C.white, align: "center", margin: 0 });
    const urows = [
      ["佣金率", u.commission],
      ["配套产品", u.products],
      ["目标客户", u.target],
    ];
    urows.forEach((r, j) => {
      const y = 2.0 + j * 0.6;
      s.addText(r[0], { x: x + 0.1, y, w: 0.7, h: 0.25, fontSize: 9, bold: true, color: C.gray, margin: 0 });
      s.addText(r[1], { x: x + 0.1, y: y + 0.25, w: 1.5, h: 0.3, fontSize: 9, color: C.dark, margin: 0 });
    });
    s.addText(u.sub, { x: x + 0.1, y: 3.85, w: 1.5, h: 0.8, fontSize: 8.5, color: C.gray, margin: 0 });
  });
  footer(s, 15);
}

// ── Slide 16: 新旧团单对比 ──────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.light };
  slideHeader(s, "新旧团单核心对比", "品牌从「抽你钱」变成「帮你赚钱」");
  s.addText("核心维度对比", { x: 0.4, y: 1.25, w: 9, h: 0.38, fontSize: 14, bold: true, color: C.dark, margin: 0 });
  const compRows = [
    { dim: "品牌抽佣", old: "品牌抽走90%，门店仅得10%", new: "品牌抽走7%-27%，门店到手73%-95%", change: "↑" },
    { dim: "服务时长", old: "引流团单流程冗长（60-90分钟）", new: "引流团单标准化30分钟，不拖沓", change: "↑" },
    { dim: "产品配套", old: "产品散乱，没有统一配套", new: "每个团单配套套盒+面膜+产品赠送", change: "↑" },
    { dim: "转化路径", old: "引流进来，直接推高客单，客户跑", new: "引流→锁客→利润三层漏斗，循序渐进", change: "↑" },
    { dim: "佣金递增", old: "无论价格高低，抽佣比例固定90%", new: "越高端的团单，门店佣金比例越高（83%-95%）", change: "↑" },
    { dim: "区域灵活", old: "全国统一，无法因地制宜", new: "区域差异化上线（广东暂不上线引流团单）", change: "↑" },
  ];
  compRows.forEach((r, i) => {
    const y = 1.7 + i * 0.58;
    const bg = i % 2 === 0 ? C.white : C.lightGray;
    s.addShape("rect", { x: 0.4, y, w: 9.2, h: 0.52, fill: { color: bg }, line: { width: 0 } });
    s.addText(r.dim, { x: 0.55, y: y + 0.08, w: 1.5, h: 0.36, fontSize: 11, bold: true, color: C.dark, margin: 0 });
    s.addShape("rect", { x: 2.1, y: y + 0.08, w: 4, h: 0.36, fill: { color: "FEF2F2" }, line: { width: 0 } });
    s.addText(r.old, { x: 2.2, y: y + 0.08, w: 3.8, h: 0.36, fontSize: 9.5, color: C.red, margin: 0 });
    s.addShape("rect", { x: 6.2, y: y + 0.08, w: 2.8, h: 0.36, fill: { color: "F0FDF4" }, line: { width: 0 } });
    s.addText(r.new, { x: 6.3, y: y + 0.08, w: 2.6, h: 0.36, fontSize: 9.5, color: C.primary, margin: 0 });
    s.addText(r.change, { x: 9.1, y: y + 0.08, w: 0.3, h: 0.36, fontSize: 16, bold: true, color: C.secondary, align: "center", margin: 0 });
  });
  footer(s, 16);
}

// ── Slide 17: 门店行动要求 ─────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  slideHeader(s, "门店行动要求", "培训只是开始，执行才是关键");
  s.addText("今日学习后，门店需要做的事", { x: 0.4, y: 1.25, w: 9, h: 0.38, fontSize: 14, bold: true, color: C.dark, margin: 0 });
  const actions = [
    { num: "1", title: "全员学习新团单SOP", desc: "今天起3天内，所有美容顾问必须熟记引流团单30分钟6步SOP，能完整复述", deadline: "3天内", color: C.primary },
    { num: "2", title: "话术实战演练", desc: "门店店长组织角色扮演，演练新客首邀+到店接待+转化升单三个场景，全员考核", deadline: "1周内", color: C.secondary },
    { num: "3", title: "老客邀约启动", desc: "梳理近3个月未到店老客名单，按新团单话术开始分批邀约，重点推58元2次卡和169元3次卡", deadline: "下周起", color: C.accent },
    { num: "4", title: "数据跟踪机制", desc: "每天登记团单转化数据：引流团单到店数、二转锁客率、升单金额，形成周报表", deadline: "持续", color: C.gold },
    { num: "5", title: "客户满意度跟进", desc: "所有成交客户，24小时内完成满意度回访（参照跟进话术模板），防止差评", deadline: "每日", color: "8B5CF6" },
    { num: "6", title: "区域差异化执行", desc: "广东区域暂不推引流团单，其他团单正常上线；其他区域全面执行新团单体系", deadline: "立即", color: C.red },
  ];
  actions.forEach((a, i) => {
    const x = 0.4 + (i % 3) * 3.1;
    const y = 1.72 + Math.floor(i / 3) * 1.8;
    card(s, x, y, 2.9, 1.65);
    s.addShape("ellipse", { x: x + 0.15, y: y + 0.15, w: 0.5, h: 0.5, fill: { color: a.color }, line: { width: 0 } });
    s.addText(a.num, { x: x + 0.15, y: y + 0.17, w: 0.5, h: 0.46, fontSize: 18, bold: true, color: C.white, align: "center", margin: 0 });
    s.addText(a.title, { x: x + 0.75, y: y + 0.2, w: 2, h: 0.4, fontSize: 12, bold: true, color: C.dark, margin: 0 });
    s.addText(a.desc, { x: x + 0.15, y: y + 0.75, w: 2.65, h: 0.65, fontSize: 9.5, color: C.gray, margin: 0 });
    s.addShape("rect", { x: x + 0.15, y: y + 1.4, w: 2.65, h: 0.22, fill: { color: a.color, transparency: 80 }, line: { width: 0 } });
    s.addText("⏰ " + a.deadline, { x: x + 0.25, y: y + 1.41, w: 2.4, h: 0.2, fontSize: 9, bold: true, color: a.color, margin: 0 });
  });
  footer(s, 17);
}

// ── Slide 18: 结束页 ──────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape("ellipse", { x: -2, y: -1, w: 6, h: 6, fill: { color: C.primary, transparency: 30 }, line: { width: 0 } });
  s.addShape("ellipse", { x: 6, y: 2, w: 6, h: 6, fill: { color: C.secondary, transparency: 20 }, line: { width: 0 } });
  s.addText("先让客户感受到价值", { x: 0.5, y: 1.6, w: 9, h: 0.8, fontSize: 36, bold: true, color: C.white, align: "center", margin: 0 });
  s.addText("再谈价格", { x: 0.5, y: 2.4, w: 9, h: 0.8, fontSize: 36, bold: true, color: C.secondary, align: "center", margin: 0 });
  s.addShape("rect", { x: 4, y: 3.3, w: 2, h: 0.06, fill: { color: C.accent }, line: { width: 0 } });
  s.addText("先建立信任，再自然转化", { x: 0.5, y: 3.55, w: 9, h: 0.5, fontSize: 18, color: "A8D5C2", align: "center", margin: 0 });
  s.addText("荟宝 × 美业连锁品牌增长方案", { x: 0.5, y: 4.6, w: 9, h: 0.4, fontSize: 13, color: "5E8A7C", align: "center", margin: 0 });
}

// ── Save ───────────────────────────────────────────────────────
pres.writeFile({ fileName: "/Users/flyingspur/Desktop/荟宝团单升级培训.pptx" })
  .then(() => console.log("✅ PPT已生成：/Users/flyingspur/Desktop/荟宝团单升级培训.pptx"))
  .catch(err => console.error("❌ 错误：", err));
