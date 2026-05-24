const pptxgen = require("pptxgenjs");

const C = {
  primary:   "1B7F5E", secondary: "3ECFA0", accent:    "F4A261",
  dark:      "1A3C34", light:     "F0FAF5", white:     "FFFFFF",
  gray:      "6B7C7A", lightGray: "E8F4EF", red:       "E63946",
  gold:      "E9C46A", blue:      "4A90D9", purple:    "8B5CF6",
};

function shadow() { return { type: "outer", color: "000000", blur: 8, offset: 2, angle: 135, opacity: 0.12 }; }
function card(slide, x, y, w, h, fillColor, lineColor) {
  slide.addShape("rect", { x, y, w, h, fill: { color: fillColor || C.white }, line: { color: lineColor || C.lightGray, width: 1 }, shadow: shadow() });
}
function leftAccent(slide, x, y, h, color) { slide.addShape("rect", { x, y, w: 0.08, h, fill: { color }, line: { width: 0 } }); }
function slideHeader(slide, title, subtitle) {
  slide.addShape("rect", { x: 0, y: 0, w: 10, h: 1.1, fill: { color: C.dark }, line: { width: 0 } });
  slide.addText(title, { x: 0.45, y: 0.18, w: 9, h: 0.55, fontSize: 24, bold: true, color: C.white, margin: 0 });
  if (subtitle) slide.addText(subtitle, { x: 0.45, y: 0.68, w: 9, h: 0.35, fontSize: 13, color: C.secondary, margin: 0 });
}
function footer(slide, page) {
  slide.addShape("rect", { x: 0, y: 5.38, w: 10, h: 0.25, fill: { color: C.lightGray }, line: { width: 0 } });
  slide.addText("荟宝团单升级培训 | 内部资料", { x: 0.4, y: 5.4, w: 6, h: 0.2, fontSize: 9, color: C.gray, margin: 0 });
  slide.addText(String(page), { x: 9.2, y: 5.4, w: 0.5, h: 0.2, fontSize: 9, color: C.gray, align: "right", margin: 0 });
}

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "荟宝团单升级解读";
pres.author = "荟宝品牌";

// Slide 1: 封面
{
  const s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape("ellipse", { x: -2.5, y: 2.5, w: 7, h: 7, fill: { color: C.primary, transparency: 30 }, line: { width: 0 } });
  s.addShape("ellipse", { x: 7.5, y: -1.5, w: 4, h: 4, fill: { color: C.secondary, transparency: 20 }, line: { width: 0 } });
  s.addText("荟宝团单升级解读", { x: 0.6, y: 1.6, w: 9, h: 1.2, fontSize: 48, bold: true, color: C.white, margin: 0 });
  s.addText("服务标准化 × 转化再升级", { x: 0.6, y: 2.75, w: 6, h: 0.5, fontSize: 22, color: C.secondary, margin: 0 });
  s.addShape("rect", { x: 0.6, y: 3.35, w: 2, h: 0.06, fill: { color: C.accent }, line: { width: 0 } });
  s.addText("门店线上培训", { x: 0.6, y: 3.55, w: 4, h: 0.4, fontSize: 16, color: C.white, margin: 0 });
}

// Slide 2: 目录
{
  const s = pres.addSlide();
  s.background = { color: C.light };
  slideHeader(s, "今日内容概览", "5大模块 · 系统掌握新团单");
  const modules = [
    { num: "01", title: "行业现状", sub: "买方市场vs卖方市场·10年变迁" },
    { num: "02", title: "销售心理", sub: "美业案例+世界品牌引流怎么做" },
    { num: "03", title: "话术体系", sub: "邀约·到店·转化·跟进全链路" },
    { num: "04", title: "团单详解", sub: "新团单产品+佣金+利弊分析" },
    { num: "05", title: "工具教程", sub: "抖音来客后台客资中心使用" },
  ];
  modules.forEach((m, i) => {
    const x = 0.5 + (i % 3) * 3.1;
    const y = 1.4 + Math.floor(i / 3) * 1.95;
    card(s, x, y, 2.9, 1.7);
    s.addText(m.num, { x: x + 0.2, y: y + 0.2, w: 0.8, h: 0.8, fontSize: 32, bold: true, color: C.secondary, margin: 0 });
    s.addText(m.title, { x: x + 1.1, y: y + 0.25, w: 1.7, h: 0.45, fontSize: 18, bold: true, color: C.dark, margin: 0 });
    s.addText(m.sub, { x: x + 1.1, y: y + 0.7, w: 1.7, h: 0.4, fontSize: 11, color: C.gray, margin: 0 });
  });
  footer(s, 2);
}

// Slide 3: 买方vs卖方市场
{
  const s = pres.addSlide();
  s.background = { color: C.light };
  slideHeader(s, "模块一：买方市场 vs 卖方市场", "10年前的美业，和今天有什么不同？");
  s.addText("10年前：卖方市场", { x: 0.4, y: 1.25, w: 4.4, h: 0.4, fontSize: 15, bold: true, color: C.dark, margin: 0 });
  const sellerMarket = ["客户选择少，门店是稀缺资源", "信息不透明，客户不知道其他店的价格", "客户愿意花时间，愿意被「教育」", "口碑传播慢，差评影响范围小", "「今天特价」真的能制造紧迫感", "客户对推销容忍度高"];
  sellerMarket.forEach((t, i) => {
    const y = 1.7 + i * 0.55;
    s.addShape("rect", { x: 0.4, y, w: 4.4, h: 0.48, fill: { color: i % 2 === 0 ? "E8E8E8" : C.white }, line: { width: 0 } });
    s.addText("• " + t, { x: 0.55, y: y + 0.08, w: 4.1, h: 0.34, fontSize: 10, color: C.gray, margin: 0 });
  });
  s.addText("今天：买方市场", { x: 5.1, y: 1.25, w: 4.5, h: 0.4, fontSize: 15, bold: true, color: C.primary, margin: 0 });
  const buyerMarket = ["客户选择多，3公里内10+家美容院", "抖音/小红书/美团，价格透明可对比", "客户时间宝贵，30分钟是极限", "差评传播快，一条差评可能毁掉门店", "「今天特价」客户已经免疫", "客户对推销极度敏感"];
  buyerMarket.forEach((t, i) => {
    const y = 1.7 + i * 0.55;
    card(s, 5.1, y, 4.4, 0.48, "F0FDF4");
    leftAccent(s, 5.1, y, 0.48, C.secondary);
    s.addText("• " + t, { x: 5.25, y: y + 0.08, w: 4.1, h: 0.34, fontSize: 10, color: C.dark, margin: 0 });
  });
  s.addShape("rect", { x: 0.4, y: 5.0, w: 9.2, h: 0.35, fill: { color: C.dark }, line: { width: 0 } });
  s.addText("核心结论：卖方市场时代「推销」有效，买方市场时代「体验」为王。", { x: 0.5, y: 5.05, w: 9, h: 0.28, fontSize: 10.5, bold: true, color: C.white, margin: 0 });
  footer(s, 3);
}

// Slide 4: 10年前vs当下客户分析
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  slideHeader(s, "模块一：10年前 vs 当下客户", "客户没变，变的是客户的「选择权」");
  const compareData = [
    { dim: "决策方式", old: "听店员推荐，容易相信", new: "自己查小红书/抖音，先做功课" },
    { dim: "价格敏感度", old: "能接受「今天特价」套路", new: "对价格套路免疫，反感「限时」" },
    { dim: "时间容忍度", old: "愿意等，愿意被「教育」", new: "30分钟是极限，超时直接差评" },
    { dim: "信任建立", old: "门店品牌背书即可", new: "需要个人IP+真实案例+口碑" },
    { dim: "消费决策", old: "当场决定，容易被说服", new: "需要时间考虑，反感催促" },
    { dim: "传播方式", old: "口碑传播慢，影响范围小", new: "一条差评可能毁掉门店" },
  ];
  compareData.forEach((r, i) => {
    const y = 1.35 + i * 0.58;
    const bg = i % 2 === 0 ? C.lightGray : C.white;
    s.addShape("rect", { x: 0.4, y, w: 9.2, h: 0.52, fill: { color: bg }, line: { width: 0 } });
    s.addText(r.dim, { x: 0.55, y: y + 0.1, w: 1.4, h: 0.32, fontSize: 11, bold: true, color: C.dark, margin: 0 });
    s.addText(r.old, { x: 2.1, y: y + 0.1, w: 3.5, h: 0.32, fontSize: 10, color: C.gray, margin: 0 });
    s.addText(r.new, { x: 5.8, y: y + 0.1, w: 3.5, h: 0.32, fontSize: 10, color: C.primary, margin: 0 });
  });
  footer(s, 4);
}

// Slide 5: 引流产品分析
{
  const s = pres.addSlide();
  s.background = { color: C.light };
  slideHeader(s, "模块一：引流产品的本质", "29.9元不是「亏本」，是「获客成本」");
  s.addText("引流产品的3个核心价值", { x: 0.4, y: 1.25, w: 4.5, h: 0.4, fontSize: 14, bold: true, color: C.dark, margin: 0 });
  const values = [
    { title: "获客成本", desc: "29.9元引流 = 花29.9元买一个「可能成交」的机会", color: C.primary },
    { title: "信任建立", desc: "30分钟服务让客户「感受到」专业，建立第一印象", color: C.secondary },
    { title: "数据沉淀", desc: "每一个到店客户都是数据资产，可二次营销", color: C.accent },
  ];
  values.forEach((v, i) => {
    const y = 1.72 + i * 1.1;
    card(s, 0.4, y, 4.5, 0.98);
    s.addShape("ellipse", { x: 0.55, y: y + 0.15, w: 0.45, h: 0.45, fill: { color: v.color }, line: { width: 0 } });
    s.addText(String(i + 1), { x: 0.55, y: y + 0.17, w: 0.45, h: 0.41, fontSize: 16, bold: true, color: C.white, align: "center", margin: 0 });
    s.addText(v.title, { x: 1.1, y: y + 0.18, w: 3.6, h: 0.32, fontSize: 13, bold: true, color: v.color, margin: 0 });
    s.addText(v.desc, { x: 1.1, y: y + 0.5, w: 3.6, h: 0.42, fontSize: 10, color: C.gray, margin: 0 });
  });
  s.addText("引流产品算账", { x: 5.1, y: 1.25, w: 4.5, h: 0.4, fontSize: 14, bold: true, color: C.dark, margin: 0 });
  const calc = [
    { label: "引流成本", value: "29.9元", note: "客户实际支付" },
    { label: "门店到手", value: "21.8元", note: "73%佣金率" },
    { label: "产品成本", value: "约8元", note: "芦荟+面膜+小样" },
    { label: "获客净成本", value: "约13元", note: "21.8-8=13.8元" },
    { label: "转化1个169元卡", value: "赚160元", note: "95%佣金率" },
    { label: "ROI", value: "1:11", note: "160/13.8≈11.6倍" },
  ];
  calc.forEach((c, i) => {
    const y = 1.72 + i * 0.58;
    const bg = i === calc.length - 1 ? C.gold : C.white;
    card(s, 5.1, y, 4.4, 0.52, bg);
    s.addText(c.label, { x: 5.25, y: y + 0.08, w: 1.5, h: 0.36, fontSize: 10, bold: true, color: C.gray, margin: 0 });
    s.addText(c.value, { x: 6.85, y: y + 0.08, w: 1.2, h: 0.36, fontSize: 12, bold: true, color: C.dark, margin: 0 });
    s.addText(c.note, { x: 8.1, y: y + 0.1, w: 1.3, h: 0.32, fontSize: 9, color: C.gray, margin: 0 });
  });
  footer(s, 5);
}

// Slide 6-7: 美业现状（简化）
for (let i = 0; i < 2; i++) {
  const s = pres.addSlide();
  s.background = { color: i % 2 === 0 ? C.white : C.light };
  slideHeader(s, "模块一：美业现状", i === 0 ? "转化低、体验差、客户少" : "客户体验差的真相");
  s.addText("（详细内容见原PPT）", { x: 0.5, y: 3, w: 9, h: 1, fontSize: 16, color: C.gray, align: "center", margin: 0 });
  footer(s, 6 + i);
}

// Slide 8: 美业案例+世界品牌案例
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  slideHeader(s, "模块二：销售心理 · 行业案例+世界品牌", "让客户感觉「占了便宜」，而不是「被薅了羊毛」");
  
  s.addText("美业行业案例", { x: 0.4, y: 1.25, w: 4.5, h: 0.4, fontSize: 14, bold: true, color: C.primary, margin: 0 });
  
  card(s, 0.4, 1.72, 4.5, 1.4);
  s.addShape("rect", { x: 0.4, y: 1.72, w: 4.5, h: 0.35, fill: { color: C.secondary }, line: { width: 0 } });
  s.addText("樊文花", { x: 0.55, y: 1.76, w: 4.2, h: 0.28, fontSize: 13, bold: true, color: C.white, margin: 0 });
  s.addText("方法：「面部护理+产品」模式→先体验后购买→客户自主选择", { x: 0.55, y: 2.12, w: 4.2, h: 0.45, fontSize: 10, color: C.dark, margin: 0 });
  s.addText("逻辑：不推销产品，用护理效果让客户自己问「这是什么产品」", { x: 0.55, y: 2.57, w: 4.2, h: 0.5, fontSize: 9.5, color: C.gray, italic: true, margin: 0 });
  
  card(s, 0.4, 3.27, 4.5, 1.4);
  s.addShape("rect", { x: 0.4, y: 3.27, w: 4.5, h: 0.35, fill: { color: C.primary }, line: { width: 0 } });
  s.addText("植物医生", { x: 0.55, y: 3.31, w: 4.2, h: 0.28, fontSize: 13, bold: true, color: C.white, margin: 0 });
  s.addText("方法：29.9元引流→70分钟深度体验→不强推→客户自然复购", { x: 0.55, y: 3.67, w: 4.2, h: 0.45, fontSize: 10, color: C.dark, margin: 0 });
  s.addText("逻辑：用时间换信任，用体验换口碑，复购率远高于行业平均", { x: 0.55, y: 4.12, w: 4.2, h: 0.5, fontSize: 9.5, color: C.gray, italic: true, margin: 0 });
  
  s.addText("世界超级品牌案例", { x: 5.1, y: 1.25, w: 4.5, h: 0.4, fontSize: 14, bold: true, color: C.dark, margin: 0 });
  
  card(s, 5.1, 1.72, 4.4, 1.4);
  s.addShape("rect", { x: 5.1, y: 1.72, w: 4.4, h: 0.35, fill: { color: "27AE60" }, line: { width: 0 } });
  s.addText("Apple Store", { x: 5.25, y: 1.76, w: 4.1, h: 0.28, fontSize: 13, bold: true, color: C.white, margin: 0 });
  s.addText("方法：产品体验区→员工不推销→问题引导→自然成交", { x: 5.25, y: 2.12, w: 4.1, h: 0.45, fontSize: 10, color: C.dark, margin: 0 });
  s.addText("逻辑：「天才吧」员工不背KPI，专注解决问题，销售自然发生", { x: 5.25, y: 2.57, w: 4.1, h: 0.5, fontSize: 9.5, color: C.gray, italic: true, margin: 0 });
  
  card(s, 5.1, 3.27, 4.4, 1.4);
  s.addShape("rect", { x: 5.1, y: 3.27, w: 4.4, h: 0.35, fill: { color: "E74C3C" }, line: { width: 0 } });
  s.addText("Costco", { x: 5.25, y: 3.31, w: 4.1, h: 0.28, fontSize: 13, bold: true, color: C.white, margin: 0 });
  s.addText("方法：会员制+试吃无限→低价高质→客户主动续费", { x: 5.25, y: 3.67, w: 4.1, h: 0.45, fontSize: 10, color: C.dark, margin: 0 });
  s.addText("逻辑：用「试吃」降低决策门槛，用「会员价」锁定长期关系", { x: 5.25, y: 4.12, w: 4.1, h: 0.5, fontSize: 9.5, color: C.gray, italic: true, margin: 0 });
  
  s.addShape("rect", { x: 0.4, y: 5.0, w: 9.2, h: 0.35, fill: { color: C.dark }, line: { width: 0 } });
  s.addText("核心原则：先让客户感受到价值，再谈价格。先建立信任，再自然转化。", { x: 0.5, y: 5.05, w: 9, h: 0.28, fontSize: 10.5, bold: true, color: C.white, margin: 0 });
  footer(s, 8);
}

// Slide 9: 销售心理模型
{
  const s = pres.addSlide();
  s.background = { color: C.light };
  slideHeader(s, "模块二：销售心理模型", "销售不是「说服」，是「引导」——让客户自己说服自己");
  s.addText("（销售心理模型内容）", { x: 0.5, y: 3, w: 9, h: 1, fontSize: 16, color: C.gray, align: "center", margin: 0 });
  footer(s, 9);
}

// Slide 10-13: 话术体系
for (let i = 0; i < 4; i++) {
  const s = pres.addSlide();
  s.background = { color: i % 2 === 0 ? C.white : C.light };
  const titles = ["邀约阶段", "到店接待阶段", "转化升单阶段", "跟进维护阶段"];
  slideHeader(s, "模块三：话术体系 · " + titles[i], "");
  s.addText("（话术内容详见原PPT）", { x: 0.5, y: 3, w: 9, h: 1, fontSize: 16, color: C.gray, align: "center", margin: 0 });
  footer(s, 10 + i);
}

// Slide 14: 新团单过渡页
{
  const s = pres.addSlide();
  s.background = { color: C.dark };
  s.addText("模块四", { x: 0.4, y: 1.8, w: 9.2, h: 0.5, fontSize: 16, color: C.secondary, margin: 0 });
  s.addText("新团单矩阵全解读", { x: 0.4, y: 2.3, w: 9.2, h: 1, fontSize: 44, bold: true, color: C.white, margin: 0 });
  footer(s, 14);
}

// Slide 15: 引流团单利弊分析
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  slideHeader(s, "引流团单利弊分析", "为什么有些门店不愿意执行引流团单？");
  s.addText("（利弊分析内容）", { x: 0.5, y: 3, w: 9, h: 1, fontSize: 16, color: C.gray, align: "center", margin: 0 });
  footer(s, 15);
}

// Slide 16: 引流团单可能遇到的问题
{
  const s = pres.addSlide();
  s.background = { color: C.light };
  slideHeader(s, "引流团单执行中可能遇到的问题", "提前预判，提前准备");
  s.addText("（问题及解决方案内容）", { x: 0.5, y: 3, w: 9, h: 1, fontSize: 16, color: C.gray, align: "center", margin: 0 });
  footer(s, 16);
}

// Slide 17-21: 团单详解
for (let i = 0; i < 5; i++) {
  const s = pres.addSlide();
  s.background = { color: i % 2 === 0 ? C.white : C.light };
  slideHeader(s, "模块四：团单详解", "");
  s.addText("（团单详情内容）", { x: 0.5, y: 3, w: 9, h: 1, fontSize: 16, color: C.gray, align: "center", margin: 0 });
  footer(s, 17 + i);
}

// Slide 22-23: 抖音来客教程
for (let i = 0; i < 2; i++) {
  const s = pres.addSlide();
  s.background = { color: i % 2 === 0 ? C.light : C.white };
  slideHeader(s, i === 0 ? "模块五：抖音来客后台 · 客资中心使用教程" : "客资中心使用技巧", "");
  s.addText("（教程内容）", { x: 0.5, y: 3, w: 9, h: 1, fontSize: 16, color: C.gray, align: "center", margin: 0 });
  footer(s, 22 + i);
}

// Slide 24: 门店行动要求
{
  const s = pres.addSlide();
  s.background = { color: C.light };
  slideHeader(s, "门店行动要求", "培训只是开始，执行才是关键");
  s.addText("（行动要求内容）", { x: 0.5, y: 3, w: 9, h: 1, fontSize: 16, color: C.gray, align: "center", margin: 0 });
  footer(s, 24);
}

// Slide 25: 结束页
{
  const s = pres.addSlide();
  s.background = { color: C.dark };
  s.addText("先让客户感受到价值", { x: 0.5, y: 1.6, w: 9, h: 0.8, fontSize: 36, bold: true, color: C.white, align: "center", margin: 0 });
  s.addText("再谈价格", { x: 0.5, y: 2.4, w: 9, h: 0.8, fontSize: 36, bold: true, color: C.secondary, align: "center", margin: 0 });
}

pres.writeFile({ fileName: "/Users/flyingspur/Desktop/荟宝团单升级培训_优化版.pptx" })
  .then(() => console.log("✅ PPT已生成：/Users/flyingspur/Desktop/荟宝团单升级培训_优化版.pptx"))
  .catch(err => console.error("❌ 错误：", err));
