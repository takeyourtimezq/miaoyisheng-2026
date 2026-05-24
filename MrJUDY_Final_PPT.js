const pptxgen = require("pptxgenjs");
let pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';
pres.author = '连锁帝国CGO';
pres.title = 'Mr.JUDY 抖音全案运营方案';

const C = { primary: "1E3A5F", secondary: "2E5A8C", accent: "4A90D9", light: "E8F4FC", white: "FFFFFF", dark: "1A1A2E", gray: "64748B", green: "10B981", purple: "7C3AED" };

function title(slide, t, st) {
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 1.1, fill: { color: C.primary } });
  slide.addText(t, { x: 0.5, y: 0.25, w: 9, h: 0.5, fontSize: 26, fontFace: "Arial Black", color: C.white, bold: true });
  if (st) slide.addText(st, { x: 0.5, y: 0.72, w: 9, h: 0.3, fontSize: 13, fontFace: "Arial", color: C.accent });
}

// 1-2: 封面+目录
let p1 = pres.addSlide(); p1.background = { color: C.primary };
p1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.15, fill: { color: C.accent } });
p1.addText("Mr.JUDY", { x: 0.5, y: 1.8, w: 9, h: 1, fontSize: 54, fontFace: "Arial Black", color: C.white, bold: true, align: "center" });
p1.addText("抖音全案运营方案", { x: 0.5, y: 2.8, w: 9, h: 0.8, fontSize: 36, fontFace: "Arial", color: C.accent, bold: true, align: "center" });
p1.addText("参考TCJF云连锁打法，定制洗发养发品牌抖音线上运营策略", { x: 1, y: 4, w: 8, h: 0.6, fontSize: 16, fontFace: "Arial", color: C.white, align: "center" });
p1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.4, w: 10, h: 0.225, fill: { color: C.accent } });
p1.addText("2026年5月", { x: 0.5, y: 4.8, w: 9, h: 0.4, fontSize: 14, fontFace: "Arial", color: C.white, align: "center" });

let p2 = pres.addSlide(); title(p2, "方案目录", null);
let y2 = 1.5;
["01  TCJF云连锁打法深度拆解（3页）", "02  Mr.JUDY品牌分析与差异化定位（3页）", "03  整体策略与内容规划（3页）", "04  团购套餐与转化设计（3页）", "05  达人推广体系（3页）", "06  商家矩阵账号体系（3页）", "07  执行落地与阶段目标（3页）"].forEach((it, i) => {
  p2.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: y2, w: 0.5, h: 0.5, fill: { color: C.accent } });
  p2.addText(String(i + 1).padStart(2, '0'), { x: 0.8, y: y2, w: 0.5, h: 0.5, fontSize: 18, fontFace: "Arial", color: C.white, bold: true, align: "center", valign: "middle" });
  p2.addText(it.substring(4), { x: 1.5, y: y2, w: 7.5, h: 0.5, fontSize: 17, fontFace: "Arial", color: C.dark, valign: "middle" });
  y2 += 0.6;
});

// 3-5: TCJF
let p3 = pres.addSlide(); title(p3, "01  TCJF云连锁打法深度拆解", "（1/3）模式概述与核心数据");
let x3 = 0.5;
[{ n: "2700+", l: "全国门店", d: "云连锁串联不同品牌/法人" }, { n: "29.9", l: "引流款起价", d: "洗吹造型三次卡获客" }, { n: "388", l: "利润款均价", d: "总监烫染套餐盈利" }, { n: "15天", l: "售后承诺", d: "不满意免费重做" }].forEach(s => {
  p3.addShape(pres.shapes.RECTANGLE, { x: x3, y: 1.4, w: 2.2, h: 1.6, fill: { color: C.light } });
  p3.addText(s.n, { x: x3, y: 1.55, w: 2.2, h: 0.5, fontSize: 32, fontFace: "Arial Black", color: C.accent, bold: true, align: "center" });
  p3.addText(s.l, { x: x3, y: 2.15, w: 2.2, h: 0.3, fontSize: 14, fontFace: "Arial", color: C.primary, bold: true, align: "center" });
  p3.addText(s.d, { x: x3 + 0.1, y: 2.5, w: 2.0, h: 0.4, fontSize: 11, fontFace: "Arial", color: C.gray, align: "center" });
  x3 += 2.4;
});
p3.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.3, w: 9, h: 2.0, fill: { color: C.white } });
p3.addText("云连锁模式定义", { x: 0.7, y: 3.45, w: 8.6, h: 0.4, fontSize: 18, fontFace: "Arial", color: C.primary, bold: true });
p3.addText("由品牌方/服务商牵头，将线下同类目可标准化产品的门店认领到一起，实现不同品牌、不同营业执照、不同法人、不同地区的门店通过主品牌串联，共同享受品牌流量辐射，形成品牌矩阵，集中火力打造中央直播间，增加团购套餐核销量。", { x: 0.7, y: 3.95, w: 8.6, h: 1.2, fontSize: 14, fontFace: "Arial", color: C.dark });

let p4 = pres.addSlide(); title(p4, "01  TCJF云连锁打法深度拆解", "（2/3）内容策略与流量打法");
p4.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.3, w: 4.3, h: 2.2, fill: { color: C.light } });
p4.addText("内容策略", { x: 0.7, y: 1.45, w: 4, h: 0.4, fontSize: 18, fontFace: "Arial", color: C.primary, bold: true });
["真人出镜：发型师/总监真人IP", "服务过程：烫染过程全程记录", "前后对比：发型改造强烈对比", "场景植入：五一出游、约会前", "痛点切入：15天不满意重做"].forEach((t, i) => { p4.addText("• " + t, { x: 0.7, y: 1.95 + i * 0.38, w: 3.9, h: 0.35, fontSize: 13, fontFace: "Arial", color: C.dark }); });
p4.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.3, w: 4.3, h: 2.2, fill: { color: C.light } });
p4.addText("流量来源矩阵", { x: 5.4, y: 1.45, w: 4, h: 0.4, fontSize: 18, fontFace: "Arial", color: C.primary, bold: true });
["短视频矩阵：主号+门店+员工", "达人探店：本地生活达人合作", "中央直播间：总部统一带货", "付费投流：DOU+精准投放", "自然流量：爆款二次传播"].forEach((t, i) => { p4.addText("• " + t, { x: 5.4, y: 1.95 + i * 0.38, w: 3.9, h: 0.35, fontSize: 13, fontFace: "Arial", color: C.dark }); });
p4.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.8, w: 9, h: 1.4, fill: { color: C.secondary } });
p4.addText("爆款内容公式", { x: 0.7, y: 3.9, w: 8.6, h: 0.4, fontSize: 16, fontFace: "Arial", color: C.white, bold: true });
p4.addText("3秒钩子（痛点/好奇）+ 30秒过程（服务展示）+ 10秒转化（团购引导+售后承诺）", { x: 0.7, y: 4.35, w: 8.6, h: 0.4, fontSize: 15, fontFace: "Arial", color: C.white, align: "center" });

let p5 = pres.addSlide(); title(p5, "01  TCJF云连锁打法深度拆解", "（3/3）团购策略与成功要素");
p5.addText("阶梯式团购设计", { x: 0.5, y: 1.25, w: 9, h: 0.4, fontSize: 18, fontFace: "Arial", color: C.primary, bold: true });
let x5 = 0.8;
[{ t: "引流款", p: "29.9-59", c: "洗吹造型三次卡", g: "获客到店", cl: C.green }, { t: "利润款", p: "168-388", c: "烫染/总监设计", g: "主力盈利", cl: C.accent }, { t: "形象款", p: "588-888", c: "高端定制/明星款", g: "品牌调性", cl: C.secondary }].forEach(p => {
  p5.addShape(pres.shapes.RECTANGLE, { x: x5, y: 1.75, w: 2.8, h: 2.0, fill: { color: C.white } });
  p5.addShape(pres.shapes.RECTANGLE, { x: x5, y: 1.75, w: 2.8, h: 0.12, fill: { color: p.cl } });
  p5.addText(p.t, { x: x5, y: 2.0, w: 2.8, h: 0.4, fontSize: 18, fontFace: "Arial Black", color: p.cl, bold: true, align: "center" });
  p5.addText("¥" + p.p, { x: x5, y: 2.5, w: 2.8, h: 0.4, fontSize: 24, fontFace: "Arial Black", color: C.dark, bold: true, align: "center" });
  p5.addText(p.c, { x: x5, y: 3.0, w: 2.8, h: 0.3, fontSize: 12, fontFace: "Arial", color: C.gray, align: "center" });
  p5.addShape(pres.shapes.RECTANGLE, { x: x5 + 0.5, y: 3.4, w: 1.8, h: 0.35, fill: { color: p.cl } });
  p5.addText(p.g, { x: x5 + 0.5, y: 3.43, w: 1.8, h: 0.3, fontSize: 12, fontFace: "Arial", color: C.white, bold: true, align: "center", valign: "middle" });
  x5 += 3.0;
});
p5.addText("关键成功要素", { x: 0.5, y: 4.0, w: 9, h: 0.4, fontSize: 18, fontFace: "Arial", color: C.primary, bold: true });
["标准化服务流程 — 全流程SOP确保一致性", "强售后承诺 — 15天不满意免费重做", "云连锁模式 — 轻资产扩张共享流量", "阶梯式团购 — 引流获客+利润赚钱+形象立调"].forEach((f, i) => { p5.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.45 + i * 0.38, w: 0.12, h: 0.12, fill: { color: C.accent } }); p5.addText(f, { x: 0.75, y: 4.4 + i * 0.38, w: 8.5, h: 0.3, fontSize: 13, fontFace: "Arial", color: C.dark }); });

// 6-8: 品牌分析
let p6 = pres.addSlide(); title(p6, "02  Mr.JUDY品牌分析与差异化定位", "（1/3）品牌基础与核心优势");
p6.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.3, w: 4.3, h: 3.8, fill: { color: C.light } });
p6.addText("品牌基础信息", { x: 0.7, y: 1.45, w: 4, h: 0.4, fontSize: 18, fontFace: "Arial", color: C.primary, bold: true });
[["品牌名称", "Mr.JUDY 洗个头发"], ["所属公司", "北京小悦科技"], ["成立时间", "2018年"], ["目标人群", "一线/新一线白领女性"], ["核心服务", "洗发、头皮护理、按摩"], ["服务模式", "沉浸式独立空间"], ["行业地位", "养发馆十大品牌"]].forEach(([k, v], i) => { p6.addText(k + "：", { x: 0.7, y: 1.95 + i * 0.48, w: 1.4, h: 0.4, fontSize: 12, fontFace: "Arial", color: C.primary, bold: true }); p6.addText(v, { x: 2.1, y: 1.95 + i * 0.48, w: 2.5, h: 0.4, fontSize: 12, fontFace: "Arial", color: C.dark }); });
p6.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.3, w: 4.3, h: 3.8, fill: { color: C.white } });
p6.addText("核心差异化优势", { x: 5.4, y: 1.45, w: 4, h: 0.4, fontSize: 18, fontFace: "Arial", color: C.primary, bold: true });
[["卡诗官方合作", "使用卡诗专业产品线"], ["专利洗护按摩椅", "多功能专利设备"], ["独立私密空间", "白领女性专属放松空间"], ["标准化流程", "头皮检测到护理全流程"], ["专业定位", "垂直洗发养发专家"]].forEach(([t, d], i) => { p6.addText("★", { x: 5.4, y: 1.95 + i * 0.7, w: 0.3, h: 0.35, fontSize: 14, fontFace: "Arial", color: C.accent, bold: true }); p6.addText(t, { x: 5.7, y: 1.95 + i * 0.7, w: 3.6, h: 0.35, fontSize: 13, fontFace: "Arial", color: C.dark, bold: true }); p6.addText(d, { x: 5.7, y: 2.28 + i * 0.7, w: 3.6, h: 0.3, fontSize: 11, fontFace: "Arial", color: C.gray }); });

let p7 = pres.addSlide(); title(p7, "02  Mr.JUDY品牌分析与差异化定位", "（2/3）与TCJF的深度对比分析");
let x7 = 0.5; const w7 = [1.2, 2.2, 2.2, 3.9];
["维度", "TCJF（理发）", "Mr.JUDY（养发）", "策略启示"].forEach((h, i) => { p7.addShape(pres.shapes.RECTANGLE, { x: x7, y: 1.3, w: w7[i], h: 0.5, fill: { color: C.secondary } }); p7.addText(h, { x: x7, y: 1.38, w: w7[i], h: 0.4, fontSize: 12, fontFace: "Arial", color: C.white, bold: true, align: "center", valign: "middle" }); x7 += w7[i]; });
let y7 = 1.85;
[["客单价", "中低29.9-388", "中高80-300", "拼体验价值非低价"], ["消费频次", "高频1-2月", "中频2-4周", "更高复购潜力"], ["决策门槛", "低（刚需）", "中（非刚需）", "需教育市场"], ["核心痛点", "发型效果", "头皮健康+放松", "双重价值主张"], ["目标人群", "大众消费者", "白领女性", "精准高LTV人群"], ["服务场景", "公开快节奏", "独立慢享受", "体验是壁垒"], ["竞争壁垒", "价格便利", "品质专业", "建立专家形象"]].forEach((r, i) => { x7 = 0.5; r.forEach((c, j) => { p7.addShape(pres.shapes.RECTANGLE, { x: x7, y: y7, w: w7[j], h: 0.48, fill: { color: i % 2 === 0 ? C.light : C.white } }); p7.addText(c, { x: x7 + 0.05, y: y7 + 0.12, w: w7[j] - 0.1, h: 0.3, fontSize: 10, fontFace: "Arial", color: j === 0 ? C.primary : (j === 3 ? C.accent : C.dark), bold: j === 0 || j === 3, align: j === 0 ? "left" : "center" }); x7 += w7[j]; }); y7 += 0.52; });

let p8 = pres.addSlide(); title(p8, "02  Mr.JUDY品牌分析与差异化定位", "（3/3）差异化定位与核心策略");
p8.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.3, w: 9, h: 1.0, fill: { color: C.accent } });
p8.addText("核心定位：精致护理专家", { x: 0.5, y: 1.55, w: 9, h: 0.5, fontSize: 32, fontFace: "Arial Black", color: C.white, bold: true, align: "center" });
p8.addText("三大策略支柱", { x: 0.5, y: 2.5, w: 9, h: 0.4, fontSize: 18, fontFace: "Arial", color: C.primary, bold: true });
let x8 = 0.5;
[{ n: "01", t: "体验差异化", s: "沉浸式独立空间", d: "卡诗品质+专利设备+白领专属空间", cl: C.accent }, { n: "02", t: "内容专业化", s: "头皮健康知识", d: "头皮检测到方案定制，建立专家形象", cl: C.secondary }, { n: "03", t: "转化阶梯化", s: "体验价值驱动", d: "低价引流→到店升级→储值锁客", cl: C.green }].forEach(p => { p8.addShape(pres.shapes.RECTANGLE, { x: x8, y: 3.0, w: 3, h: 2.4, fill: { color: C.light } }); p8.addShape(pres.shapes.RECTANGLE, { x: x8, y: 3.0, w: 3, h: 0.6, fill: { color: p.cl } }); p8.addText(p.n, { x: x8, y: 3.1, w: 3, h: 0.4, fontSize: 24, fontFace: "Arial Black", color: C.white, bold: true, align: "center" }); p8.addText(p.t, { x: x8, y: 3.7, w: 3, h: 0.4, fontSize: 18, fontFace: "Arial Black", color: p.cl, bold: true, align: "center" }); p8.addText(p.s, { x: x8, y: 4.15, w: 3, h: 0.3, fontSize: 13, fontFace: "Arial", color: C.dark, align: "center" }); p8.addText(p.d, { x: x8 + 0.15, y: 4.55, w: 2.7, h: 0.7, fontSize: 11, fontFace: "Arial", color: C.gray, align: "center" }); x8 += 3.2; });

// 9-11: 内容规划
let p9 = pres.addSlide(); title(p9, "03  整体策略与内容规划", "（1/3）内容策略总览与类型配比");
p9.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.3, w: 9, h: 0.7, fill: { color: C.secondary } });
p9.addText("内容公式：痛点共鸣 + 场景代入 + 效果可视化", { x: 0.5, y: 1.5, w: 9, h: 0.4, fontSize: 20, fontFace: "Arial", color: C.white, bold: true, align: "center" });
let x9 = 0.5;
[{ t: "知识科普类", p: "30%", i: ["头皮健康科普", "脱发原因解析", "养发误区纠正", "季节护理指南"], cl: C.accent }, { t: "服务展示类", p: "35%", i: ["沉浸式洗头过程", "头皮检测展示", "护理前后对比", "ASMR音效体验"], cl: C.secondary }, { t: "用户证言类", p: "20%", i: ["真实客户体验", "Vlog记录", "效果反馈", "复购故事"], cl: C.green }, { t: "品牌调性类", p: "15%", i: ["卡诗合作展示", "专利设备介绍", "独立空间环境", "白领生活方式"], cl: C.purple }].forEach(ct => { p9.addShape(pres.shapes.RECTANGLE, { x: x9, y: 2.2, w: 2.2, h: 3.1, fill: { color: C.white } }); p9.addShape(pres.shapes.RECTANGLE, { x: x9, y: 2.2, w: 2.2, h: 0.7, fill: { color: ct.cl } }); p9.addText(ct.p, { x: x9, y: 2.35, w: 2.2, h: 0.4, fontSize: 28, fontFace: "Arial Black", color: C.white, bold: true, align: "center" }); p9.addText(ct.t, { x: x9, y: 3.0, w: 2.2, h: 0.35, fontSize: 14, fontFace: "Arial", color: ct.cl, bold: true, align: "center" }); ct.i.forEach((it, j) => { p9.addText("• " + it, { x: x9 + 0.15, y: 3.45 + j * 0.38, w: 1.9, h: 0.35, fontSize: 11, fontFace: "Arial", color: C.dark }); }); x9 += 2.4; });

let p10 = pres.addSlide(); title(p10, "03  整体策略与内容规划", "（2/3）爆款内容模板与脚本结构");
let x10 = 0.5;
[{ t: "3秒钩子", g: "抓住注意力", cl: C.accent, items: ["\"头发越洗越油？\"", "\"头皮比脸皮老6岁\"", "\"白领的秘密基地\"", "\"熬夜掉发看过来\"", "\"39.9元贵妇体验\""] }, { t: "30秒过程", g: "建立信任", cl: C.secondary, items: ["头皮检测特写", "ASMR洗头音效", "按摩手法展示", "客户放松表情", "卡诗产品展示"] }, { t: "10秒转化", g: "引导行动", cl: C.green, items: ["\"洗完轻了2斤\"", "\"点击左下角39.9\"", "\"每天限10名额\"", "\"限时手慢无\"", "\"评论区扣1预约\""] }].forEach(sec => { p10.addShape(pres.shapes.RECTANGLE, { x: x10, y: 1.3, w: 2.9, h: 3.8, fill: { color: C.light } }); p10.addShape(pres.shapes.RECTANGLE, { x: x10, y: 1.3, w: 2.9, h: 0.5, fill: { color: sec.cl } }); p10.addText(sec.t, { x: x10, y: 1.38, w: 2.9, h: 0.4, fontSize: 18, fontFace: "Arial Black", color: C.white, bold: true, align: "center", valign: "middle" }); p10.addText("目标：" + sec.g, { x: x10, y: 1.9, w: 2.9, h: 0.3, fontSize: 12, fontFace: "Arial", color: C.gray, align: "center" }); sec.items.forEach((it, i) => { p10.addText(it, { x: x10 + 0.2, y: 2.3 + i * 0.55, w: 2.5, h: 0.5, fontSize: 12, fontFace: "Arial", color: C.dark }); }); x10 += 3.15; });

let p11 = pres.addSlide(); title(p11, "03  整体策略与内容规划", "（3/3）内容日历与发布节奏");
p11.addText("商家门店号内容日历（2-3条/天）", { x: 0.5, y: 1.25, w: 4.5, h: 0.35, fontSize: 14, fontFace: "Arial", color: C.primary, bold: true });
let x11 = 0.5;
["星期", "上午", "下午", "晚间"].forEach((h, i) => { const w = i === 0 ? 0.9 : 1.2; p11.addShape(pres.shapes.RECTANGLE, { x: x11, y: 1.7, w: w, h: 0.4, fill: { color: C.secondary } }); p11.addText(h, { x: x11, y: 1.75, w: w, h: 0.3, fontSize: 11, fontFace: "Arial", color: C.white, bold: true, align: "center", valign: "middle" }); x11 += w; });
let y11 = 2.15;
[["周一", "门店环境", "服务快剪", "好评截图"], ["周二", "知识科普", "员工日常", "福利预告"], ["周三", "达人转发", "检测过程", "预约提醒"], ["周四", "产品展示", "ASMR洗头", "活动预告"], ["周五", "周末福利", "客户Vlog", "预约提醒"], ["周六", "高峰实拍", "排队场景", "当日战报"], ["周日", "周总结", "UGC合集", "互动问答"]].forEach((r, i) => { x11 = 0.5; r.forEach((c, j) => { const w = j === 0 ? 0.9 : 1.2; p11.addShape(pres.shapes.RECTANGLE, { x: x11, y: y11, w: w, h: 0.38, fill: { color: i % 2 === 0 ? C.light : C.white } }); p11.addText(c, { x: x11, y: y11 + 0.08, w: w, h: 0.25, fontSize: 9, fontFace: "Arial", color: j === 0 ? C.primary : C.dark, bold: j === 0, align: "center", valign: "middle" }); x11 += w; }); y11 += 0.42; });
p11.addText("必拍8个标准镜头", { x: 5.5, y: 1.25, w: 4, h: 0.35, fontSize: 14, fontFace: "Arial", color: C.primary, bold: true });
["门头招牌（3秒）- 品牌露出", "接待区（5秒）- 环境展示", "独立护理间（5秒）- 私密空间", "头皮检测（10秒）- 专业感", "洗护过程（15秒）- ASMR", "放松表情（5秒）- 效果感", "前后对比（5秒）- 可视化", "团购引导（5秒）- 转化点"].forEach((s, i) => { p11.addShape(pres.shapes.RECTANGLE, { x: 5.5, y: 1.7 + i * 0.5, w: 0.25, h: 0.25, fill: { color: C.accent } }); p11.addText(String(i + 1), { x: 5.5, y: 1.7 + i * 0.5, w: 0.25, h: 0.25, fontSize: 10, fontFace: "Arial", color: C.white, bold: true, align: "center", valign: "middle" }); p11.addText(s, { x: 5.85, y: 1.68 + i * 0.5, w: 3.5, h: 0.3, fontSize: 11, fontFace: "Arial", color: C.dark }); });

// 由于代码长度限制，这里只生成前11页，剩余页面需要继续添加
// 保存当前进度
pres.writeFile({ fileName: "/Users/flyingspur/.qclaw/workspace-agent-6f07984c/MrJUDY_抖音全案运营方案_24页.pptx" });
console.log("PPT已生成：MrJUDY_抖音全案运营方案_24页.pptx（前11页）");
console.log("继续生成剩余13页...");
