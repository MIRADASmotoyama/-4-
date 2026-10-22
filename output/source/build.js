// 第1回 ミラキン業界研究会「金属業界」全41枚 PowerPoint 生成スクリプト
//   cd output/source && npm install && node render_assets.js && node build.js
const path = require("path");
const fs = require("fs");
const pptxgen = require("pptxgenjs");
const sharp = require("sharp");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const Tb = require("react-icons/tb");
const { applyTheme } = require("./apply_theme.js");
const { S, CHAPTERS, chapterOf, notesFor } = require("./content");

const OUT = path.resolve(__dirname, "..", "MIRAKIN_金属業界研究会_全41枚.pptx");
const ASSETS = path.join(__dirname, "assets");
const CROPS = path.join(ASSETS, "crops");
fs.mkdirSync(CROPS, { recursive: true });

// ---- テーマ -------------------------------------------------------------
const FONT = "Noto Sans CJK JP";
const HEX = {
  navy: "12243A", teal: "067F82", cyan: "20A8BC", pale: "EDF5F7", orange: "EEA348",
  text: "26384A", gray: "5B6B7C", line: "C9DDE2", red: "C8483D", white: "FFFFFF",
  navy2: "1C3A57", orangePale: "FCEBD3", tealPale: "D6ECEC", redPale: "F6DEDB",
};
const THEME = {
  name: "MIRAKIN Metal",
  headFontFace: FONT,
  bodyFontFace: FONT,
  colors: {
    dk1: HEX.text, lt1: HEX.white, dk2: HEX.navy, lt2: HEX.pale,
    accent1: HEX.teal, accent2: HEX.cyan, accent3: HEX.orange, accent4: HEX.gray,
    accent5: HEX.line, accent6: HEX.red, hlink: HEX.teal, folHlink: HEX.navy2,
  },
};

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5 inch
pres.theme = { headFontFace: FONT, bodyFontFace: FONT };
pres.title = "第1回 ミラキン業界研究会「金属業界」";
pres.author = "ミラキン業界研究会";
const C = pres.SchemeColor;
const SH = pres.shapes;
// 淡色の塗りはテーマ外の補助色（16進）で指定する
const FILL = {
  navy: C.text2, teal: C.accent1, cyan: C.accent2, orange: C.accent3, gray: C.accent4,
  line: C.accent5, red: C.accent6, pale: C.background2, white: C.background1, text: C.text1,
};

// ---- レイアウト（スライドマスター） ---------------------------------------
const FOOT = "第1回 ミラキン業界研究会｜金属業界";
function footerObjs(dark, w = 8) {
  return [{ text: { text: FOOT, options: { x: 0.6, y: 7.0, w, h: 0.3, fontSize: 11, color: dark ? "A9C4CF" : "6E7F8F", fontFace: FONT, margin: 0 } } }];
}
const chapPh = (color, x = 0.6, y = 0.38, w = 10) => ({
  placeholder: { options: { name: "chap", type: "body", x, y, w, h: 0.32, fontSize: 14, bold: true, color, margin: 0, valign: "middle" }, text: "" },
});
const titlePh = (color, o = {}) => ({
  placeholder: { options: Object.assign({ name: "title", type: "title", x: 0.6, y: 0.7, w: 12.13, h: 0.82, fontSize: 30, bold: true, color, margin: 0, valign: "middle", align: "left" }, o), text: "" },
});
const slideNo = (color, x = 12.13) => ({ x, y: 6.98, w: 0.6, h: 0.32, fontSize: 12, color, fontFace: FONT, align: "right" });

pres.defineSlideMaster({ title: "LIGHT", background: { color: C.background1 },
  objects: [chapPh(C.accent1), titlePh(C.text2), ...footerObjs(false)], slideNumber: slideNo("6E7F8F") });
pres.defineSlideMaster({ title: "WORK", background: { color: C.background2 },
  objects: [chapPh(C.accent1), titlePh(C.text2), ...footerObjs(false)], slideNumber: slideNo("6E7F8F") });
pres.defineSlideMaster({ title: "DARK", background: { color: C.text2 },
  objects: [chapPh(C.accent2), titlePh(C.background1), ...footerObjs(true)], slideNumber: slideNo("A9C4CF") });
pres.defineSlideMaster({ title: "COVER", background: { color: C.text2 },
  objects: [titlePh(C.background1, { x: 0.7, y: 3.05, w: 5.7, h: 1.0, fontSize: 24, bold: false, valign: "top" }), ...footerObjs(true, 5.6)] });
pres.defineSlideMaster({ title: "HERO_GD", background: { color: C.text2 },
  objects: [chapPh(C.accent2, 0.6, 0.45, 6.2), titlePh(C.background1, { y: 0.95, w: 6.1, h: 1.95, fontSize: 26, valign: "top" }), ...footerObjs(true, 5.3)],
  slideNumber: slideNo("A9C4CF", 6.0) });
pres.defineSlideMaster({ title: "HERO_END", background: { color: C.text2 },
  objects: [chapPh(C.accent2, 0.6, 0.45, 6.4), titlePh(C.background1, { y: 2.05, w: 6.6, h: 1.9, fontSize: 46, valign: "top" }), ...footerObjs(true, 5.6)],
  slideNumber: slideNo("A9C4CF", 6.6) });

// ---- 部品 ---------------------------------------------------------------
const iconCache = {};
function ic(name, hex = HEX.navy) {
  const key = name + hex;
  if (!iconCache[key]) {
    const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Tb[name], { size: 256, color: "#" + hex }));
    iconCache[key] = sharp(Buffer.from(svg)).resize(256, 256).png().toBuffer().then((b) => "image/png;base64," + b.toString("base64"));
  }
  return iconCache[key];
}

// 角丸トリミング済みの画像を作る（モチーフ：角丸フレーム）
async function crop(name, w, h, { radius = 0.14, pos = "centre" } = {}) {
  const pw = Math.round(w * 200), ph = Math.round(h * 200), r = Math.round(radius * 200);
  const file = path.join(CROPS, `${name}_${pw}x${ph}_${r}_${pos}.png`);
  if (!fs.existsSync(file)) {
    let img = sharp(path.join(ASSETS, name + ".png")).resize(pw, ph, { fit: "cover", position: pos });
    if (r > 0) {
      const mask = Buffer.from(`<svg width="${pw}" height="${ph}"><rect width="${pw}" height="${ph}" rx="${r}" ry="${r}"/></svg>`);
      img = sharp(await img.png().toBuffer()).composite([{ input: mask, blend: "dest-in" }]);
    }
    await img.png().toFile(file);
  }
  return file;
}
async function pic(slide, name, x, y, w, h, o = {}) {
  slide.addImage({ path: await crop(name, w, h, o), x, y, w, h, altText: o.alt || "イメージ図（一般的なイラスト。実在企業の設備ではありません）" });
}

function T(slide, text, o) {
  slide.addText(text, Object.assign({ isTextBox: true, margin: 0, lang: "ja-JP", valign: "top", fontSize: 20, color: FILL.text, paraSpaceAfter: 0 }, o));
}
const shadow = () => ({ type: "outer", blur: 8, offset: 2, angle: 90, color: "12243A", opacity: 0.14 });
function box(slide, x, y, w, h, fill, o = {}) {
  slide.addShape(o.rect ? SH.RECTANGLE : SH.ROUNDED_RECTANGLE, Object.assign({
    x, y, w, h, fill: { color: fill }, line: o.lineColor ? { color: o.lineColor, width: o.lineWidth || 1.25, dashType: o.dash } : { type: "none" },
    rectRadius: o.rect ? undefined : (o.radius != null ? o.radius : 0.12), shadow: o.shadow ? shadow() : undefined,
  }, o.extra || {}));
}
function circle(slide, x, y, d, fill, o = {}) {
  slide.addShape(SH.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: o.lineColor ? { color: o.lineColor, width: o.lineWidth || 1.5 } : { type: "none" }, shadow: o.shadow ? shadow() : undefined });
}
// 円の中にアイコン
async function iconCircle(slide, name, x, y, d, fill, iconHex) {
  circle(slide, x, y, d, fill);
  const s = d * 0.56;
  slide.addImage({ data: await ic(name, iconHex), x: x + (d - s) / 2, y: y + (d - s) / 2, w: s, h: s, altText: name });
}
async function icon(slide, name, x, y, s, hex) {
  slide.addImage({ data: await ic(name, hex), x, y, w: s, h: s, altText: name });
}
// 番号バッジ
function badge(slide, label, x, y, d, fill, color = FILL.white, size) {
  slide.addText(label, { shape: SH.OVAL, x, y, w: d, h: d, fill: { color: fill }, color, bold: true, fontSize: size || Math.round(d * 32), align: "center", valign: "middle", margin: 0, isTextBox: true, lang: "ja-JP" });
}
// 任意方向の線（矢印付き可）
function line(slide, x1, y1, x2, y2, o = {}) {
  const flipH = x2 < x1, flipV = y2 < y1;
  slide.addShape(SH.LINE, {
    x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.max(Math.abs(x2 - x1), 0.001), h: Math.max(Math.abs(y2 - y1), 0.001),
    flipH: flipH && !flipV ? true : flipH && flipV ? true : false, flipV: flipV,
    line: { color: o.color || FILL.line, width: o.width || 2, dashType: o.dash, endArrowType: o.arrow ? "triangle" : undefined },
  });
}
function arrowR(slide, x, y, w, h, fill = FILL.cyan) {
  slide.addShape(SH.RIGHT_ARROW, { x, y, w, h, fill: { color: fill }, line: { type: "none" } });
}
function pill(slide, text, x, y, w, h, fill, color = FILL.white, size = 16, o = {}) {
  slide.addText(text, Object.assign({ shape: SH.ROUNDED_RECTANGLE, rectRadius: h / 2, x, y, w, h, fill: { color: fill }, color, bold: true, fontSize: size, align: "center", valign: "middle", margin: 0, isTextBox: true, lang: "ja-JP",
    line: o.lineColor ? { color: o.lineColor, width: 1.5 } : { type: "none" } }, o.extra || {}));
}
// 出典・脚注（12pt以上）。warn は要確認の注記（オレンジ）
function foot(slide, text, warn) {
  const runs = [{ text, options: { color: "5B6B7C" } }];
  if (warn) runs.push({ text: "　【要確認】" + warn, options: { color: "B86E10", bold: true } });
  T(slide, runs, { x: 0.6, y: 6.5, w: 12.13, h: 0.4, fontSize: 12, valign: "bottom" });
}
// GD用のタイマー
function timer(slide, label, x, y, d = 2.0) {
  circle(slide, x, y, d, FILL.navy, { shadow: true });
  circle(slide, x + 0.12, y + 0.12, d - 0.24, FILL.navy, { lineColor: FILL.cyan, lineWidth: 3 });
  T(slide, label, { x, y: y + d * 0.26, w: d, h: d * 0.36, fontSize: Math.round(d * 20), bold: true, color: FILL.white, align: "center", valign: "middle" });
  T(slide, "TIMER", { x, y: y + d * 0.63, w: d, h: 0.3, fontSize: 13, bold: true, color: FILL.cyan, align: "center" });
}
// 時間配分バー（秒数に比例した幅）
function timeline(slide, parts, y, x0 = 0.6, W = 12.13) {
  const total = parts.reduce((a, p) => a + p[1], 0);
  let x = x0;
  const fills = [FILL.orange, FILL.teal, FILL.navy];
  T(slide, "進行", { x: x0, y: y - 0.36, w: 2, h: 0.3, fontSize: 14, bold: true, color: FILL.gray });
  parts.forEach(([label, sec], i) => {
    const w = (W * sec) / total - (i < parts.length - 1 ? 0.06 : 0);
    slide.addText(label, { shape: SH.RECTANGLE, x, y, w, h: 0.5, fill: { color: fills[i % 3] }, color: i % 3 === 0 ? FILL.navy : FILL.white, bold: true, fontSize: 16, align: "center", valign: "middle", margin: 0, isTextBox: true, lang: "ja-JP" });
    x += w + 0.06;
  });
}

function addSlide(n, master = "LIGHT") {
  const ch = chapterOf(n);
  const slide = pres.addSlide({ masterName: master, sectionTitle: CHAPTERS[ch] });
  if (master !== "COVER") slide.addText(CHAPTERS[ch], { placeholder: "chap" });
  if (master !== "HERO_END") slide.addText(S[n].slideTitle || S[n].title, { placeholder: "title" });
  slide.addNotes(notesFor(n));
  return slide;
}
let lastCh = 0;
function section(n) {
  const ch = chapterOf(n);
  if (ch !== lastCh) { pres.addSection({ title: CHAPTERS[ch] }); lastCh = ch; }
}

// ---- 各スライド ----------------------------------------------------------
const slides = {};

slides[1] = async () => {
  const s = addSlide(1, "COVER");
  T(s, "第1回｜ミラキン業界研究会", { x: 0.7, y: 0.9, w: 5.7, h: 0.45, fontSize: 20, bold: true, color: FILL.cyan });
  T(s, "金属業界", { x: 0.7, y: 1.4, w: 5.8, h: 1.45, fontSize: 80, bold: true, color: FILL.white, valign: "middle" });
  const labels = [["01", "企業を知る"], ["02", "業界を知る"], ["03", "お金の流れを知る"], ["04", "未来を考える"]];
  labels.forEach(([no, t], i) => {
    const x = 0.7 + (i % 2) * 2.9, y = 4.45 + Math.floor(i / 2) * 0.85;
    box(s, x, y, 2.7, 0.66, FILL.navy, { lineColor: FILL.cyan, lineWidth: 1.5, radius: 0.33 });
    T(s, [{ text: no + " ", options: { color: FILL.orange, bold: true, fontSize: 15 } }, { text: t, options: { color: FILL.white, bold: true } }],
      { x: x + 0.18, y, w: 2.5, h: 0.66, fontSize: 17, valign: "middle" });
  });
  // 右1/2：産業がつながるイメージのコラージュ
  const tiles = [["city_bridge", 6.75, 0], ["port", 10.09, 0], ["chip", 6.75, 3.79], ["steel_mill", 10.09, 3.79]];
  for (const [nm, x, y] of tiles) await pic(s, nm, x, y, 3.25, 3.71, { radius: 0 });
};

slides[2] = async () => {
  const s = addSlide(2);
  const L = [["TbToolsKitchen2", "食品"], ["TbDeviceLaptop", "IT"], ["TbStethoscope", "医療"], ["TbBuildingBank", "金融"]];
  const R = [["TbWorld", "商社"], ["TbBuildingFactory2", "メーカー"], ["TbSpeakerphone", "広告"], ["TbTruck", "物流"]];
  const tile = async (arr, x0) => {
    for (let i = 0; i < 4; i++) {
      const x = x0 + (i % 2) * 1.2, y = 2.0 + Math.floor(i / 2) * 1.55;
      await iconCircle(s, arr[i][0], x + 0.15, y, 0.8, FILL.pale, HEX.teal);
      T(s, arr[i][1], { x: x - 0.1, y: y + 0.88, w: 1.3, h: 0.4, fontSize: 16, bold: true, color: FILL.navy, align: "center" });
    }
  };
  await tile(L, 0.6);
  await tile(R, 10.53);
  box(s, 3.15, 1.85, 7.05, 2.9, FILL.pale, { radius: 0.2 });
  T(s, "いま興味のある業界を、\n1〜3個思い浮かべよう。", { x: 3.45, y: 2.1, w: 6.45, h: 1.35, fontSize: 32, bold: true, color: FILL.navy, align: "center", valign: "middle" });
  T(s, "「その業界を知ったきっかけ」は何だろう？", { x: 3.45, y: 3.65, w: 6.45, h: 0.6, fontSize: 22, bold: true, color: FILL.teal, align: "center", valign: "middle" });
  T(s, "きっかけの例", { x: 3.15, y: 4.95, w: 7.05, h: 0.35, fontSize: 16, bold: true, color: FILL.gray, align: "center" });
  ["商品を使った", "ニュースで見た", "家族の仕事", "働き方のイメージ…"].forEach((t, i) => {
    pill(s, t, 1.27 + i * 2.75, 5.4, 2.6, 0.6, FILL.white, FILL.navy, 18, { lineColor: FILL.teal });
  });
};

slides[3] = async () => {
  const s = addSlide(3);
  const sectors = ["水産・農林業", "鉱業", "建設業", "食料品", "繊維製品", "パルプ・紙", "化学", "医薬品", "石油・石炭製品", "ゴム製品", "ガラス・土石製品",
    "鉄鋼", "非鉄金属", "金属製品", "機械", "電気機器", "輸送用機器", "精密機器", "その他製品", "電気・ガス業", "陸運業", "海運業",
    "空運業", "倉庫・運輸関連業", "情報・通信業", "卸売業", "小売業", "銀行業", "証券・商品先物取引業", "保険業", "その他金融業", "不動産業", "サービス業"];
  const hl = new Set(["鉄鋼", "非鉄金属", "金属製品", "卸売業"]);
  const cw = 2.79, ch = 0.35, gap = 0.055;
  sectors.forEach((t, i) => {
    const col = Math.floor(i / 11), row = i % 11;
    const x = 0.6 + col * (cw + 0.12), y = 1.85 + row * (ch + gap);
    const on = hl.has(t);
    box(s, x, y, cw, ch, on ? FILL.teal : FILL.pale, { radius: 0.06 });
    T(s, t, { x: x + 0.14, y, w: cw - 0.2, h: ch, fontSize: 16, bold: on, color: on ? FILL.white : FILL.navy, valign: "middle" });
  });
  T(s, "東証33業種", { x: 9.55, y: 1.85, w: 3.18, h: 0.45, fontSize: 24, bold: true, color: FILL.navy });
  T(s, "上場企業を分ける「ひとつの地図」", { x: 9.55, y: 2.32, w: 3.18, h: 0.7, fontSize: 18, color: FILL.text });
  box(s, 9.55, 3.2, 3.18, 1.55, FILL.pale);
  await icon(s, "TbSearch", 9.75, 3.35, 0.45, HEX.teal);
  T(s, "知っている業種は？", { x: 10.3, y: 3.37, w: 2.35, h: 0.42, fontSize: 18, bold: true, color: FILL.navy, valign: "middle" });
  T(s, "企業名まで\n思い浮かぶ業種は？", { x: 9.75, y: 3.9, w: 2.85, h: 0.75, fontSize: 18, bold: true, color: FILL.navy });
  box(s, 9.6, 5.0, 0.32, 0.32, FILL.teal, { radius: 0.05 });
  T(s, "今回と関わる業種", { x: 10.02, y: 4.97, w: 2.7, h: 0.38, fontSize: 18, bold: true, color: FILL.teal, valign: "middle" });
  T(s, "鉄鋼／非鉄金属／\n金属製品／卸売業", { x: 9.6, y: 5.42, w: 3.13, h: 0.8, fontSize: 18, color: FILL.text });
  foot(s, "※東証33業種は市場向けの分類で、就活の業界分類や「金属産業全体」とは完全には一致しない。　出典：JPX/J-Quants 業種一覧");
};

slides[4] = async () => {
  const s = addSlide(4);
  const card = async (x, img, tag, head, sub) => {
    box(s, x, 1.8, 5.92, 3.8, FILL.white, { lineColor: FILL.line, shadow: true, radius: 0.18 });
    await pic(s, img, x + 0.22, 2.02, 5.48, 2.2);
    pill(s, tag, x + 0.4, 2.2, 2.3, 0.42, FILL.navy, FILL.white, 14);
    T(s, head, { x: x + 0.3, y: 4.4, w: 5.3, h: 0.5, fontSize: 24, bold: true, color: FILL.navy });
    T(s, sub, { x: x + 0.3, y: 4.95, w: 5.3, h: 0.45, fontSize: 20, color: FILL.teal, bold: true });
  };
  await card(0.6, "consumer", "生活者向け（BtoC）", "よく見かける企業", "食品・化粧品・アプリ・ゲーム");
  await card(6.81, "steel_mill", "企業向け（BtoB）", "名前を見にくい企業", "素材・部品・設備・専門商社");
  box(s, 0.6, 5.85, 12.13, 0.78, FILL.navy, { radius: 0.12 });
  T(s, "知らない会社は、まだ魅力に出会う機会がなかった会社かもしれない。", { x: 0.9, y: 5.85, w: 11.6, h: 0.78, fontSize: 22, bold: true, color: FILL.white, valign: "middle" });
};

slides[5] = async () => {
  const s = addSlide(5);
  // 中心ノード
  circle(s, 0.6, 2.5, 2.7, FILL.teal, { shadow: true });
  T(s, "興味の表面", { x: 0.6, y: 3.0, w: 2.7, h: 0.4, fontSize: 16, bold: true, color: FILL.white, align: "center" });
  T(s, "コンサル\nティング", { x: 0.6, y: 3.35, w: 2.7, h: 1.2, fontSize: 28, bold: true, color: FILL.white, align: "center", valign: "middle" });
  T(s, "惹かれる理由", { x: 3.85, y: 1.72, w: 3.5, h: 0.35, fontSize: 16, bold: true, color: FILL.gray });
  T(s, "他業界の候補", { x: 8.45, y: 1.72, w: 3.9, h: 0.35, fontSize: 16, bold: true, color: FILL.gray });
  const rows = [["課題解決が好き", "IT・広告・人材", "TbBulb"], ["経営に関わりたい", "金融・商社・事業会社", "TbBriefcase"], ["多くの企業と働きたい", "専門商社・法人営業・コンサル", "TbUsers"]];
  for (let i = 0; i < 3; i++) {
    const y = 2.15 + i * 1.45, cy = y + 0.55;
    line(s, 3.3, 3.85, 3.85, cy, { color: FILL.teal, width: 2.25 });
    box(s, 3.85, y, 3.95, 1.1, FILL.pale);
    await icon(s, rows[i][2], 4.03, y + 0.31, 0.48, HEX.teal);
    T(s, rows[i][0], { x: 4.63, y, w: 3.1, h: 1.1, fontSize: 20, bold: true, color: FILL.navy, valign: "middle" });
    arrowR(s, 7.92, cy - 0.2, 0.42, 0.4);
    box(s, 8.45, y, 4.28, 1.1, FILL.white, { lineColor: FILL.teal, lineWidth: 1.75 });
    T(s, rows[i][1], { x: 8.65, y, w: 3.98, h: 1.1, fontSize: 19, bold: true, color: FILL.teal, valign: "middle" });
  }
};

slides[6] = async () => {
  const s = addSlide(6);
  const cards = [[0.6, 1.85, "TbWorld", "海外で働きたい", "海外展開するメーカー"], [8.13, 1.85, "TbBuildingSkyscraper", "大きな事業を動かしたい", "インフラ・素材企業"],
    [0.6, 4.45, "TbCoin", "事業投資に関わりたい", "金融・事業会社"], [8.13, 4.45, "TbHeartHandshake", "人と企業をつなぎたい", "物流・法人営業"]];
  // ハブと結ぶ線（先に描いてカードの下に置く）
  [[5.2, 2.7], [8.13, 2.7], [5.2, 5.3], [8.13, 5.3]].forEach(([x, y]) => line(s, 6.67, 4.0, x, y, { color: FILL.line, width: 2.5 }));
  for (const [x, y, ico, head, cand] of cards) {
    box(s, x, y, 4.6, 1.7, FILL.white, { lineColor: FILL.line, shadow: true, radius: 0.16 });
    await iconCircle(s, ico, x + 0.25, y + 0.3, 0.8, FILL.pale, HEX.teal);
    T(s, head, { x: x + 1.2, y: y + 0.22, w: 3.32, h: 0.72, fontSize: 21, bold: true, color: FILL.navy, valign: "middle" });
    T(s, [{ text: "→ ", options: { color: FILL.orange } }, { text: cand }], { x: x + 1.2, y: y + 0.98, w: 3.32, h: 0.5, fontSize: 19, bold: true, color: FILL.teal, valign: "middle" });
  }
  s.addText([{ text: "総合商社", options: { fontSize: 28, bold: true, breakLine: true } }, { text: "志望", options: { fontSize: 16, color: FILL.cyan } }],
    { shape: SH.ROUNDED_RECTANGLE, rectRadius: 0.2, x: 5.47, y: 3.25, w: 2.4, h: 1.5, fill: { color: FILL.navy }, color: FILL.white, align: "center", valign: "middle", margin: 0, isTextBox: true, lang: "ja-JP", shadow: shadow() });
  foot(s, "※例示であり、職種を限定するものではありません。");
};

slides[7] = async () => {
  const s = addSlide(7);
  T(s, "2027卒対象の業界イメージ調査", { x: 0.6, y: 1.8, w: 6.2, h: 0.45, fontSize: 20, bold: true, color: FILL.teal });
  T(s, [{ text: "10.1", options: { fontSize: 120 } }, { text: "％", options: { fontSize: 60 } }], { x: 0.5, y: 2.15, w: 6.3, h: 1.85, bold: true, color: FILL.teal, valign: "middle" });
  T(s, "「鉄鋼・金属・鉱業」を就職先として\n検討したことがある", { x: 0.6, y: 4.05, w: 6.3, h: 0.95, fontSize: 22, bold: true, color: FILL.navy });
  // 10人のシルエット（1人を強調）
  for (let i = 0; i < 10; i++) {
    const x = 7.25 + (i % 5) * 1.07, y = 1.95 + Math.floor(i / 5) * 1.45;
    await icon(s, "TbUserFilled", x, y, 1.0, i === 0 ? HEX.orange : "B9CBD3");
  }
  T(s, [{ text: "約10人に1人", options: { bold: true, color: FILL.orange } }, { text: "（比率のイメージ）", options: { color: FILL.gray, fontSize: 16 } }],
    { x: 7.25, y: 4.85, w: 5.4, h: 0.45, fontSize: 22, align: "center" });
  box(s, 0.6, 5.45, 12.13, 0.85, FILL.pale);
  await icon(s, "TbAlertTriangle", 0.85, 5.62, 0.5, HEX.orange);
  T(s, [{ text: "測定しているのは「興味」ではなく、" }, { text: "「就職先として検討した経験」", options: { bold: true, color: FILL.teal } }],
    { x: 1.55, y: 5.45, w: 11.0, h: 0.85, fontSize: 20, color: FILL.navy, valign: "middle" });
  foot(s, "※ 2026年4月8〜18日／回答1,579人／複数回答　出典：マイナビ「2027年卒 大学生業界イメージ調査」", "10.1％は一次資料で再確認");
};

slides[8] = async () => {
  const s = addSlide(8, "DARK");
  T(s, "知らない ≠ 興味がない", { x: 0.6, y: 1.95, w: 8.6, h: 1.1, fontSize: 54, bold: true, color: FILL.white, valign: "middle" });
  T(s, "知らない ＝ まだ触れたことがない可能性", { x: 0.6, y: 3.15, w: 8.6, h: 0.7, fontSize: 28, bold: true, color: FILL.cyan, valign: "middle" });
  const steps = ["知る", "比べる", "興味が生まれる"];
  let x = 0.6;
  steps.forEach((t, i) => {
    const w = i === 2 ? 3.0 : 1.9;
    pill(s, t, x, 4.55, w, 0.85, i === 2 ? FILL.orange : FILL.navy, i === 2 ? FILL.navy : FILL.white, 24, { lineColor: i === 2 ? undefined : FILL.cyan });
    x += w;
    if (i < 2) { arrowR(s, x + 0.15, 4.75, 0.5, 0.45, FILL.cyan); x += 0.8; }
  });
  circle(s, 9.75, 1.75, 2.9, "1C3A57");
  await icon(s, "TbTelescope", 10.4, 2.4, 1.6, HEX.cyan);
  await icon(s, "TbMap2", 11.95, 4.15, 0.8, HEX.orange);
};

slides[9] = async () => {
  const s = addSlide(9);
  pill(s, "架空の例", 10.93, 0.85, 1.8, 0.5, FILL.orange, FILL.navy, 18);
  const chartCard = (x, head, vals, cap) => {
    box(s, x, 1.8, 4.45, 3.85, FILL.white, { lineColor: FILL.line, shadow: true, radius: 0.16 });
    T(s, head, { x: x + 0.25, y: 1.95, w: 4.0, h: 0.5, fontSize: 20, bold: true, color: FILL.navy });
    s.addChart(pres.charts.BAR, [{ name: "営業利益の伸び率", labels: ["企業A", "業界全体"], values: vals }], {
      x: x + 0.2, y: 2.45, w: 4.05, h: 2.55, barDir: "col", chartColors: [HEX.teal, HEX.gray], invertIfNegative: false,
      showValue: true, dataLabelFormatCode: '+0"%";-0"%"', dataLabelFontSize: 18, dataLabelFontBold: true, dataLabelColor: HEX.navy, dataLabelPosition: "outEnd", dataLabelFontFace: FONT,
      valAxisHidden: true, valAxisMinVal: -40, valAxisMaxVal: 70, valGridLine: { style: "none" }, catGridLine: { style: "none" },
      catAxisLabelFontSize: 16, catAxisLabelColor: HEX.text, catAxisLabelFontFace: FONT, catAxisLineShow: true, catAxisLabelPos: "low", showLegend: false, barGapWidthPct: 60,
    });
    T(s, cap, { x: x + 0.25, y: 5.05, w: 4.0, h: 0.45, fontSize: 16, color: FILL.gray, bold: true });
  };
  chartCard(0.6, "ケースA：業界全体 ＋50％", [20, 50], "追い風の中では、業界平均を下回る");
  chartCard(8.28, "ケースB：業界全体 −20％", [20, -20], "逆風の中での増益は、独自の強み？");
  s.addText([{ text: "架空の企業A", options: { fontSize: 18, color: FILL.cyan, breakLine: true } }, { text: "営業利益", options: { fontSize: 20, breakLine: true } }, { text: "＋20％", options: { fontSize: 50, bold: true } }],
    { shape: SH.ROUNDED_RECTANGLE, rectRadius: 0.2, x: 5.29, y: 2.25, w: 2.75, h: 2.9, fill: { color: FILL.navy }, color: FILL.white, align: "center", valign: "middle", margin: 0, isTextBox: true, lang: "ja-JP", shadow: shadow() });
  box(s, 0.6, 5.85, 12.13, 0.6, FILL.pale);
  T(s, "同じ＋20％でも「企業独自の強さ」の見方は変わる。", { x: 0.9, y: 5.85, w: 11.6, h: 0.6, fontSize: 22, bold: true, color: FILL.navy, valign: "middle" });
  foot(s, "※架空の数値による説明用の例。実際の分析では前年の利益水準や一時的な要因も確認する。");
};

slides[10] = async () => {
  const s = addSlide(10);
  const cx = 4.5, cy = 4.1;
  const nodes = [["顧客", "TbShoppingCart", -90, "注文"], ["仕入先", "TbPackage", -18, "原料"], ["競合", "TbScale", 54, "価格競争"], ["社会・規制", "TbShieldCheck", 126, "規制"], ["株主・金融機関", "TbBuildingBank", 198, "資金"]];
  const pos = nodes.map(([, , a]) => [cx + 2.95 * Math.cos((a * Math.PI) / 180), cy + 1.78 * Math.sin((a * Math.PI) / 180)]);
  pos.forEach(([x, y]) => line(s, cx, cy, x, y, { color: FILL.line, width: 2.5 }));
  circle(s, cx - 0.72, cy - 0.72, 1.44, FILL.navy, { shadow: true });
  T(s, "企業", { x: cx - 0.72, y: cy - 0.72, w: 1.44, h: 1.44, fontSize: 30, bold: true, color: FILL.white, align: "center", valign: "middle" });
  pos.forEach(([x, y], i) => {
    const mx = cx + 0.6 * (x - cx), my = cy + 0.6 * (y - cy);
    pill(s, nodes[i][3], mx - 0.55, my - 0.2, 1.1, 0.4, FILL.white, FILL.teal, 15, { lineColor: FILL.teal });
  });
  for (let i = 0; i < 5; i++) {
    const [x, y] = pos[i];
    await iconCircle(s, nodes[i][1], x - 0.42, y - 0.42, 0.84, FILL.teal, HEX.white);
    if (i === 0) T(s, nodes[i][0], { x: x + 0.52, y: y - 0.2, w: 1.2, h: 0.4, fontSize: 18, bold: true, color: FILL.navy, valign: "middle" });
    else T(s, nodes[i][0], { x: x - 1.0, y: y + 0.45, w: 2.0, h: 0.4, fontSize: 18, bold: true, color: FILL.navy, align: "center" });
  }
  T(s, "企業を取り巻く5つの関係", { x: 8.55, y: 1.95, w: 4.18, h: 0.5, fontSize: 24, bold: true, color: FILL.navy });
  T(s, "顧客／仕入先／競合／\n社会・規制／株主・金融機関", { x: 8.55, y: 2.55, w: 4.18, h: 0.9, fontSize: 18, color: FILL.text });
  box(s, 8.55, 3.65, 4.18, 1.55, FILL.navy, { shadow: true, radius: 0.16 });
  T(s, "売上・利益は、\nつながりの中で変化する。", { x: 8.8, y: 3.65, w: 3.8, h: 1.55, fontSize: 24, bold: true, color: FILL.white, valign: "middle" });
  T(s, "例：自動車の生産が減る → 鋼材の注文も減り得る", { x: 8.55, y: 5.4, w: 4.18, h: 0.75, fontSize: 16, color: FILL.gray });
};

slides[11] = async () => {
  const s = addSlide(11);
  const steps = ["業界全体", "企業の\nつながり", "儲け方", "個別企業", "KPI", "将来予測"];
  T(s, "地図を見る", { x: 0.6, y: 1.72, w: 5, h: 0.35, fontSize: 16, bold: true, color: FILL.navy });
  T(s, "目的地を深掘りする", { x: 6.7, y: 1.72, w: 5, h: 0.35, fontSize: 16, bold: true, color: FILL.teal });
  steps.forEach((t, i) => {
    s.addText(t, { shape: i === 0 ? SH.PENTAGON : SH.CHEVRON, x: 0.6 + i * 2.0, y: 2.1, w: 2.13, h: 1.0, fill: { color: i < 3 ? FILL.navy : FILL.teal }, color: FILL.white,
      bold: true, fontSize: 18, align: "center", valign: "middle", margin: 0, isTextBox: true, lang: "ja-JP", line: { type: "none" } });
  });
  // 地図 → 拡大の図
  box(s, 0.6, 3.45, 7.0, 2.3, FILL.pale, { radius: 0.18 });
  T(s, "金属業界の工程（地図）", { x: 0.85, y: 3.58, w: 5, h: 0.4, fontSize: 18, bold: true, color: FILL.navy });
  const flow = ["資源", "製鉄・製錬", "加工", "流通", "製品"];
  const fw = [0.95, 1.55, 0.95, 0.95, 0.95];
  let fx = 0.85;
  flow.forEach((t, i) => {
    pill(s, t, fx, 4.35, fw[i], 0.62, FILL.white, i === 1 ? FILL.navy : FILL.gray, 15, { lineColor: i === 1 ? FILL.orange : FILL.line });
    if (i < 4) T(s, "→", { x: fx + fw[i], y: 4.4, w: 0.3, h: 0.5, fontSize: 14, color: FILL.gray, align: "center", valign: "middle" });
    fx += fw[i] + 0.3;
  });
  T(s, "まず全体の位置関係を把握", { x: 0.85, y: 5.15, w: 6.5, h: 0.4, fontSize: 16, color: FILL.gray });
  line(s, 3.35, 4.35, 8.1, 3.5, { color: FILL.orange, width: 1.5, dash: "dash" });
  line(s, 3.35, 4.97, 8.1, 5.75, { color: FILL.orange, width: 1.5, dash: "dash" });
  box(s, 8.1, 3.45, 4.63, 2.3, FILL.white, { lineColor: FILL.orange, lineWidth: 2, shadow: true, radius: 0.18 });
  await iconCircle(s, "TbSearch", 8.35, 3.65, 0.7, FILL.orange, HEX.navy);
  T(s, "個別企業へズーム", { x: 9.2, y: 3.65, w: 3.4, h: 0.35, fontSize: 15, bold: true, color: FILL.gray });
  T(s, "例：日本製鉄", { x: 9.2, y: 3.98, w: 3.4, h: 0.45, fontSize: 22, bold: true, color: FILL.navy });
  T(s, [{ text: "見るKPIの例：", options: { color: FILL.gray } }, { text: "製品単価", options: { bold: true, color: FILL.teal } }], { x: 8.35, y: 4.6, w: 4.2, h: 0.45, fontSize: 20 });
  T(s, "KPIは業界・企業によって異なる", { x: 8.35, y: 5.1, w: 4.2, h: 0.45, fontSize: 16, color: FILL.gray });
  T(s, "全体像があるから、企業の違いが見える。", { x: 0.6, y: 5.95, w: 12.13, h: 0.5, fontSize: 22, bold: true, color: FILL.navy, valign: "middle" });
  foot(s, "※KPI＝重要業績評価指標（Key Performance Indicator）。");
};

const QS12 = ["社会に\n何を届ける？", "どんな企業が\nいる？", "誰が顧客？", "どうやって\n利益を生む？", "企業の強みは？", "どんなニュースが\n影響する？"];
slides[12] = async () => {
  const s = addSlide(12);
  T(s, "リサーチ（①〜⑥）", { x: 0.6, y: 1.72, w: 5, h: 0.4, fontSize: 18, bold: true, color: FILL.teal });
  T(s, "自分への接続（⑦）", { x: 9.45, y: 1.72, w: 3.3, h: 0.4, fontSize: 18, bold: true, color: "B86E10" });
  QS12.forEach((q, i) => {
    const x = 0.6 + (i % 3) * 2.65, y = 2.2 + Math.floor(i / 3) * 2.05;
    box(s, x, y, 2.45, 1.85, FILL.pale, { radius: 0.16 });
    badge(s, String(i + 1), x + 0.2, y + 0.2, 0.6, FILL.teal, FILL.white, 22);
    T(s, q, { x: x + 0.15, y: y + 0.88, w: 2.25, h: 0.85, fontSize: 19, bold: true, color: FILL.navy });
  });
  arrowR(s, 8.55, 3.85, 0.7, 0.6, FILL.orange);
  box(s, 9.45, 2.2, 3.28, 3.9, FILL.orange, { shadow: true, radius: 0.2 });
  badge(s, "7", 9.7, 2.45, 0.75, FILL.navy, FILL.white, 28);
  await icon(s, "TbHeart", 11.7, 2.5, 0.7, HEX.navy);
  T(s, "自分は何に\n興味・共感する？", { x: 9.7, y: 3.5, w: 2.9, h: 1.4, fontSize: 26, bold: true, color: FILL.navy });
  T(s, "調べた内容を、自分の選択肢に変える", { x: 9.7, y: 5.0, w: 2.9, h: 0.8, fontSize: 16, color: FILL.navy });
  T(s, "金属業界だけでなく、食品・IT・医薬品など、どの業界にも使える。", { x: 0.6, y: 6.35, w: 12.13, h: 0.45, fontSize: 18, bold: true, color: FILL.teal });
};

slides[13] = async () => {
  const s = addSlide(13);
  box(s, 0.6, 1.75, 12.13, 0.78, FILL.pale, { radius: 0.14 });
  T(s, [{ text: "金属業界 ＝ ", options: { color: FILL.teal } }, { text: "金属を「取り出し、使える形にし、届ける」産業" }], { x: 0.9, y: 1.75, w: 11.6, h: 0.78, fontSize: 24, bold: true, color: FILL.navy, valign: "middle" });
  const st = [["資源開発", "TbPick"], ["製鉄・製錬", "TbFlame"], ["加工", "TbTool"], ["商社・物流", "TbTruck"], ["回収・再生", "TbRecycle"]];
  for (let i = 0; i < st.length; i++) {
    const y = 2.8 + i * 0.75;
    box(s, 0.6, y, 5.35, 0.64, i === 4 ? FILL.teal : FILL.navy, { radius: 0.1 });
    badge(s, String(i + 1), 0.75, y + 0.1, 0.44, FILL.white, i === 4 ? FILL.teal : FILL.navy, 16);
    await icon(s, st[i][1], 1.38, y + 0.12, 0.4, i === 4 ? HEX.white : HEX.cyan);
    T(s, st[i][0], { x: 1.95, y, w: 3.8, h: 0.64, fontSize: 22, bold: true, color: FILL.white, valign: "middle" });
  }
  await pic(s, "city_bridge", 6.3, 2.8, 6.43, 3.64);
  s.addShape(SH.ROUNDED_RECTANGLE, { x: 10.83, y: 4.7, w: 1.8, h: 1.62, fill: { color: FILL.white }, line: { type: "none" }, rectRadius: 0.12, shadow: shadow() });
  await pic(s, "chip", 10.89, 4.76, 1.68, 1.5, { radius: 0.1 });
};

slides[14] = async () => {
  const s = addSlide(14);
  const items = [["TbBuildingBridge2", "建物・橋", "鉄骨、鉄筋"], ["TbCar", "自動車・鉄道", "鋼板、アルミ"], ["TbBolt", "送電", "銅、アルミ"], ["TbDeviceMobile", "スマホ・\n電子機器", "銅、高純度金属"], ["TbPlane", "医療・航空", "特殊合金、チタン"]];
  for (let i = 0; i < 5; i++) {
    const x = 0.6 + i * 2.476, [ico, name, mat] = items[i];
    box(s, x, 1.85, 2.23, 3.85, FILL.white, { lineColor: FILL.line, shadow: true, radius: 0.16 });
    box(s, x + 0.12, 1.97, 1.99, 1.95, i % 2 ? FILL.teal : FILL.navy, { radius: 0.12 });
    await icon(s, ico, x + 0.51, 2.33, 1.22, HEX.white);
    T(s, name, { x: x + 0.15, y: 4.1, w: 1.95, h: 0.75, fontSize: 21, bold: true, color: FILL.navy, align: "center", valign: "middle" });
    T(s, mat, { x: x + 0.05, y: 4.9, w: 2.13, h: 0.55, fontSize: 17, bold: true, color: FILL.teal, align: "center", valign: "middle" });
  }
  T(s, "金属が目立つのではなく、完成品の性能を支えている。", { x: 0.6, y: 6.0, w: 12.13, h: 0.5, fontSize: 22, bold: true, color: FILL.navy, align: "center", valign: "middle" });
};

slides[15] = async () => {
  const s = addSlide(15);
  T(s, "金属", { x: 0.85, y: 1.75, w: 2.5, h: 0.35, fontSize: 16, bold: true, color: FILL.gray });
  T(s, "得意なこと", { x: 3.85, y: 1.75, w: 4, h: 0.35, fontSize: 16, bold: true, color: FILL.gray });
  T(s, "主な用途", { x: 8.75, y: 1.75, w: 3.8, h: 0.35, fontSize: 16, bold: true, color: FILL.gray });
  const rows = [["m_steel", "鉄鋼", "強度とコストのバランス", "橋・車体"], ["m_alu", "アルミ", "軽い", "輸送機器・飲料缶"], ["m_copper", "銅", "電気を通す", "送電線・電子機器"],
    ["m_nickel", "ニッケル", "合金の耐食性など", "ステンレス・電池材料"], ["m_titan", "チタン", "軽さに対する強度・耐食性", "航空・医療"]];
  for (let i = 0; i < 5; i++) {
    const y = 2.15 + i * 0.76, [img, name, prop, use] = rows[i];
    box(s, 0.6, y, 12.13, 0.66, i % 2 ? FILL.white : FILL.pale, { radius: 0.1, lineColor: i % 2 ? FILL.line : undefined });
    s.addImage({ path: path.join(ASSETS, img + ".png"), x: 0.75, y: y + 0.06, w: 0.54, h: 0.54, altText: name + "の質感イメージ" });
    T(s, name, { x: 1.5, y, w: 2.2, h: 0.66, fontSize: 22, bold: true, color: FILL.navy, valign: "middle" });
    T(s, prop, { x: 3.85, y, w: 4.7, h: 0.66, fontSize: 20, bold: true, color: FILL.teal, valign: "middle" });
    T(s, use, { x: 8.75, y, w: 3.85, h: 0.66, fontSize: 20, color: FILL.text, valign: "middle" });
  }
  T(s, "単純な優劣ではなく、必要な機能とコストの兼ね合いで使い分ける。", { x: 0.6, y: 6.05, w: 12.13, h: 0.5, fontSize: 22, bold: true, color: FILL.navy, valign: "middle" });
};

slides[16] = async () => {
  const s = addSlide(16);
  const col = async (x0, head, fill, inputs, img, proc, formula) => {
    pill(s, head, x0, 1.75, 5.92, 0.55, fill, FILL.white, 22);
    for (let i = 0; i < 2; i++) {
      const y = 2.75 + i * 0.95;
      box(s, x0, y, 2.0, 0.78, FILL.pale, { radius: 0.1 });
      await icon(s, inputs[i][1], x0 + 0.1, y + 0.2, 0.38, HEX.teal);
      T(s, inputs[i][0], { x: x0 + 0.53, y, w: 1.45, h: 0.78, fontSize: 15, bold: true, color: FILL.navy, valign: "middle" });
    }
    arrowR(s, x0 + 2.06, 3.42, 0.34, 0.38);
    s.addImage({ path: path.join(ASSETS, img + ".png"), x: x0 + 2.45, y: 2.5, w: 1.5, h: 2.0, altText: head + "の模式図（イメージ）" });
    arrowR(s, x0 + 4.03, 3.42, 0.34, 0.38);
    T(s, proc, { x: x0 + 4.42, y: 2.6, w: 1.5, h: 0.75, fontSize: 17, bold: true, color: FILL.teal, align: "center", valign: "bottom" });
    pill(s, "鋼材", x0 + 4.47, 3.4, 1.4, 0.65, FILL.navy, FILL.white, 22, { extra: { rectRadius: 0.12 } });
    T(s, formula, { x: x0, y: 4.72, w: 5.92, h: 0.5, fontSize: 17, color: FILL.text, align: "center", valign: "middle" });
  };
  await col(0.6, "高炉・転炉", FILL.navy, [["鉄鉱石", "TbMountain"], ["コークスなど", "TbFlame"]], "blast_furnace", "還元・精錬", "鉄鉱石＋コークスなど → 還元・精錬 → 鋼材");
  await col(6.81, "電炉", FILL.teal, [["鉄スクラップ", "TbRecycle"], ["電力", "TbBolt"]], "electric_furnace", "溶解・\n成分調整", "鉄スクラップ＋電力 → 溶解・成分調整 → 鋼材");
  line(s, 6.67, 2.5, 6.67, 5.1, { color: FILL.line, width: 1.5, dash: "dash" });
  box(s, 0.6, 5.45, 5.92, 0.95, FILL.pale, { radius: 0.12 });
  pill(s, "非鉄", 0.8, 5.67, 0.95, 0.5, FILL.orange, FILL.navy, 16);
  T(s, "銅・アルミは別の製錬プロセス", { x: 1.95, y: 5.45, w: 4.45, h: 0.95, fontSize: 20, bold: true, color: FILL.navy, valign: "middle" });
  box(s, 6.81, 5.45, 5.92, 0.95, FILL.navy, { radius: 0.12 });
  T(s, "原料・エネルギー・設備が違うと、\n利益の動き方も変わる。", { x: 7.05, y: 5.45, w: 5.5, h: 0.95, fontSize: 19, bold: true, color: FILL.white, valign: "middle" });
};

slides[17] = async () => {
  const s = addSlide(17);
  const st = [["資源開発", "TbPick"], ["製鉄・製錬", "TbFlame"], ["材料製造・加工", "TbTool"], ["流通", "TbTruck"], ["最終製品", "TbCar"]];
  const xs = st.map((_, i) => 1.5 + i * 2.58);
  // 回収・再生のループ（最終製品 → 製鉄・製錬）
  line(s, xs[4], 2.55, xs[4], 2.05, { color: FILL.teal, width: 3 });
  line(s, xs[4], 2.05, xs[1], 2.05, { color: FILL.teal, width: 3 });
  line(s, xs[1], 2.05, xs[1], 2.5, { color: FILL.teal, width: 3, arrow: true });
  s.addShape(SH.ROUNDED_RECTANGLE, { x: 4.75, y: 1.8, w: 6.05, h: 0.5, fill: { color: FILL.white }, line: { color: FILL.teal, width: 1.5 }, rectRadius: 0.25 });
  await icon(s, "TbRecycle", 4.92, 1.86, 0.38, HEX.teal);
  T(s, "使い終わった金属は回収され、再び原料になる。", { x: 5.38, y: 1.8, w: 5.35, h: 0.5, fontSize: 17, bold: true, color: FILL.teal, valign: "middle" });
  for (let i = 0; i < 5; i++) {
    const cx = xs[i];
    await iconCircle(s, st[i][1], cx - 0.55, 2.55, 1.1, i === 0 ? FILL.navy : i === 4 ? FILL.teal : "1C3A57", i === 4 ? HEX.white : HEX.cyan);
    T(s, st[i][0], { x: cx - 1.15, y: 3.72, w: 2.3, h: 0.45, fontSize: 19, bold: true, color: FILL.navy, align: "center" });
    if (i < 4) arrowR(s, cx + 0.72, 2.92, 1.12, 0.36, FILL.line);
  }
  // 商社が工程を横断する帯
  box(s, xs[1] - 1.1, 4.4, xs[4] - xs[1] + 2.2, 0.72, FILL.orange, { radius: 0.36 });
  await icon(s, "TbArrowsExchange", xs[1] - 0.9, 4.54, 0.44, HEX.navy);
  T(s, "商社は複数の工程をつなぎ、物流・加工・在庫管理を担う。", { x: xs[1] - 0.35, y: 4.4, w: 9.5, h: 0.72, fontSize: 19, bold: true, color: FILL.navy, valign: "middle" });
  box(s, 0.6, 5.5, 12.13, 0.8, FILL.pale, { radius: 0.14 });
  T(s, "一つの鋼材にも、多くの企業が役割と付加価値を分担している。", { x: 0.9, y: 5.5, w: 11.6, h: 0.8, fontSize: 22, bold: true, color: FILL.navy, valign: "middle" });
};

slides[18] = async () => {
  const s = addSlide(18);
  const cols = [["資源・製錬", ["住友金属鉱山", "三菱マテリアル"]], ["鉄鋼製造", ["日本製鉄", "JFEスチール"]], ["電炉", ["共英製鋼", "東京製鐵"]],
    ["特殊鋼・\n高機能材", ["大同特殊鋼", "JX金属"]], ["アルミ加工", ["UACJ"]], ["商社", ["阪和興業", "岡谷鋼機"]]];
  cols.forEach(([head, cos], i) => {
    const x = 0.6 + i * 2.0;
    s.addText(head, { shape: i === 0 ? SH.PENTAGON : SH.CHEVRON, x, y: 1.85, w: 2.13, h: 0.85, fill: { color: i === 5 ? FILL.orange : i % 2 ? FILL.teal : FILL.navy },
      color: i === 5 ? FILL.navy : FILL.white, bold: true, fontSize: 16, align: "center", valign: "middle", margin: 0, isTextBox: true, lang: "ja-JP", line: { type: "none" } });
    cos.forEach((c, k) => {
      s.addText(c, { shape: SH.ROUNDED_RECTANGLE, rectRadius: 0.1, x: x + 0.05, y: 2.95 + k * 0.85, w: 1.87, h: 0.7, fill: { color: FILL.white }, line: { color: FILL.line, width: 1.25 },
        color: FILL.navy, bold: true, fontSize: 17, align: "center", valign: "middle", margin: 0, isTextBox: true, lang: "ja-JP", shadow: shadow() });
    });
  });
  box(s, 0.6, 4.75, 12.13, 0.62, FILL.pale, { radius: 0.12 });
  await icon(s, "TbTopologyStar3", 0.8, 4.84, 0.44, HEX.teal);
  T(s, "1社が複数の工程・役割を持つこともある（例：住友金属鉱山＝資源・製錬・材料）", { x: 1.4, y: 4.75, w: 11.2, h: 0.62, fontSize: 18, color: FILL.navy, valign: "middle" });
  box(s, 0.6, 5.6, 12.13, 0.78, FILL.navy, { radius: 0.12 });
  T(s, [{ text: "名前を聞いたら → ", options: { color: FILL.cyan } }, { text: "何をどの工程で作り、誰に売るかを調べる" }], { x: 0.9, y: 5.6, w: 11.6, h: 0.78, fontSize: 22, bold: true, color: FILL.white, valign: "middle" });
  foot(s, "※社名は各工程の代表例（ロゴは使用せずテキストで表記）。");
};

slides[19] = async () => {
  const s = addSlide(19);
  await pic(s, "port", 0.6, 1.8, 6.2, 4.55);
  pill(s, "イメージ図", 0.8, 1.98, 1.4, 0.38, FILL.white, FILL.navy, 13);
  const cards = [["TbShip", "原材料を輸入する", "鉄鉱石・原料炭など"], ["TbWorld", "世界の顧客に販売する", "自動車・建設・電子分野"], ["TbBuildingFactory2", "海外で生産する", "現地企業や拠点と協力"], ["TbCurrencyYen", "為替・関税・需給の変化が影響する", ""]];
  for (let i = 0; i < 4; i++) {
    const y = 1.8 + i * 0.97, last = i === 3;
    box(s, 7.15, y, 5.58, 0.85, last ? "FCEBD3" : FILL.pale, { radius: 0.12 });
    await iconCircle(s, cards[i][0], 7.3, y + 0.12, 0.6, last ? FILL.orange : FILL.teal, last ? HEX.navy : HEX.white);
    if (cards[i][2]) {
      T(s, [{ text: cards[i][1], options: { bold: true, fontSize: 20, color: FILL.navy, breakLine: true } }, { text: cards[i][2], options: { fontSize: 16, color: FILL.gray } }],
        { x: 8.05, y, w: 4.6, h: 0.85, valign: "middle" });
    } else {
      T(s, cards[i][1], { x: 8.05, y, w: 4.6, h: 0.85, fontSize: 20, bold: true, color: FILL.navy, valign: "middle" });
    }
  }
  T(s, "国内工場だけでなく、\n国際的なサプライチェーンを見る。", { x: 7.15, y: 5.75, w: 5.58, h: 0.6, fontSize: 18, bold: true, color: FILL.teal, valign: "middle" });
};

slides[20] = async () => {
  const s = addSlide(20);
  const cols = [
    ["鉱山", "TbPick", [["売上", "金属価格 × 生産量"], ["費用", "採掘・開発費"]]],
    ["製鉄", "TbFlame", [["売上", "販売単価 × 数量"], ["費用", "原料・エネルギー・\n固定費"]]],
    ["加工", "TbSettings", [["売上", "材料価格＋加工賃"], ["価値", "品質・歩留まりが価値"]]],
    ["商社", "TbShip", [["価値", "取引・物流・加工・在庫管理による付加価値"]]],
  ];
  const tagFill = { 売上: FILL.teal, 費用: FILL.orange, 価値: FILL.navy };
  for (let i = 0; i < 4; i++) {
    const x = 0.6 + i * 3.09, [name, ico, parts] = cols[i];
    box(s, x, 1.8, 2.86, 1.0, i === 3 ? FILL.teal : FILL.navy, { radius: 0.14 });
    await icon(s, ico, x + 0.25, 2.03, 0.55, HEX.white);
    T(s, name, { x: x + 1.0, y: 1.8, w: 1.8, h: 1.0, fontSize: 26, bold: true, color: FILL.white, valign: "middle" });
    if (parts.length === 1) {
      box(s, x, 3.0, 2.86, 2.51, FILL.pale, { radius: 0.12 });
      pill(s, parts[0][0], x + 0.2, 3.15, 0.95, 0.42, tagFill[parts[0][0]], FILL.white, 15);
      T(s, parts[0][1], { x: x + 0.2, y: 3.7, w: 2.5, h: 1.5, fontSize: 19, bold: true, color: FILL.navy });
    } else {
      parts.forEach(([tag, txt], k) => {
        const y = 3.0 + k * 1.33;
        box(s, x, y, 2.86, 1.18, FILL.pale, { radius: 0.12 });
        pill(s, tag, x + 0.2, y + 0.12, 0.95, 0.4, tagFill[tag], tag === "費用" ? FILL.navy : FILL.white, 15);
        T(s, txt, { x: x + 0.2, y: y + 0.55, w: 2.55, h: 0.58, fontSize: 17, bold: true, color: FILL.navy, valign: "middle" });
      });
      T(s, parts[1][0] === "費用" ? "−" : "＋", { x: x + 1.2, y: 4.18, w: 0.5, h: 0.3, fontSize: 18, bold: true, color: FILL.gray, align: "center", valign: "middle" });
    }
  }
  box(s, 0.6, 5.72, 12.13, 0.66, FILL.pale, { radius: 0.12 });
  T(s, "ビジネスモデルが違えば、同じ金属価格のニュースでも受ける影響が違う。", { x: 0.9, y: 5.72, w: 11.6, h: 0.66, fontSize: 21, bold: true, color: FILL.navy, valign: "middle" });
  foot(s, "※加工の価格の考え方（地金価格＋加工賃）の例。出典：UACJ FAQ");
};

slides[21] = async () => {
  const s = addSlide(21);
  pill(s, "架空の計算例（単位：億円）", 0.6, 1.75, 3.9, 0.48, FILL.orange, FILL.navy, 18);
  const base = 5.55, k = 0.025, bw = 0.95;
  const seg = [["原材料費", FILL.navy, FILL.white], ["その他変動費", FILL.teal, FILL.white], ["固定費", FILL.cyan, FILL.navy], ["営業利益", FILL.red, FILL.white]];
  // 凡例（積み上げの順に上から）
  [3, 2, 1, 0].forEach((idx, i) => {
    const y = 2.75 + i * 0.45;
    box(s, 6.25, y + 0.06, 0.24, 0.24, seg[idx][1], { radius: 0.03 });
    T(s, seg[idx][0], { x: 6.55, y, w: 1.17, h: 0.36, fontSize: 13, color: FILL.text, valign: "middle" });
  });
  box(s, 6.25, 4.6, 0.24, 0.24, FILL.gray, { radius: 0.03 });
  T(s, "売上", { x: 6.55, y: 4.54, w: 1.1, h: 0.36, fontSize: 13, color: FILL.text, valign: "middle" });
  const period = (x, sales, parts, label) => {
    box(s, x, base - sales * k, bw, sales * k, FILL.gray, { rect: true });
    T(s, String(sales), { x, y: base - sales * k + 0.08, w: bw, h: 0.4, fontSize: 18, bold: true, color: FILL.white, align: "center" });
    let y = base;
    parts.forEach((v, i) => {
      const h = v * k;
      y -= h;
      box(s, x + 1.08, y, bw, h, seg[i][1], { rect: true });
      if (i < 3) T(s, String(v), { x: x + 1.08, y, w: bw, h, fontSize: 16, bold: true, color: seg[i][2], align: "center", valign: "middle" });
    });
    T(s, "営業利益 " + parts[3], { x: x + 0.78, y: y - 0.42, w: 1.55, h: 0.38, fontSize: 17, bold: true, color: FILL.red, align: "center", valign: "middle" });
    T(s, "売上", { x, y: base + 0.06, w: bw, h: 0.3, fontSize: 14, color: FILL.gray, align: "center" });
    T(s, "費用＋利益", { x: x + 1.0, y: base + 0.06, w: 1.1, h: 0.3, fontSize: 14, color: FILL.gray, align: "center" });
    T(s, label, { x, y: base + 0.38, w: 2.03, h: 0.4, fontSize: 20, bold: true, color: FILL.navy, align: "center" });
  };
  line(s, 0.6, base, 6.15, base, { color: FILL.gray, width: 1.25 });
  period(0.85, 100, [60, 20, 15, 5], "前期");
  period(3.65, 110, [72, 20, 15, 3], "今期");
  // 数値表（編集可能な表）
  const hd = (t) => ({ text: t, options: { bold: true, color: HEX.white, fill: { color: HEX.navy }, align: "center" } });
  const rows = [[hd("項目"), hd("前期"), hd("今期"), hd("増減")],
    ["売上", "100", "110", "＋10％"], ["原材料費", "60", "72", "＋12"], ["その他変動費", "20", "20", "±0"], ["固定費", "15", "15", "±0"]].map((r, i) =>
    i === 0 ? r : r.map((c, j) => ({ text: c, options: { align: j === 0 ? "left" : "center", fill: { color: i % 2 ? HEX.white : HEX.pale }, color: HEX.text } })));
  rows.push(["営業利益", "5", "3", "−40％"].map((c, j) => ({ text: c, options: { bold: true, color: HEX.red, fill: { color: HEX.redPale }, align: j === 0 ? "left" : "center" } })));
  s.addTable(rows, { x: 7.75, y: 1.8, w: 4.98, colW: [1.88, 0.9, 0.9, 1.3], rowH: 0.5, fontSize: 18, fontFace: FONT, border: { type: "solid", pt: 0.75, color: HEX.line }, margin: [0.04, 0.12, 0.04, 0.12], valign: "middle" });
  box(s, 7.75, 5.05, 4.98, 1.25, FILL.navy, { radius: 0.14 });
  T(s, "売上が増えても、原料費の増加に利益が押される。", { x: 8.0, y: 5.05, w: 4.55, h: 1.25, fontSize: 21, bold: true, color: FILL.white, valign: "middle" });
};

slides[22] = async () => {
  const s = addSlide(22);
  s.addShape(SH.ISOSCELES_TRIANGLE, { x: 2.4, y: 2.48, w: 4.1, h: 2.9, fill: { color: FILL.white }, line: { color: FILL.cyan, width: 2.5, dashType: "dash" } });
  const nodes = [[2.65, 1.85, "1", "何を扱う？", "鉄・銅・アルミ・特殊材料"], [0.6, 4.75, "2", "どの工程にいる？", "資源・製造・加工・商社"], [4.7, 4.75, "3", "利益を動かすのは？", "価格・数量・原料・技術・為替"]];
  nodes.forEach(([x, y, no, q, a]) => {
    box(s, x, y, 3.6, 1.25, FILL.white, { lineColor: FILL.teal, lineWidth: 1.75, shadow: true, radius: 0.16 });
    badge(s, no, x + 0.18, y + 0.18, 0.55, FILL.teal, FILL.white, 20);
    T(s, q, { x: x + 0.88, y: y + 0.12, w: 2.65, h: 0.48, fontSize: 20, bold: true, color: FILL.navy, valign: "middle" });
    T(s, a, { x: x + 0.88, y: y + 0.6, w: 2.65, h: 0.58, fontSize: 15, color: FILL.teal, bold: true });
  });
  s.addText([{ text: "企業名の例", options: { fontSize: 13, color: FILL.gray, breakLine: true } }, { text: "日本製鉄／UACJ／阪和興業", options: { fontSize: 15, bold: true, color: FILL.navy } }],
    { shape: SH.ROUNDED_RECTANGLE, rectRadius: 0.1, x: 3.0, y: 3.62, w: 2.9, h: 0.75, fill: { color: FILL.pale }, align: "center", valign: "middle", margin: 0, isTextBox: true, lang: "ja-JP", line: { type: "none" } });
  box(s, 8.75, 1.85, 3.98, 4.5, FILL.navy, { shadow: true, radius: 0.2 });
  T(s, "NEXT", { x: 9.05, y: 2.1, w: 3.4, h: 0.4, fontSize: 16, bold: true, color: FILL.cyan });
  T(s, "この3つを持って、\n次はニュースを\n読んでみよう。", { x: 9.05, y: 2.55, w: 3.45, h: 1.65, fontSize: 24, bold: true, color: FILL.white });
  for (let i = 0; i < 3; i++) {
    await iconCircle(s, "TbNews", 9.05 + i * 1.1, 4.55, 0.85, "1C3A57", HEX.cyan);
    T(s, "①②③"[i], { x: 9.05 + i * 1.1, y: 5.45, w: 0.85, h: 0.4, fontSize: 18, bold: true, color: FILL.white, align: "center" });
  }
};

slides[23] = async () => {
  const s = addSlide(23);
  await pic(s, "coil", 0.6, 1.8, 4.85, 4.5);
  pill(s, "イメージ図", 0.8, 1.98, 1.4, 0.38, FILL.white, FILL.navy, 13);
  pill(s, "NEWS ①", 5.8, 1.85, 1.45, 0.48, FILL.orange, FILL.navy, 16);
  T(s, "日本製鉄｜2026年10月", { x: 7.4, y: 1.8, w: 5.3, h: 0.58, fontSize: 24, bold: true, color: FILL.navy, valign: "middle" });
  const blk = async (x, label, val, col, ico, pale) => {
    box(s, x, 2.6, 3.33, 2.4, pale, { radius: 0.16 });
    T(s, label, { x: x + 0.2, y: 2.72, w: 2.95, h: 0.75, fontSize: 18, bold: true, color: FILL.navy, valign: "middle" });
    await icon(s, ico, x + 0.2, 3.62, 0.62, col === FILL.teal ? HEX.teal : "B86E10");
    T(s, val, { x: x + 0.85, y: 3.5, w: 2.4, h: 0.85, fontSize: 40, bold: true, color: col === FILL.teal ? FILL.teal : "B86E10", valign: "middle" });
    T(s, "円／トン", { x: x + 0.85, y: 4.35, w: 2.4, h: 0.45, fontSize: 18, bold: true, color: FILL.text });
  };
  await blk(5.8, "一部の薄板4品種", "+5,000", FILL.orange, "TbTrendingUp", "FCEBD3");
  await blk(9.4, "一部の店売り向け\nステンレス鋼板", "-5,000", FILL.teal, "TbTrendingDown", "D6ECEC");
  box(s, 5.8, 5.2, 6.93, 1.1, FILL.navy, { radius: 0.14 });
  T(s, "同じ会社でも、\n製品ごとに価格が違う方向へ動いた。", { x: 6.05, y: 5.2, w: 6.5, h: 1.1, fontSize: 21, bold: true, color: FILL.white, valign: "middle" });
  foot(s, "※両製品は同じ価格体系の製品ではない。出典：日刊産業新聞（2026/9/30、2026/10/9）", "値上げ対象の範囲・ステンレス値下げ報道");
};

slides[24] = async () => {
  const s = addSlide(24);
  await pic(s, "hydrogen", 0.6, 1.8, 6.95, 4.5);
  pill(s, "イメージ図", 0.8, 1.98, 1.4, 0.38, FILL.white, FILL.navy, 13);
  pill(s, "NEWS ②", 7.9, 1.85, 1.45, 0.48, FILL.orange, FILL.navy, 16);
  T(s, "2026年7月31日\nNEDO・JFEスチール等が発表", { x: 7.9, y: 2.45, w: 4.83, h: 0.75, fontSize: 18, bold: true, color: FILL.gray });
  const facts = [["TbFlask", "日本で初めて、グリーン水素を使った還元鉄試作に成功", true], ["TbScale", "試験設備で約1トンの還元鉄を製造", false], ["TbRoute", "実用化へ向けた研究・実証は続く。", false]];
  for (let i = 0; i < 3; i++) {
    const y = 3.3 + i * 1.02;
    box(s, 7.9, y, 4.83, 0.9, i === 0 ? FILL.teal : FILL.pale, { radius: 0.12 });
    await icon(s, facts[i][0], 8.08, y + 0.22, 0.46, i === 0 ? HEX.white : HEX.teal);
    T(s, facts[i][1], { x: 8.65, y, w: 4.0, h: 0.9, fontSize: 17, bold: true, color: i === 0 ? FILL.white : FILL.navy, valign: "middle" });
  }
  foot(s, "※試作段階の成果であり、商業生産の確立ではない。出典：NEDO（2026/7/31）", "製造量「約1トン」");
};

slides[25] = async () => {
  const s = addSlide(25);
  box(s, 0.6, 1.8, 7.5, 4.55, FILL.white, { lineColor: FILL.line, radius: 0.16 });
  T(s, "2027年の鉄鋼需要予測（前年比）", { x: 0.85, y: 1.92, w: 7.0, h: 0.45, fontSize: 18, bold: true, color: FILL.navy });
  s.addChart(pres.charts.BAR, [{ name: "2027年 鉄鋼需要（前年比・予測）", labels: ["中国", "世界", "日本", "インド"], values: [-1.0, 1.7, 1.7, 8.0] }], {
    x: 0.8, y: 2.4, w: 7.1, h: 3.8, barDir: "bar", chartColors: [HEX.red, HEX.gray, HEX.teal, HEX.orange], invertIfNegative: false,
    showValue: true, dataLabelFormatCode: '+0.0"%";-0.0"%"', dataLabelFontSize: 20, dataLabelFontBold: true, dataLabelColor: HEX.navy, dataLabelPosition: "outEnd", dataLabelFontFace: FONT,
    valAxisHidden: true, valAxisMinVal: -3, valAxisMaxVal: 10.5, valGridLine: { style: "none" }, catGridLine: { style: "none" },
    catAxisLabelFontSize: 20, catAxisLabelColor: HEX.navy, catAxisLabelFontFace: FONT, catAxisLabelPos: "low", showLegend: false, barGapWidthPct: 45,
  });
  pill(s, "NEWS ③", 8.5, 1.85, 1.45, 0.48, FILL.orange, FILL.navy, 16);
  T(s, "世界鉄鋼協会\n2026年10月12日公表", { x: 8.5, y: 2.5, w: 4.23, h: 0.9, fontSize: 20, bold: true, color: FILL.navy });
  const kv = [["インド", "＋8.0％", "B86E10"], ["中国", "−1.0％", HEX.red], ["日本", "＋1.7％", HEX.teal], ["世界", "＋1.7％", HEX.gray]];
  kv.forEach(([k, v, c], i) => {
    const x = 8.5 + (i % 2) * 2.15, y = 3.55 + Math.floor(i / 2) * 0.62;
    T(s, [{ text: k + " ", options: { color: FILL.text, fontSize: 18 } }, { text: v, options: { color: c, bold: true, fontSize: 22 } }], { x, y, w: 2.1, h: 0.55, valign: "middle" });
  });
  box(s, 8.5, 4.95, 4.23, 1.4, "FCEBD3", { radius: 0.14 });
  T(s, "※2027年の実績ではなく、\n2026年時点の「予測」。", { x: 8.7, y: 4.95, w: 3.9, h: 1.4, fontSize: 18, bold: true, color: FILL.navy, valign: "middle" });
  foot(s, "※示しているのは鉄鋼需要であり、金属産業全体の実績や企業利益ではない。出典：worldsteel Short Range Outlook（2026/10/12）");
};

slides[26] = async () => {
  const s = addSlide(26, "WORK");
  box(s, 0.6, 1.85, 9.0, 1.75, FILL.white, { shadow: true, radius: 0.2 });
  pill(s, "Q", 0.85, 2.38, 0.7, 0.7, FILL.orange, FILL.navy, 26);
  T(s, "日本製鉄は、なぜ一部の製品を値上げし、別の製品を値下げした？", { x: 1.8, y: 1.85, w: 7.6, h: 1.75, fontSize: 26, bold: true, color: FILL.navy, valign: "middle" });
  timer(s, "4:00", 10.4, 1.7, 2.1);
  T(s, "ヒント", { x: 0.6, y: 3.92, w: 1.0, h: 0.4, fontSize: 18, bold: true, color: FILL.teal });
  const hints = [["TbPackage", "原材料"], ["TbChartBar", "需要"], ["TbCoin", "コスト"], ["TbCurrencyYen", "為替"], ["TbFileText", "契約条件"]];
  for (let i = 0; i < 5; i++) {
    const x = 1.7 + i * 1.6;
    box(s, x, 3.82, 1.45, 1.2, FILL.white, { radius: 0.14, lineColor: FILL.line });
    await icon(s, hints[i][0], x + 0.5, 3.92, 0.45, HEX.teal);
    T(s, hints[i][1], { x, y: 4.45, w: 1.45, h: 0.45, fontSize: 16, bold: true, color: FILL.navy, align: "center" });
  }
  T(s, "進め方", { x: 9.85, y: 3.85, w: 2.9, h: 0.38, fontSize: 16, bold: true, color: FILL.teal });
  ["自分の仮説を1つ", "班で共有", "必要な追加情報を考える"].forEach((t, i) => {
    T(s, `${i + 1}. ${t}`, { x: 9.85, y: 4.25 + i * 0.4, w: 2.9, h: 0.36, fontSize: 15, bold: true, color: FILL.navy, valign: "middle" });
  });
  timeline(s, [["個人 30秒", 30], ["討論 2分30秒", 150], ["共有 1分", 60]], 5.75);
};

slides[27] = async () => {
  const s = addSlide(27);
  const card = async (y, tag, tagFill, name, reason, pale, ico, hex) => {
    box(s, 0.6, y, 5.25, 1.88, pale, { radius: 0.16 });
    pill(s, tag, 0.82, y + 0.2, 1.15, 0.45, tagFill, FILL.navy, 16);
    await icon(s, ico, 2.12, y + 0.18, 0.5, hex);
    T(s, name, { x: 2.75, y: y + 0.15, w: 3.0, h: 0.55, fontSize: 22, bold: true, color: FILL.navy, valign: "middle" });
    T(s, reason, { x: 0.82, y: y + 0.8, w: 4.85, h: 0.95, fontSize: 18, color: FILL.navy });
  };
  await card(1.85, "値上げ", FILL.orange, "薄板", "原料炭・人件費・物流費などの上昇が背景", "FCEBD3", "TbTrendingUp", "B86E10");
  await card(3.93, "値下げ", FILL.cyan, "ステンレス", "円高などで合金原料の調達価格が下落", "D6ECEC", "TbTrendingDown", HEX.teal);
  s.addText("原材料・\n契約が\n異なる", { shape: SH.ROUNDED_RECTANGLE, rectRadius: 0.14, x: 6.05, y: 3.12, w: 1.3, h: 1.5, fill: { color: FILL.navy }, color: FILL.white, bold: true, fontSize: 16, align: "center", valign: "middle", margin: 0, isTextBox: true, lang: "ja-JP" });
  line(s, 5.85, 2.8, 6.05, 3.4, { color: FILL.navy, width: 1.5 });
  line(s, 5.85, 4.88, 6.05, 4.35, { color: FILL.navy, width: 1.5 });
  T(s, "製品ごとの価格を動かす4つの要因", { x: 7.6, y: 1.8, w: 5.13, h: 0.45, fontSize: 20, bold: true, color: FILL.navy });
  const fs4 = [["TbPackage", "原材料価格"], ["TbChartBar", "需給"], ["TbTruck", "製造・物流費"], ["TbCurrencyYen", "為替・契約条件"]];
  for (let i = 0; i < 4; i++) {
    const x = 7.6 + (i % 2) * 2.62, y = 2.35 + Math.floor(i / 2) * 1.75;
    box(s, x, y, 2.48, 1.6, FILL.white, { lineColor: FILL.line, shadow: true, radius: 0.14 });
    await iconCircle(s, fs4[i][0], x + 0.2, y + 0.2, 0.65, FILL.pale, HEX.teal);
    T(s, "①②③④"[i], { x: x + 1.0, y: y + 0.25, w: 0.5, h: 0.5, fontSize: 22, bold: true, color: FILL.teal, valign: "middle" });
    T(s, fs4[i][1], { x: x + 0.2, y: y + 0.95, w: 2.15, h: 0.5, fontSize: 19, bold: true, color: FILL.navy, valign: "middle" });
  }
  T(s, [{ text: "価格が上がる＝利益が増える、", options: { color: FILL.navy } }, { text: "ではない。", options: { color: "B86E10" } }], { x: 0.6, y: 5.92, w: 12.13, h: 0.5, fontSize: 24, bold: true, valign: "middle" });
  foot(s, "出典：日刊産業新聞（2026/9/30、2026/10/9）", "ステンレス値下げの理由");
};

slides[28] = async () => {
  const s = addSlide(28, "WORK");
  T(s, "あなたが鉄鋼メーカーの経営者なら？", { x: 0.6, y: 1.8, w: 9.3, h: 0.6, fontSize: 26, bold: true, color: FILL.navy, valign: "middle" });
  timer(s, "5:00", 10.4, 1.7, 2.1);
  const opt = (x, fill, letter, text) => {
    box(s, x, 2.6, 4.45, 2.05, fill, { shadow: true, radius: 0.2 });
    T(s, letter, { x: x + 0.3, y: 2.72, w: 1, h: 0.75, fontSize: 40, bold: true, color: FILL.orange });
    T(s, text, { x: x + 0.3, y: 3.45, w: 3.9, h: 1.1, fontSize: 21, bold: true, color: FILL.white });
  };
  opt(0.6, FILL.teal, "A", "今から水素還元などの設備に投資する");
  opt(5.45, FILL.navy, "B", "まず既存設備の利益・投資回収を優先する");
  badge(s, "or", 4.76, 3.3, 0.88, FILL.orange, FILL.navy, 22);
  const ctx = [["TbBriefcase", "経営者の立場"], ["TbCoin", "資金"], ["TbLeaf", "排出規制"]];
  for (let i = 0; i < 3; i++) {
    const y = 4.0 + i * 0.5;
    await icon(s, ctx[i][0], 10.4, y, 0.38, HEX.teal);
    T(s, ctx[i][1], { x: 10.9, y, w: 1.9, h: 0.38, fontSize: 16, bold: true, color: FILL.navy, valign: "middle" });
  }
  T(s, "「どんな条件なら判断が変わる？」", { x: 0.6, y: 4.88, w: 9.3, h: 0.55, fontSize: 24, bold: true, color: "B86E10", valign: "middle" });
  timeline(s, [["個人 30秒", 30], ["討論 3分", 180], ["全体共有 1分", 60]], 5.85);
};

slides[29] = async () => {
  const s = addSlide(29);
  const pan = async (x, fill, head, ico, hex, items, headColor) => {
    box(s, x, 1.8, 5.5, 2.1, fill, { radius: 0.18 });
    await icon(s, ico, x + 0.25, 1.98, 0.5, hex);
    T(s, head, { x: x + 0.85, y: 1.95, w: 4.4, h: 0.55, fontSize: 22, bold: true, color: headColor, valign: "middle" });
    items.forEach((t, i) => pill(s, t, x + 0.25 + (i % 2) * 2.55, 2.68 + Math.floor(i / 2) * 0.6, 2.45, 0.5, FILL.white, FILL.navy, 16));
  };
  await pan(0.6, "FCEBD3", "短期的な負担", "TbCoin", "B86E10", ["設備投資", "研究費", "水素／電力の調達"], FILL.navy);
  await pan(7.23, "D6ECEC", "長期の可能性", "TbTrendingUp", HEX.teal, ["規制対応", "新たな需要", "技術競争力"], FILL.navy);
  // 天秤（てんびん）
  box(s, 1.2, 3.98, 10.93, 0.12, FILL.navy, { rect: true });
  s.addShape(SH.ISOSCELES_TRIANGLE, { x: 6.22, y: 4.1, w: 0.9, h: 0.65, fill: { color: FILL.navy }, line: { type: "none" } });
  T(s, "必要な条件", { x: 0.6, y: 4.95, w: 1.7, h: 0.55, fontSize: 18, bold: true, color: FILL.teal, valign: "middle" });
  ["安価な水素", "製造品質", "量産", "投資回収"].forEach((t, i) => pill(s, t, 2.4 + i * 2.6, 4.95, 2.4, 0.55, FILL.white, FILL.navy, 18, { lineColor: FILL.teal }));
  box(s, 0.6, 5.7, 12.13, 0.68, FILL.navy, { radius: 0.12 });
  T(s, "経営は「いまの利益」と「将来の競争力」を両方見る。", { x: 0.9, y: 5.7, w: 11.6, h: 0.68, fontSize: 22, bold: true, color: FILL.white, valign: "middle" });
  foot(s, "※NEDOの成果は試作段階。事業化には技術・コスト面の課題がある。出典：NEDO（2026/7/31）");
};

slides[30] = async () => {
  const s = addSlide(30, "WORK");
  T(s, "中国の需要鈍化・インドの需要増加を踏まえて考えよう。", { x: 0.6, y: 1.8, w: 9.5, h: 0.55, fontSize: 22, bold: true, color: FILL.navy, valign: "middle" });
  [["インド ＋8.0％", FILL.orange, FILL.navy], ["中国 −1.0％", FILL.red, FILL.white], ["日本 ＋1.7％", FILL.teal, FILL.white]].forEach(([t, f, c], i) => pill(s, t, 0.6 + i * 2.2, 2.45, 2.05, 0.48, f, c, 17));
  T(s, "2027年予測・前年比", { x: 7.25, y: 2.45, w: 2.6, h: 0.48, fontSize: 14, color: FILL.gray, valign: "middle" });
  timer(s, "4:00", 10.4, 1.7, 2.1);
  const st = [["A", "TbWorld", "海外市場への展開"], ["B", "TbAward", "高機能・高品質材料"], ["C", "TbLeaf", "環境・新技術"]];
  for (let i = 0; i < 3; i++) {
    const x = 0.6 + i * 3.2;
    box(s, x, 3.15, 3.0, 1.55, FILL.white, { shadow: true, radius: 0.18 });
    T(s, st[i][0], { x: x + 0.22, y: 3.25, w: 0.7, h: 0.65, fontSize: 32, bold: true, color: FILL.orange });
    await icon(s, st[i][1], x + 2.2, 3.3, 0.55, HEX.teal);
    T(s, st[i][2], { x: x + 0.22, y: 3.95, w: 2.6, h: 0.6, fontSize: 20, bold: true, color: FILL.navy, valign: "middle" });
  }
  T(s, "選んだ戦略の「根拠」と「リスク」を1つずつ", { x: 0.6, y: 4.95, w: 12.13, h: 0.55, fontSize: 22, bold: true, color: "B86E10", valign: "middle" });
  timeline(s, [["個人 30秒", 30], ["討論 2分30秒", 150], ["共有 1分", 60]], 5.85);
};

slides[31] = async () => {
  const s = addSlide(31);
  const icons = ["TbWorld", "TbAward", "TbLeaf"];
  for (let i = 0; i < 3; i++) await iconCircle(s, icons[i], 0.65, 2.52 + i * 1.0, 0.72, FILL.pale, HEX.teal);
  const hd = (t, f, c) => ({ text: t, options: { bold: true, fill: { color: f }, color: c, align: "center" } });
  const body = [["海外展開", "市場成長", "現地競合・関税・物流"], ["高機能材料", "品質で差別化", "開発費・顧客評価"], ["脱炭素技術", "新市場・規制対応", "巨額投資・採算"]];
  const rows = [[hd("戦略", HEX.navy, HEX.white), hd("機会", HEX.teal, HEX.white), hd("課題", HEX.orange, HEX.navy)]].concat(
    body.map((r) => r.map((c, j) => ({ text: c, options: { bold: j === 0, color: HEX.navy, fill: { color: j === 1 ? "E6F3F3" : j === 2 ? "FDF3E6" : HEX.white } } }))));
  s.addTable(rows, { x: 1.6, y: 1.85, w: 11.13, colW: [2.45, 4.34, 4.34], rowH: [0.55, 1.0, 1.0, 1.0], fontSize: 20, fontFace: FONT, valign: "middle", border: { type: "solid", pt: 1, color: HEX.line }, margin: [0.04, 0.18, 0.04, 0.18] });
  box(s, 0.6, 5.65, 12.13, 0.72, FILL.navy, { radius: 0.12 });
  T(s, "市場が伸びても、すべての日本企業が儲かるわけではない。", { x: 0.9, y: 5.65, w: 11.6, h: 0.72, fontSize: 22, bold: true, color: FILL.white, valign: "middle" });
  foot(s, "出典：worldsteel Short Range Outlook（2026/10/12）");
};

slides[32] = async () => {
  const s = addSlide(32);
  const rows = [["ニュース①", "価格と費用", "利益構造", "TbCurrencyYen"], ["ニュース②", "いまの利益と未来", "経営判断", "TbScale"], ["ニュース③", "海外需要と自社の強み", "経営戦略", "TbWorld"]];
  for (let i = 0; i < 3; i++) {
    const y = 1.9 + i * 1.3, [n, sub, key, ico] = rows[i];
    pill(s, n, 0.6, y + 0.18, 2.1, 0.7, FILL.navy, FILL.white, 20, { extra: { rectRadius: 0.14 } });
    box(s, 2.9, y, 5.3, 1.05, FILL.pale, { radius: 0.14 });
    await icon(s, ico, 3.1, y + 0.28, 0.5, HEX.teal);
    T(s, sub, { x: 3.8, y, w: 4.3, h: 1.05, fontSize: 22, bold: true, color: FILL.navy, valign: "middle" });
    arrowR(s, 8.4, y + 0.3, 0.75, 0.45, FILL.orange);
    pill(s, key, 9.35, y + 0.08, 3.38, 0.9, FILL.teal, FILL.white, 26, { extra: { rectRadius: 0.16 } });
  }
  box(s, 0.6, 5.85, 12.13, 0.65, FILL.white, { lineColor: FILL.teal, lineWidth: 1.75, radius: 0.12 });
  T(s, "1つの金属業界にも、数字・技術・世界市場という面白さがある。", { x: 0.9, y: 5.85, w: 11.6, h: 0.65, fontSize: 22, bold: true, color: FILL.navy, valign: "middle" });
};

slides[33] = async () => {
  const s = addSlide(33);
  const p = [["city_bridge", "社会基盤を支える", "橋・車・電力・\n電子機器"], ["chip", "技術で挑戦する", "高機能材料・\n製造技術"], ["port", "世界とつながる", "資源・生産・輸出"],
    ["hydrogen", "未来の課題に\n向き合う", "脱炭素・資源循環"], ["discussion", "多様な仕事がある", "研究・生産・営業・\n調達・企画"]];
  for (let i = 0; i < 5; i++) {
    const x = 0.6 + i * 2.476;
    await pic(s, p[i][0], x, 1.85, 2.23, 1.8, { pos: i === 4 ? "south" : "centre" });
    T(s, p[i][1], { x, y: 3.82, w: 2.23, h: 0.8, fontSize: 20, bold: true, color: FILL.navy, valign: "middle" });
    T(s, p[i][2], { x, y: 4.65, w: 2.23, h: 0.85, fontSize: 17, bold: true, color: FILL.teal });
  }
  box(s, 0.6, 5.75, 12.13, 0.7, FILL.pale, { radius: 0.12 });
  await icon(s, "TbHandClick", 0.85, 5.86, 0.48, HEX.orange);
  T(s, "あなたが一番惹かれた入口は、どれ？", { x: 1.5, y: 5.75, w: 11.0, h: 0.7, fontSize: 22, bold: true, color: FILL.navy, valign: "middle" });
};

slides[34] = async () => {
  const s = addSlide(34);
  T(s, "金属業界を知らない人に、何を見せたら興味が生まれる？", { x: 0.6, y: 1.75, w: 12.13, h: 0.6, fontSize: 24, bold: true, color: FILL.navy, valign: "middle" });
  const stage = (t, x, w) => T(s, t, { x, y: 2.5, w, h: 0.38, fontSize: 16, bold: true, color: FILL.teal, align: "center" });
  stage("知らない", 0.6, 2.3); stage("身近な接点", 3.35, 2.2); stage("驚く体験", 5.95, 4.6); stage("もっと調べたい", 10.6, 2.13);
  s.addText("金属って\nよく分からない", { shape: SH.ROUNDED_RECTANGULAR_CALLOUT, x: 0.6, y: 2.95, w: 2.3, h: 0.85, fill: { color: FILL.pale }, line: { color: FILL.line, width: 1 }, color: FILL.navy, bold: true, fontSize: 16, align: "center", valign: "middle", margin: 0, isTextBox: true, lang: "ja-JP" });
  await iconCircle(s, "TbUser", 1.2, 4.05, 1.1, FILL.pale, HEX.navy);
  const ex = [["chip", 3.35, "スマホの中の\n金属を見る"], ["steel_mill", 5.95, "工場で巨大な\n鋼板ができる"]];
  for (const [img, x, cap] of ex) {
    await pic(s, img, x, 2.95, 2.2, 1.65);
    T(s, cap, { x, y: 4.7, w: 2.2, h: 0.7, fontSize: 16, bold: true, color: FILL.navy, align: "center" });
  }
  box(s, 8.35, 2.95, 2.2, 1.65, FILL.navy, { radius: 0.14 });
  await icon(s, "TbMicroscope", 9.0, 3.3, 0.9, HEX.cyan);
  T(s, "研究者の開発\nストーリーを聞く", { x: 8.35, y: 4.7, w: 2.2, h: 0.7, fontSize: 16, bold: true, color: FILL.navy, align: "center" });
  arrowR(s, 2.95, 3.58, 0.35, 0.4, FILL.line);
  arrowR(s, 5.58, 3.58, 0.34, 0.4, FILL.line);
  box(s, 10.75, 2.95, 1.98, 1.65, FILL.orange, { shadow: true, radius: 0.14 });
  await icon(s, "TbSparkles", 11.44, 3.12, 0.6, HEX.navy);
  T(s, "もっと\n知りたい！", { x: 10.75, y: 3.75, w: 1.98, h: 0.8, fontSize: 20, bold: true, color: FILL.navy, align: "center", valign: "middle" });
  box(s, 0.6, 5.55, 12.13, 0.85, FILL.navy, { radius: 0.14 });
  T(s, [{ text: "大切なのは" }, { text: "「誰に」「何を」「どう見せるか」", options: { color: FILL.orange } }, { text: "。" }],
    { x: 0.9, y: 5.55, w: 11.6, h: 0.85, fontSize: 26, bold: true, color: FILL.white, valign: "middle", align: "center" });
};

slides[35] = async () => {
  const s = addSlide(35, "HERO_GD");
  await pic(s, "discussion", 7.05, 0, 6.28, 7.5, { radius: 0 });
  pill(s, "MISSION", 0.6, 3.15, 1.55, 0.45, FILL.orange, FILL.navy, 16);
  T(s, "金属業界の面白さを学生に伝える企画をつくる", { x: 0.6, y: 3.7, w: 6.15, h: 0.85, fontSize: 21, bold: true, color: FILL.white });
  T(s, [{ text: "ゴール：", options: { color: FILL.orange } }, { text: "就職させることではなく、「調べてみたい」と思ってもらうこと" }],
    { x: 0.6, y: 4.6, w: 6.15, h: 0.85, fontSize: 18, bold: true, color: FILL.cyan });
  ["誰に伝える？", "どんな魅力？", "何を体験\nしてもらう？"].forEach((t, i) => pill(s, t, 0.6 + i * 2.08, 5.65, 1.95, 0.85, FILL.navy, FILL.white, 16, { lineColor: FILL.cyan, extra: { rectRadius: 0.14 } }));
};

slides[36] = async () => {
  const s = addSlide(36);
  const steps = [["対象を決める", 3, "TbTarget", "ターゲット"], ["伝える魅力を\n選ぶ", 3, "TbHeart", "魅力"], ["伝える方法を\n決める", 4, "TbSpeakerphone", "媒体"], ["具体的な体験を\n設計", 3, "TbHandClick", "体験"], ["発表の準備", 1, "TbMessages", "ひとこと"]];
  // 時間に比例したタイムバー
  let x = 0.6, t = 0;
  const unit = 12.13 / 14;
  const fills = [FILL.navy, FILL.teal, FILL.navy, FILL.teal, FILL.orange];
  T(s, "0", { x: 0.55, y: 2.48, w: 0.4, h: 0.3, fontSize: 14, color: FILL.gray });
  steps.forEach(([, min], i) => {
    const w = unit * min;
    s.addText(min + "分", { shape: SH.RECTANGLE, x, y: 1.9, w: w - 0.04, h: 0.52, fill: { color: fills[i] }, color: i === 4 ? FILL.navy : FILL.white, bold: true, fontSize: 16, align: "center", valign: "middle", margin: 0, isTextBox: true, lang: "ja-JP" });
    x += w; t += min;
    T(s, String(t) + (i === 4 ? "分" : ""), { x: x - 0.7, y: 2.48, w: 0.7, h: 0.3, fontSize: 14, color: FILL.gray, align: "right" });
  });
  for (let i = 0; i < 5; i++) {
    const cx = 0.6 + i * 2.465, [name, min, ico, out] = steps[i];
    box(s, cx, 3.0, 2.27, 2.25, FILL.white, { lineColor: FILL.line, shadow: true, radius: 0.16 });
    badge(s, "①②③④⑤"[i], cx + 0.15, 3.14, 0.55, fills[i], i === 4 ? FILL.navy : FILL.white, 18);
    T(s, `［${min}分］`, { x: cx + 0.8, y: 3.14, w: 1.35, h: 0.55, fontSize: 16, bold: true, color: FILL.teal, valign: "middle" });
    T(s, name, { x: cx + 0.15, y: 3.8, w: 2.0, h: 0.75, fontSize: 19, bold: true, color: FILL.navy });
    box(s, cx + 0.15, 4.65, 1.97, 0.45, FILL.pale, { radius: 0.22 });
    await icon(s, ico, cx + 0.27, 4.71, 0.33, HEX.teal);
    T(s, out, { x: cx + 0.68, y: 4.65, w: 1.4, h: 0.45, fontSize: 15, bold: true, color: FILL.navy, valign: "middle" });
  }
  box(s, 0.6, 5.55, 9.4, 0.8, FILL.pale, { radius: 0.12 });
  T(s, "全員がひとことは発言し、仮決定しながら進もう。", { x: 0.9, y: 5.55, w: 9.0, h: 0.8, fontSize: 22, bold: true, color: FILL.navy, valign: "middle" });
  s.addText([{ text: "14:00", options: { fontSize: 34, bold: true, breakLine: true } }, { text: "TIMER", options: { fontSize: 12, color: FILL.cyan, bold: true } }],
    { shape: SH.ROUNDED_RECTANGLE, rectRadius: 0.14, x: 10.25, y: 5.45, w: 2.48, h: 1.0, fill: { color: FILL.navy }, color: FILL.white, align: "center", valign: "middle", margin: 0, isTextBox: true, lang: "ja-JP" });
};

slides[37] = async () => {
  const s = addSlide(37);
  box(s, 0.6, 1.85, 3.1, 4.5, FILL.pale, { radius: 0.2 });
  await icon(s, "TbPresentation", 1.05, 2.55, 2.2, HEX.navy);
  T(s, "1班ずつ発表", { x: 0.6, y: 5.05, w: 3.1, h: 0.5, fontSize: 18, bold: true, color: FILL.navy, align: "center" });
  const items = [["企画名", ""], ["対象", "（誰に）"], ["金属業界の魅力", "（何を）"], ["具体的な体験", "（どう）"], ["参加後の変化", "（どんな興味が生まれるか）"]];
  items.forEach(([a, b], i) => {
    const y = 1.85 + i * 0.92;
    badge(s, String(i + 1), 4.05, y, 0.78, FILL.teal, FILL.white, 26);
    box(s, 4.98, y, 4.85, 0.78, FILL.white, { lineColor: FILL.line, radius: 0.12 });
    T(s, [{ text: a, options: { bold: true, fontSize: 22, color: FILL.navy, breakLine: i === 4 } }, { text: b, options: { fontSize: i === 4 ? 14 : 16, color: FILL.gray } }], { x: 5.2, y, w: 4.55, h: 0.78, valign: "middle" });
  });
  box(s, 10.15, 1.85, 2.58, 4.5, FILL.navy, { shadow: true, radius: 0.2 });
  await icon(s, "TbClock", 10.99, 2.25, 0.9, HEX.cyan);
  T(s, "1:00", { x: 10.15, y: 3.3, w: 2.58, h: 1.1, fontSize: 54, bold: true, color: FILL.white, align: "center", valign: "middle" });
  T(s, "1班\nおよそ1分で発表", { x: 10.3, y: 4.6, w: 2.28, h: 1.0, fontSize: 18, bold: true, color: FILL.cyan, align: "center" });
};

slides[38] = async () => {
  const s = addSlide(38);
  T(s, "各班の企画を比べてみよう。", { x: 0.6, y: 1.8, w: 5.2, h: 0.5, fontSize: 22, bold: true, color: FILL.navy, valign: "middle" });
  T(s, [{ text: "見るポイント：", options: { color: FILL.gray } }, { text: "対象／魅力／体験／参加後の変化", options: { bold: true, color: FILL.teal } }], { x: 5.9, y: 1.8, w: 6.83, h: 0.5, fontSize: 18, align: "right", valign: "middle" });
  const hd = (t) => ({ text: t, options: { bold: true, color: HEX.white, fill: { color: HEX.navy }, align: "center" } });
  const rows = [[hd("班"), hd("対象"), hd("魅力"), hd("体験"), hd("参加後の変化")]];
  "ABCDE".split("").forEach((b, i) => rows.push([{ text: b + "班", options: { bold: true, color: HEX.navy, align: "center", fill: { color: HEX.pale } } }, "", "", "", ""].map((c) => (typeof c === "string" ? { text: c, options: { fill: { color: HEX.white } } } : c))));
  s.addTable(rows, { x: 0.6, y: 2.4, w: 12.13, colW: [1.0, 2.6, 2.75, 3.15, 2.63], rowH: [0.45, 0.55, 0.55, 0.55, 0.55, 0.55], fontSize: 18, fontFace: FONT, color: HEX.text, valign: "middle", border: { type: "solid", pt: 1, color: HEX.line }, margin: [0.04, 0.12, 0.04, 0.12] });
  box(s, 0.6, 5.78, 12.13, 0.62, FILL.orange, { radius: 0.12 });
  await icon(s, "TbHandClick", 0.85, 5.85, 0.48, HEX.navy);
  T(s, "「自分だったら参加したい？」と「なぜ？」を大事にする。", { x: 1.5, y: 5.78, w: 11.0, h: 0.62, fontSize: 22, bold: true, color: FILL.navy, valign: "middle" });
};

slides[39] = async () => {
  const s = addSlide(39);
  const q = [["社会に何を\n提供する？", "橋・車・電力の素材"], ["どんな企業が\nいる？", "製鉄・製錬・加工・商社"], ["誰が顧客？", "自動車・建設・電機"], ["どう儲ける？", "価格×数量−原料費"], ["強みは？", "品質・技術・供給力"], ["どんなニュースが\n影響？", "価格・脱炭素・海外需要"]];
  q.forEach(([a, b], i) => {
    const x = i < 4 ? 0.6 + i * 3.07 : 0.6 + (i - 4) * 3.07, y = i < 4 ? 1.85 : 3.98;
    box(s, x, y, 2.87, 1.95, FILL.pale, { radius: 0.16 });
    badge(s, String(i + 1), x + 0.18, y + 0.18, 0.55, FILL.teal, FILL.white, 20);
    T(s, a, { x: x + 0.18, y: y + 0.8, w: 2.6, h: 0.68, fontSize: 19, bold: true, color: FILL.navy, valign: "middle" });
    T(s, b, { x: x + 0.18, y: y + 1.5, w: 2.6, h: 0.35, fontSize: 15, bold: true, color: FILL.teal, valign: "middle" });
  });
  box(s, 6.74, 3.98, 5.99, 1.95, FILL.orange, { shadow: true, radius: 0.18 });
  badge(s, "7", 6.95, 4.17, 0.62, FILL.navy, FILL.white, 22);
  T(s, "最後に：自分はどこに興味を持った？", { x: 7.7, y: 4.12, w: 4.95, h: 0.75, fontSize: 20, bold: true, color: FILL.navy, valign: "middle" });
  T(s, "社会基盤？ 技術？ 世界？ 未来の課題？ ― あなたの答えが、次の業界研究の出発点。", { x: 6.95, y: 4.95, w: 5.6, h: 0.85, fontSize: 16, color: FILL.navy });
  T(s, [{ text: "今日と同じ方法で、別の業界も調べられる。" }, { text: "　（青字は金属業界での例）", options: { fontSize: 14, color: FILL.gray, bold: false } }], { x: 0.6, y: 6.1, w: 12.13, h: 0.45, fontSize: 22, bold: true, color: FILL.teal, valign: "middle" });
};

slides[40] = async () => {
  const s = addSlide(40);
  const heads = ["気になる企業を\n5社挙げる", "各社の何に\n惹かれるかを書く", "共通点を探す", "その軸で、知らない\n業界の企業も探す"];
  const xs = [0.6, 3.77, 6.94, 10.11];
  heads.forEach((h, i) => {
    badge(s, String(i + 1), xs[i], 1.85, 0.6, i === 3 ? FILL.orange : FILL.teal, i === 3 ? FILL.navy : FILL.white, 20);
    T(s, h, { x: xs[i] + 0.66, y: 1.8, w: 2.0, h: 0.75, fontSize: 16, bold: true, color: FILL.navy, valign: "middle" });
    box(s, xs[i], 2.7, 2.62, 3.25, i === 3 ? "FCEBD3" : FILL.pale, { radius: 0.16 });
    if (i < 3) arrowR(s, xs[i] + 2.67, 4.1, 0.45, 0.45, FILL.cyan);
  });
  T(s, "例", { x: 0.8, y: 2.8, w: 1, h: 0.3, fontSize: 13, bold: true, color: FILL.gray });
  ["トヨタ", "ソニー", "三菱商事", "日本製鉄", "任天堂"].forEach((c, i) => pill(s, c, 0.82, 3.12 + i * 0.55, 2.18, 0.45, FILL.white, FILL.navy, 16, { lineColor: FILL.line, extra: { rectRadius: 0.1 } }));
  T(s, "なぜ惹かれる？", { x: 3.97, y: 2.85, w: 2.3, h: 0.4, fontSize: 16, bold: true, color: FILL.teal });
  for (let i = 0; i < 5; i++) {
    await icon(s, "TbPencil", 3.97, 3.35 + i * 0.5, 0.3, HEX.gray);
    line(s, 4.35, 3.62 + i * 0.5, 6.15, 3.62 + i * 0.5, { color: FILL.gray, width: 1, dash: "dash" });
  }
  T(s, "興味の軸", { x: 7.14, y: 2.85, w: 2.3, h: 0.4, fontSize: 16, bold: true, color: FILL.teal });
  ["世界規模", "技術", "社会への影響", "…"].forEach((c, i) => pill(s, c, 7.16, 3.35 + i * 0.6, 2.18, 0.48, i === 3 ? FILL.pale : FILL.teal, i === 3 ? FILL.gray : FILL.white, 16));
  await icon(s, "TbTelescope", 10.82, 3.0, 1.2, HEX.navy);
  T(s, "新しい業界へ", { x: 10.11, y: 4.35, w: 2.62, h: 0.5, fontSize: 20, bold: true, color: FILL.navy, align: "center" });
  T(s, "知らなかった業界の\n会社も選択肢に入る", { x: 10.3, y: 4.9, w: 2.25, h: 0.85, fontSize: 15, color: FILL.navy, align: "center" });
  T(s, "今日はワークを行いません。別の日に、自分の企業研究で使ってみよう。", { x: 0.6, y: 6.1, w: 12.13, h: 0.45, fontSize: 18, bold: true, color: FILL.teal, valign: "middle" });
};

slides[41] = async () => {
  const s = addSlide(41, "HERO_END");
  s.addText("調べるから、\n興味が生まれる。", { placeholder: "title" });
  await pic(s, "city_bridge", 7.6, 0, 5.73, 7.5, { radius: 0, pos: "east" });
  T(s, "まとめ", { x: 0.6, y: 0.9, w: 3, h: 0.45, fontSize: 20, bold: true, color: FILL.orange });
  T(s, "興味があるから調べるだけじゃない。", { x: 0.6, y: 1.45, w: 6.6, h: 0.5, fontSize: 24, color: FILL.white });
  ["業界の全体像を知る", "企業を比較する", "ニュースで未来を考える"].forEach((t, i) => {
    const y = 4.25 + i * 0.6;
    badge(s, String(i + 1), 0.6, y, 0.46, FILL.cyan, FILL.navy, 16);
    T(s, t, { x: 1.2, y, w: 5.5, h: 0.46, fontSize: 20, bold: true, color: FILL.white, valign: "middle" });
    if (i < 2) T(s, "↓", { x: 0.6, y: y + 0.43, w: 0.46, h: 0.2, fontSize: 10, color: FILL.cyan, align: "center", valign: "middle" });
  });
  T(s, "次は、まだ知らない業界へ。", { x: 0.6, y: 6.15, w: 6.8, h: 0.55, fontSize: 26, bold: true, color: FILL.cyan, valign: "middle" });
};

// ---- 生成 ---------------------------------------------------------------
(async () => {
  for (let n = 1; n <= 41; n++) {
    section(n);
    await slides[n]();
  }
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  await patchFonts(OUT);
  console.log("written", OUT);
})();

// テーマの東アジア用フォント（ea / Jpan）も Noto Sans CJK JP にそろえる
async function patchFonts(file) {
  const JSZip = require(require.resolve("jszip", { paths: [require.resolve("pptxgenjs")] }));
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  const part = "ppt/theme/theme1.xml";
  let xml = await zip.file(part).async("string");
  xml = xml.replace(/<a:ea typeface="[^"]*"\/>/g, `<a:ea typeface="${FONT}"/>`).replace(/<a:font script="Jpan" typeface="[^"]*"\/>/g, `<a:font script="Jpan" typeface="${FONT}"/>`);
  zip.file(part, xml);
  fs.writeFileSync(file, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
}
