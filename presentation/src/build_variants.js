// 홍콩 ELS · 소비자 권리 발표 덱 — 12장, 시안 A/B/C 생성 스크립트 (pptxgenjs)
// 실행: PPTX_SKILL_DIR=<pptx skill dir> NODE_PATH=<node_modules> node build_variants.js <outdir>
const path = require("path");
const fs = require("fs");
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
const { applyTheme } = require(path.join(process.env.PPTX_SKILL_DIR, "scripts/apply_theme.js"));

const OUTDIR = process.argv[2] || ".";
const NOTES = JSON.parse(fs.readFileSync(path.join(__dirname, "notes_orig.json"), "utf8"));
const SLOT = { tx1: "dk1", bg1: "lt1", tx2: "dk2", bg2: "lt2" };

const VARIANTS = {
  A: {
    file: "시안A_네이비레드.pptx", radius: 0.08, outline: false, onAccentHex: "FFFFFF",
    theme: { name: "ELS A Navy Red", headFontFace: "Pretendard", bodyFontFace: "Pretendard", colors: {
      dk1: "0E1726", lt1: "FFFFFF", dk2: "1F2A44", lt2: "F2F4F8", accent1: "E5484D", accent2: "0E9384", accent3: "D9942B",
      accent4: "5B667A", accent5: "D5DAE3", accent6: "2B3857", hlink: "0E9384", folHlink: "5B667A" } },
    role: { bg: "bg1", fg: "tx1", muted: "accent4", faint: "accent5", card: "bg2", strong: "tx1", onStrong: "bg1", onStrongMuted: "accent5",
      neutral: "tx2", loss: "accent1", right: "accent2", warn: "accent3", onAccent: "bg1", pillBg: "bg1", eyebrow: "accent1",
      coverBg: "tx1", coverFg: "bg1", coverSub: "accent5", coverPanel: "accent6", coverEyebrow: "accent2",
      closeBg: "tx1", closeFg: "bg1", closeCard: "accent6", closeMuted: "accent4" },
  },
  B: {
    file: "시안B_다크.pptx", radius: 0.12, outline: false, onAccentHex: "0B1220",
    theme: { name: "ELS B Dark", headFontFace: "Pretendard", bodyFontFace: "Pretendard", colors: {
      dk1: "0B1220", lt1: "FFFFFF", dk2: "18233D", lt2: "E8ECF4", accent1: "FF6B6B", accent2: "2DD4BF", accent3: "F5B544",
      accent4: "93A0B8", accent5: "34425F", accent6: "24304F", hlink: "2DD4BF", folHlink: "93A0B8" } },
    role: { bg: "tx1", fg: "bg1", muted: "accent4", faint: "accent5", card: "tx2", strong: "bg2", onStrong: "tx1", onStrongMuted: "accent5",
      neutral: "bg2", loss: "accent1", right: "accent2", warn: "accent3", onAccent: "tx1", pillBg: "tx1", eyebrow: "accent2",
      coverBg: "tx1", coverFg: "bg1", coverSub: "accent4", coverPanel: "accent5", coverEyebrow: "accent2",
      closeBg: "tx1", closeFg: "bg1", closeCard: "tx2", closeMuted: "accent4" },
  },
  C: {
    file: "시안C_미니멀.pptx", radius: 0, outline: true, onAccentHex: "FFFFFF",
    theme: { name: "ELS C Minimal", headFontFace: "Pretendard", bodyFontFace: "Pretendard", colors: {
      dk1: "111111", lt1: "FFFFFF", dk2: "111111", lt2: "F4F4F2", accent1: "1F4BFF", accent2: "111111", accent3: "8A8A8A",
      accent4: "6B6B6B", accent5: "DADADA", accent6: "2A2A2A", hlink: "1F4BFF", folHlink: "6B6B6B" } },
    role: { bg: "bg1", fg: "tx1", muted: "accent4", faint: "accent5", card: "bg1", strong: "tx1", onStrong: "bg1", onStrongMuted: "accent5",
      neutral: "tx2", loss: "accent1", right: "accent2", warn: "accent3", onAccent: "bg1", pillBg: "bg2", eyebrow: "accent1",
      coverBg: "bg1", coverFg: "tx1", coverSub: "accent4", coverPanel: "accent5", coverEyebrow: "accent1",
      closeBg: "tx1", closeFg: "bg1", closeCard: "accent6", closeMuted: "accent3" },
  },
};

const iconCache = {};
async function icon(Comp, hex, size = 256) {
  const key = Comp.name + hex;
  if (iconCache[key]) return iconCache[key];
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Comp, { color: "#" + hex, size }));
  const buf = await sharp(Buffer.from(svg)).resize(size, size, { fit: "contain", background: { r: 0, g: 0, b: 0, alpha: 0 } }).png().toBuffer();
  return (iconCache[key] = "image/png;base64," + buf.toString("base64"));
}

async function build(V) {
  const OUT = path.join(OUTDIR, V.file);
  const THEME = V.theme;
  const c = (r) => V.role[r];
  const hx = (r) => THEME.colors[SLOT[V.role[r]] || V.role[r]];
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9"; // 10 x 5.625
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  pres.title = "홍콩 ELS 사태로 알아보는 소비자 권리";
  const CF = "+mn-lt";

  // ---------- layouts ----------
  const ph = (name, type, o) => ({ placeholder: { options: { name, type, margin: 0, ...o }, text: "" } });
  pres.defineSlideMaster({
    title: "Cover", background: { color: c("coverBg") },
    objects: [
      ph("eyebrow", "body", { x: 0.6, y: 1.45, w: 5.2, h: 0.3, fontSize: 12, bold: true, color: c("coverEyebrow"), charSpacing: 2 }),
      ph("title", "title", { x: 0.6, y: 1.85, w: 5.3, h: 1.35, fontSize: 31, bold: true, color: c("coverFg"), valign: "top", align: "left" }),
      ph("sub", "body", { x: 0.6, y: 3.4, w: 5.2, h: 0.5, fontSize: 14, color: c("coverSub"), valign: "top" }),
    ],
  });
  pres.defineSlideMaster({
    title: "Content", background: { color: c("bg") },
    objects: [
      ph("eyebrow", "body", { x: 0.5, y: 0.32, w: 6, h: 0.22, fontSize: 10, bold: true, color: c("eyebrow"), charSpacing: 2 }),
      ph("title", "title", { x: 0.5, y: 0.55, w: 9, h: 0.55, fontSize: 26, bold: true, color: c("fg"), valign: "middle", align: "left" }),
      ph("source", "body", { x: 0.5, y: 5.12, w: 8.3, h: 0.28, fontSize: 8, color: c("muted"), valign: "middle" }),
    ],
    slideNumber: { x: 9.0, y: 5.12, w: 0.5, h: 0.28, fontSize: 8, color: c("muted"), align: "right" },
  });
  pres.defineSlideMaster({
    title: "Closing", background: { color: c("closeBg") },
    objects: [
      ph("eyebrow", "body", { x: 0.6, y: 0.5, w: 6, h: 0.25, fontSize: 10, bold: true, color: c("coverEyebrow") === "accent1" ? "accent5" : c("coverEyebrow"), charSpacing: 2 }),
      ph("title", "title", { x: 0.6, y: 0.85, w: 8.8, h: 1.2, fontSize: 28, bold: true, color: c("closeFg"), valign: "top", align: "left" }),
      ph("source", "body", { x: 0.6, y: 4.75, w: 8.8, h: 0.55, fontSize: 8, color: c("closeMuted"), valign: "top" }),
    ],
  });

  // ---------- helpers ----------
  let n = 0;
  const nm = (p) => `${p} ${++n}`;
  const T = (s, text, o) => s.addText(text, { isTextBox: true, margin: 0, valign: "top", color: c("fg"), objectName: nm("Text"), ...o });
  const box = (s, x, y, w, h, fill, lineOpt) => {
    const base = { x, y, w, h, fill: { color: fill }, line: lineOpt || { type: "none" }, objectName: nm("Card") };
    if (V.radius > 0) s.addShape(pres.ShapeType.roundRect, { ...base, rectRadius: V.radius });
    else s.addShape(pres.ShapeType.rect, base);
  };
  const card = (s, x, y, w, h, role = "card") =>
    box(s, x, y, w, h, c(role), role === "card" && V.outline ? { color: c("faint"), width: 0.75 } : null);
  const dot = (s, x, y, d, fill) =>
    s.addShape(pres.ShapeType.ellipse, { x, y, w: d, h: d, fill: { color: fill }, line: { type: "none" }, objectName: nm("Dot") });
  const pill = (s, text, x, y, w, fill, color, o = {}) => {
    const h = o.h || 0.22;
    s.addText(text, { shape: pres.ShapeType.roundRect, rectRadius: V.radius > 0 ? h / 2 : 0.03, x, y, w, h, fill: { color: fill }, line: { type: "none" },
      fontSize: 8.5, bold: true, color, align: "center", valign: "middle", margin: 0, objectName: nm("Pill"), ...o });
  };
  const iconDot = async (s, Comp, x, y, d, fill) => {
    dot(s, x, y, d, fill);
    const k = d * 0.48;
    s.addImage({ data: await icon(Comp, V.onAccentHex), x: x + (d - k) / 2, y: y + (d - k) / 2, w: k, h: k, objectName: nm("Icon") });
  };
  const line = (s, x, y, w, color, pt = 1.5) =>
    s.addShape(pres.ShapeType.line, { x, y, w, h: 0, line: { color, width: pt }, objectName: nm("Line") });
  const vline = (s, x, y, h, color, pt = 1.5, dash) =>
    s.addShape(pres.ShapeType.line, { x, y, w: 0, h, line: { color, width: pt, ...(dash ? { dashType: "dash" } : {}) }, objectName: nm("Line") });
  const chartBase = () => ({
    catAxisLabelColor: hx("fg"), catAxisLabelFontFace: CF, catAxisLabelFontSize: 11,
    valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" },
    catAxisLineShow: false, valAxisLineShow: false,
    showValue: true, dataLabelColor: hx("fg"), dataLabelFontFace: CF, dataLabelFontSize: 11, dataLabelFontBold: true,
    showLegend: false, barDir: "bar", catAxisOrientation: "maxMin", barGapWidthPct: 45, dataLabelPosition: "outEnd",
  });
  const head = (s, eyebrow, title, source) => {
    s.addText(eyebrow, { placeholder: "eyebrow" });
    s.addText(title, { placeholder: "title" });
    if (source) s.addText(source, { placeholder: "source" });
  };
  const SEC = ["도입", "배경과 권리", "실제 결과", "원인과 대안", "결론"];
  const content = (sec) => pres.addSlide({ masterName: "Content", sectionTitle: sec });

  // ================= 1. 표지 =================
  pres.addSection({ title: SEC[0] });
  {
    const s = pres.addSlide({ masterName: "Cover", sectionTitle: SEC[0] });
    s.addText("주제 5 · 소비자의 권리", { placeholder: "eyebrow" });
    s.addText("홍콩 ELS 사태로 알아보는\n소비자 권리", { placeholder: "title" });
    s.addText("금융소비자보호법이 준 권리는 실제로 작동했는가", { placeholder: "sub" });
    s.addChart(pres.charts.DOUGHNUT, [{ name: "손실확정 계좌 원금", labels: ["손실", "잔여"], values: [44.2, 55.8] }], {
      x: 6.05, y: 0.95, w: 3.4, h: 3.4, holeSize: 72, chartColors: [hx("loss"), hx("coverPanel")],
      showLegend: false, showValue: false, showPercent: false, showLabel: false, showTitle: false,
      dataBorder: { pt: 0, color: hx("coverBg") }, objectName: "손실률 도넛",
    });
    T(s, "44.2%", { x: 6.75, y: 2.1, w: 2.0, h: 0.6, fontSize: 34, bold: true, color: c("coverFg"), align: "center", valign: "middle" });
    T(s, "누적 손실률", { x: 6.75, y: 2.72, w: 2.0, h: 0.25, fontSize: 11, color: c("coverSub"), align: "center" });
    T(s, "원금 10.4조원 중 4.6조원 손실 · 5대 판매은행, 2024년 말", { x: 5.75, y: 4.5, w: 4.0, h: 0.25, fontSize: 9, color: c("muted"), align: "center" });
    s.addNotes(NOTES[0]);
  }

  // ================= 2. 배경 =================
  pres.addSection({ title: SEC[1] });
  {
    const s = content(SEC[1]);
    head(s, "배경", "법이 생긴 뒤에도 더 큰 손실이 발생함",
      "강의자료 8장 1절 · 금융위원회·금융감독원(2025.2) · 원 넓이는 손실액 비례(팀 작성)");
    const ly = 3.0;
    line(s, 0.9, ly, 8.2, c("faint"), 2);
    const cols = [1.9, 4.9, 7.9];
    const yrs = ["2019", "2021.3", "2024"];
    const evs = ["DLF 사태", "금융소비자보호법 시행", "홍콩 H지수 ELS 사태"];
    cols.forEach((cx, i) => {
      T(s, yrs[i], { x: cx - 1.3, y: 1.3, w: 2.6, h: 0.36, fontSize: 22, bold: true, align: "center", color: i === 1 ? c("right") : c("fg") });
      T(s, evs[i], { x: cx - 1.3, y: 1.68, w: 2.6, h: 0.25, fontSize: 12, align: "center", color: c("muted") });
    });
    const dE = 1.9, dD = dE / Math.sqrt(46000 / 4468); // 넓이 비 10.3배
    dot(s, cols[0] - dD / 2, ly - dD / 2, dD, c("loss"));
    T(s, "손실 4,468억원", { x: cols[0] - 1.3, y: 3.5, w: 2.6, h: 0.3, fontSize: 15, bold: true, align: "center" });
    await iconDot(s, fa.FaShieldAlt, cols[1] - 0.38, ly - 0.38, 0.76, c("right"));
    T(s, "6대 판매원칙 도입", { x: cols[1] - 1.3, y: 3.5, w: 2.6, h: 0.3, fontSize: 15, bold: true, align: "center" });
    dot(s, cols[2] - dE / 2, ly - dE / 2, dE, c("loss"));
    T(s, "4.6조원", { x: cols[2] - 0.9, y: ly - 0.45, w: 1.8, h: 0.5, fontSize: 30, bold: true, color: c("onAccent"), align: "center", valign: "middle" });
    T(s, "손실률 44.2%", { x: cols[2] - 0.9, y: ly + 0.1, w: 1.8, h: 0.25, fontSize: 11, color: c("onAccent"), align: "center" });
    card(s, 0.5, 4.3, 9, 0.58);
    pill(s, "손실 약 10배", 0.72, 4.44, 1.15, c("loss"), c("onAccent"), { h: 0.3, fontSize: 10.5 });
    T(s, "규제는 강화됐지만 현장의 소비자 보호는 작동하지 않음", { x: 2.05, y: 4.3, w: 7.2, h: 0.58, fontSize: 13, valign: "middle" });
    s.addNotes(NOTES[1] + "\n\n[시각화] 원의 넓이는 손실액에 비례함. 4.6조원 ÷ 4,468억원 ≈ 10.3배(팀 계산). DLF는 전체 투자자, ELS는 5대 판매은행 손실확정 계좌 기준이라 집계 범위가 다르다는 점을 질문이 나오면 밝힘."
      + "\n\n[슬라이드에서 뺀 수치] DLF: 투자자 3,243명, 판매액 7,950억원. 금소법: 청약철회권·위법계약해지권, 설명의무 입증책임 전환. ELS: 은행 손실확정 17.0만 계좌. 하단 문장은 금감원 검사 결과의 평가임.");
  }

  // ================= 3. ELS =================
  {
    const s = content(SEC[1]);
    head(s, "ELS란", "은행 창구에서 팔린 '예금 같은' 투자상품",
      "강의자료 8장 1·2절 · 금융위원회·금융감독원(2025.2), 5대 판매은행 2024년 말 기준");
    const xa = 1.75, wa = 1.85, xb = 3.75, wb = 1.95, rh = 0.7;
    card(s, xa, 1.3, wa, 3.57);
    card(s, xb, 1.3, wb, 3.57, "strong");
    T(s, "정기예금", { x: xa, y: 1.42, w: wa, h: 0.4, fontSize: 15, bold: true, align: "center", valign: "middle" });
    T(s, "ELS", { x: xb, y: 1.42, w: wb, h: 0.4, fontSize: 15, bold: true, align: "center", valign: "middle", color: c("onStrong") });
    const rows = [["판매 창구", "은행", "은행", false], ["원금", "보장", "손실 가능", true], ["예금자보호", "대상", "대상 아님", true], ["수익", "확정 이자", "지수 연계", true]];
    const ok = await icon(fa.FaCheckCircle, hx("right")), no = await icon(fa.FaTimesCircle, hx("loss"));
    const eqA = await icon(fa.FaEquals, hx("muted")), eqB = await icon(fa.FaEquals, hx("onStrongMuted"));
    rows.forEach((r, i) => {
      const y = 1.95 + i * rh;
      if (i > 0) line(s, 0.5, y, 1.15, c("faint"), 0.75);
      T(s, r[0], { x: 0.5, y, w: 1.15, h: rh, fontSize: 12, color: c("muted"), valign: "middle" });
      s.addImage({ data: r[3] ? ok : eqA, x: xa + 0.22, y: y + rh / 2 - 0.1, w: 0.2, h: 0.2, objectName: nm("Icon") });
      T(s, r[1], { x: xa + 0.52, y, w: wa - 0.6, h: rh, fontSize: 14, bold: true, valign: "middle" });
      s.addImage({ data: r[3] ? no : eqB, x: xb + 0.22, y: y + rh / 2 - 0.1, w: 0.2, h: 0.2, objectName: nm("Icon") });
      T(s, r[2], { x: xb + 0.52, y, w: wb - 0.6, h: rh, fontSize: 14, bold: true, valign: "middle", color: c("onStrong") });
    });
    const rx = 6.1, rw = 3.4, bw = rw - 0.6;
    card(s, rx, 1.3, rw, 3.57);
    T(s, "손실확정 계좌", { x: rx + 0.3, y: 1.55, w: bw, h: 0.22, fontSize: 10.5, color: c("muted") });
    T(s, "17.0만", { x: rx + 0.3, y: 1.8, w: bw, h: 0.6, fontSize: 34, bold: true, valign: "middle" });
    T(s, "손실액", { x: rx + 0.3, y: 2.65, w: bw, h: 0.22, fontSize: 10.5, color: c("muted") });
    T(s, "4.6조원", { x: rx + 0.3, y: 2.9, w: bw, h: 0.6, fontSize: 34, bold: true, color: c("loss"), valign: "middle" });
    s.addShape(pres.ShapeType.rect, { x: rx + 0.3, y: 3.85, w: bw, h: 0.32, fill: { color: c("faint") }, line: { type: "none" }, objectName: "원금 막대" });
    s.addShape(pres.ShapeType.rect, { x: rx + 0.3, y: 3.85, w: bw * 0.442, h: 0.32, fill: { color: c("loss") }, line: { type: "none" }, objectName: "손실 막대" });
    T(s, "44.2%", { x: rx + 0.3, y: 3.85, w: bw * 0.442, h: 0.32, fontSize: 10.5, bold: true, color: c("onAccent"), align: "center", valign: "middle" });
    T(s, "원금 10.4조원", { x: rx + 0.3, y: 4.25, w: bw, h: 0.22, fontSize: 10.5, color: c("muted"), align: "right" });
    s.addNotes(NOTES[2] + "\n\n[수정 반영] 손실 조건은 \"대표적인 구조에서는 만기평가일의 기초지수와 정해진 상환조건 등에 따라 원금손실이 발생할 수 있다\"로 설명함. 장중 하락만으로 손실이 확정된다고 말하지 않음."
      + "\n\n[슬라이드에서 뺀 설명] 은행은 ELS를 신탁(ELT)·펀드(ELF) 형태로 판매함. 2021년 초 H지수 고점 판매분의 3년 만기가 2024년에 도래하며 손실이 확정됨.");
  }

  // ================= 4. 권리 =================
  {
    const s = content(SEC[1]);
    head(s, "금융소비자보호법", "금소법이 소비자에게 준 여섯 가지 장치",
      "강의자료 8장 1절 주요내용 2~4 · 금소법 제28·42·44·46·47·57조");
    const G = { a: ["권리 강화", "right"], b: ["분쟁 시 권익 확보", "neutral"], c: ["처벌 강화", "loss"] };
    const items = [
      [fa.FaUndoAlt, "청약철회권", "대상 상품은 7일 안에 철회", "a"],
      [fa.FaFileContract, "위법계약해지권", "위반 계약을 위약금 없이 해지", "a"],
      [fa.FaHandshake, "조정이탈금지", "소액분쟁 조정 중 제소 제한", "b"],
      [fa.FaFolderOpen, "자료열람요구권", "분쟁 시 자료 열람·사본 요구", "b"],
      [fa.FaBalanceScale, "입증책임 전환", "금융회사가 무과실을 입증", "b"],
      [fa.FaGavel, "징벌적 과징금", "관련 수입의 50% 이내", "c"],
    ];
    const cw = 2.85, gap = 0.225, chh = 1.68;
    for (let i = 0; i < items.length; i++) {
      const [ic, name, desc, g] = items[i];
      const x = 0.5 + (i % 3) * (cw + gap), y = 1.3 + Math.floor(i / 3) * (chh + 0.2);
      card(s, x, y, cw, chh);
      await iconDot(s, ic, x + 0.22, y + 0.2, 0.54, c(G[g][1]));
      T(s, G[g][0], { x: x + 0.95, y: y + 0.2, w: cw - 1.17, h: 0.54, fontSize: 9, bold: true, color: c(G[g][1]), align: "right", valign: "middle" });
      T(s, name, { x: x + 0.22, y: y + 0.9, w: cw - 0.44, h: 0.32, fontSize: 16, bold: true, valign: "middle" });
      T(s, desc, { x: x + 0.22, y: y + 1.25, w: cw - 0.44, h: 0.26, fontSize: 11.5, color: c("muted"), valign: "middle" });
    }
    s.addNotes(NOTES[3] + "\n\n[수정 반영] ① 청약철회는 \"대상 투자성 상품·자문은 원칙적으로 7일\"(법 제46조, 시행령 제37조)로 한정함. 철회기간 중 자금 운용에 동의한 경우 등 예외가 있어 모든 ELS 거래에 자동 적용되지 않음. ② '소송이탈금지'는 '조정이탈금지'로 표기함(법 제42조, 2,000만원 이하 소액분쟁, 조정안 제시 전이 원칙이며 예외 존재). ③ 위법계약해지권은 판매원칙 위반 계약을 수수료·위약금 없이 해지, 위반을 안 날부터 1년·계약일부터 5년 이내. ④ 입증책임 전환은 설명의무 위반 손해배상에서 금융회사가 고의·과실 없음을 입증하는 것에 한정됨(법 제44조 제2항). ⑤ 과징금은 판매원칙 위반 시 관련 수입 등의 50% 이내이며 적합성·적정성 위반은 제외됨.");
  }

  // ================= 5. 실제 구제: 31.4% =================
  pres.addSection({ title: SEC[2] });
  {
    const s = content(SEC[2]);
    head(s, "실제 구제", "실제 구제는 자율배상 중심으로 진행됨",
      "금융위원회·금융감독원(2025.2), 5대 판매은행 2024년 말 기준");
    card(s, 0.5, 1.3, 3.6, 3.57, "strong");
    T(s, "평균 배상비율", { x: 0.8, y: 1.6, w: 3.0, h: 0.25, fontSize: 12, color: c("onStrongMuted") });
    T(s, "31.4%", { x: 0.8, y: 2.1, w: 3.0, h: 1.1, fontSize: 60, bold: true, color: c("onStrong"), valign: "middle" });
    T(s, "배상 동의 15.9만 계좌 기준", { x: 0.8, y: 4.3, w: 3.0, h: 0.25, fontSize: 10.5, color: c("onStrongMuted") });
    const bx = 4.4, bwid = 5.1, part = 1.3 / 4.1;
    T(s, "손실 약 4.1조원의 구성", { x: bx, y: 1.4, w: bwid, h: 0.3, fontSize: 13, bold: true, valign: "middle" });
    s.addShape(pres.ShapeType.rect, { x: bx, y: 1.85, w: bwid, h: 0.85, fill: { color: c("faint") }, line: { type: "none" }, objectName: "소비자 부담 막대" });
    s.addShape(pres.ShapeType.rect, { x: bx, y: 1.85, w: bwid * part, h: 0.85, fill: { color: c("loss") }, line: { type: "none" }, objectName: "배상액 막대" });
    T(s, "배상 1.3조", { x: bx, y: 1.85, w: bwid * part, h: 0.85, fontSize: 14, bold: true, color: c("onAccent"), align: "center", valign: "middle" });
    T(s, "소비자 부담 2.8조", { x: bx + bwid * part, y: 1.85, w: bwid * (1 - part), h: 0.85, fontSize: 14, bold: true, align: "center", valign: "middle" });
    T(s, "구제 경로", { x: bx, y: 3.3, w: bwid, h: 0.25, fontSize: 11, color: c("muted"), valign: "middle" });
    const chips = [["금감원\n분쟁조정기준", "neutral", "onAccent"], ["은행\n자율배상", "loss", "onAccent"], ["15.9만 계좌\n배상 동의", "right", "onAccent"]];
    chips.forEach((ch, i) => {
      const x = bx + i * 1.8;
      pill(s, ch[0], x, 3.65, 1.5, c(ch[1]), c(ch[2]), { h: 0.72, fontSize: 11, rectRadius: V.radius > 0 ? 0.12 : 0.03 });
      if (i < 2) T(s, "→", { x: x + 1.5, y: 3.65, w: 0.3, h: 0.72, fontSize: 16, bold: true, align: "center", valign: "middle", color: c("muted") });
    });
    s.addNotes(NOTES[4].replace("피해 구제는 강의에서 배운 권리 행사가 아니라 금감원 분쟁조정기준과 은행 자율배상으로 이루어짐", "피해 구제는 금감원 분쟁조정기준에 따른 은행 자율배상 중심으로 진행됨")
      + "\n\n[수정 반영] 제목을 '자율배상 중심으로 진행됨'으로 바꿈. 분쟁조정·손해배상도 권리구제 수단이므로 권리 행사 전체를 부정하지 않음. 31.4%는 배상 동의 15.9만 계좌(배상진행 16.9만건의 93.8%) 기준이며, 전체 피해자에게 똑같이 적용된 확정 부담률이 아님. 분쟁조정기준 마련일은 2024.3.11."
      + "\n\n[구성] 위법계약해지권의 한계와 대표사례는 다음 장(6번)으로 옮김.");
  }

  // ================= 6. 대표사례 =================
  {
    const s = content(SEC[2]);
    head(s, "대표사례", "배상비율은 사례에 따라 30~65%",
      "금감원 분쟁조정위원회 대표사례(2024.5.13 의결) · 금융위원회 위법계약해지권 안내 · 법제처");
    s.addChart(pres.charts.BAR, [{ name: "배상비율(%)", labels: ["농협", "국민", "신한", "SC제일", "하나"], values: [65, 60, 55, 55, 30] }], {
      ...chartBase(), x: 0.4, y: 1.25, w: 5.2, h: 3.3, chartColors: [hx("neutral")], valAxisMaxVal: 100, valAxisMinVal: 0, dataLabelFormatCode: '0"%"',
      objectName: "대표사례 배상비율 차트",
    });
    T(s, "판매사 책임에 투자자 개별 사정을 반영", { x: 0.5, y: 4.6, w: 5.1, h: 0.25, fontSize: 10.5, color: c("muted"), valign: "middle" });
    card(s, 5.9, 1.3, 3.6, 1.68);
    await iconDot(s, fa.FaUniversity, 6.12, 1.5, 0.54, c("neutral"));
    pill(s, "배상 60%", 8.3, 1.63, 0.98, c("pillBg"), c("fg"), { h: 0.28, fontSize: 10 });
    T(s, "예금하러 온 고객에게 ELT 권유", { x: 6.12, y: 2.2, w: 3.2, h: 0.3, fontSize: 13.5, bold: true, valign: "middle" });
    T(s, "암 보험 진단금을 맡기려던 사례", { x: 6.12, y: 2.52, w: 3.2, h: 0.26, fontSize: 11, color: c("muted"), valign: "middle" });
    card(s, 5.9, 3.19, 3.6, 1.68, "strong");
    await iconDot(s, fa.FaExclamation, 6.12, 3.39, 0.54, c("warn"));
    T(s, "위법계약해지권의 한계", { x: 6.12, y: 4.09, w: 3.2, h: 0.3, fontSize: 13.5, bold: true, valign: "middle", color: c("onStrong") });
    T(s, "해지는 장래효, 확정된 손실은 못 되돌림", { x: 6.12, y: 4.41, w: 3.2, h: 0.26, fontSize: 11, color: c("onStrongMuted"), valign: "middle" });
    s.addNotes("[담당: 집필 A] (5번에서 분리한 장)"
      + "\n\n[핵심] 실제 배상비율은 사례마다 다름. 2024.5.13 분조위 의결(보도자료 5.14) 대표사례: 농협 65% / 국민 60% / 신한 55% / SC제일 55% / 하나 30%."
      + "\n\n[사례] 국민은행: 암 보험 진단금을 정기예금에 넣으려던 고객에게 ELT를 권유(60%). 농협은행: 70대 고객의 투자성향 파악 부실, 고령자 보호절차 미준수(65%). 신한은행: 70대 고객의 답변 유도, 위험 왜곡 설명(55%). SC제일은행: ELS 최초투자 고객, 사후 확인 통화 처리 문제(55%). 하나은행: 고객정보 파악 부실, 손실위험 설명 누락(30%)."
      + "\n\n[주의] 이 5건은 2021.3.24 이전 판매분이라 '금소법 시행 후 판매 사례'라고 말하면 안 됨. 배상비율은 판매사 책임에 투자자의 경험·가입 목적 등 개별 사정을 반영한 것이므로, 낮은 배상비율만으로 보호 실패가 입증되는 것은 아님(시장손실과 판매절차 위반 책임을 구분)."
      + "\n\n[위법계약해지권] 해지 효과는 장래를 향해 발생함. 만기에 이미 확정된 손실을 돌려받는 장치가 아니며 손해배상청구와 구별됨. 적용 요건: 계속적 거래이면서 해지 시 재산상 불이익이 있는 상품(법 제47조, 시행령 제38조). 'ELS에는 적용되지 않았다'고 단정하지 않고 구조적 한계까지만 말함.");
  }

  // ================= 7. 제재 경과 =================
  {
    const s = content(SEC[2]);
    head(s, "제재 경과", "첫 검사에서 최종 의결까지 약 2년 10개월",
      "금융위원회·금융감독원(2025.2) · 금융위원회 보도자료(2026.11.4, 제19차 정례회의 의결) · 기간은 팀 계산");
    const xs = [1.2, 3.05, 4.9, 6.75, 8.6], ty = 3.2;
    // 기간 표시
    line(s, xs[0], 2.0, xs[4] - xs[0], c("fg"), 1.25);
    vline(s, xs[0], 1.9, 0.2, c("fg"), 1.25);
    vline(s, xs[4], 1.9, 0.2, c("fg"), 1.25);
    pill(s, "약 2년 10개월", 3.7, 1.72, 2.4, c("loss"), c("onAccent"), { h: 0.56, fontSize: 18 });
    line(s, xs[0], ty, xs[4] - xs[0], c("faint"), 2.5);
    const tl = [["2024.1~3", "금감원 현장검사", "검사", "neutral"], ["2024.3.11", "분쟁조정기준", "구제", "right"], ["2024.5.13", "대표사례 의결", "구제", "right"],
      ["2025.2", "재발방지 대책", "대책", "warn"], ["2026.11.4", "과징금 최종 의결", "제재", "loss"]];
    tl.forEach((t, i) => {
      const last = i === 4, d = last ? 0.4 : 0.26;
      dot(s, xs[i] - d / 2, ty - d / 2, d, c(t[3]));
      pill(s, t[2], xs[i] - 0.35, 2.62, 0.7, c(t[3]), c("onAccent"), { h: 0.26, fontSize: 9.5 });
      T(s, t[0], { x: xs[i] - 0.9, y: 3.6, w: 1.8, h: 0.32, fontSize: 15, bold: true, align: "center", color: last ? c("loss") : c("fg") });
      T(s, t[1], { x: xs[i] - 0.9, y: 3.95, w: 1.8, h: 0.25, fontSize: 11, align: "center", color: c("muted") });
    });
    s.addNotes("[담당: 집필 B · 파트너 검토] (제재 장에서 분리한 연표)"
      + "\n\n[핵심] 금감원 현장검사(2024.1~3) → 분쟁조정기준 마련(2024.3.11) → 대표사례 배상비율 의결(2024.5.13) → 재발방지 대책 발표(2025.2) → 과징금·과태료 최종 의결(2026.11.4, 금융위 제19차 정례회의)."
      + "\n\n[시사점] 첫 현장검사(2024.1)부터 최종 의결(2026.11.4)까지 약 2년 10개월(팀 계산). 구제(자율배상)는 2024년에 시작됐지만 제재 확정은 2026년 말. 결론의 '신속한 제재' 제안으로 이어짐."
      + "\n\n[예상 질문] Q. 왜 이렇게 오래 걸렸나?\nA. 금융위는 다수의 금융소비자와 관련되고 주요 금융기관이 대상이 된 금소법 위반 첫 사례여서 '철저한 법리검토와 충분한 사실관계 확인'을 거쳐 여러 차례 심의했다고 밝힘.");
  }

  // ================= 8. 제재 규모 =================
  {
    const s = content(SEC[2]);
    head(s, "제재 규모", "최종 제재는 약 3,746억원",
      "강의자료 8장 1절 주요내용 4 · 금융위원회 보도자료(2026.11.4, 제19차 정례회의 의결)");
    card(s, 0.5, 1.3, 3.9, 3.57, "strong");
    T(s, "은행 7곳 · 증권사 6곳", { x: 0.8, y: 1.6, w: 3.3, h: 0.25, fontSize: 11.5, color: c("onStrongMuted") });
    T(s, "3,746억원", { x: 0.8, y: 2.05, w: 3.3, h: 0.9, fontSize: 44, bold: true, color: c("onStrong"), valign: "middle" });
    T(s, "과징금 3,712억 + 과태료 34억", { x: 0.8, y: 3.05, w: 3.3, h: 0.28, fontSize: 12, color: c("onStrong"), valign: "middle" });
    pill(s, "금융위 최대 규모 금전 제재", 0.8, 4.15, 2.2, c("loss"), c("onAccent"), { h: 0.32, fontSize: 10 });
    T(s, "은행별 과징금 (억원)", { x: 4.75, y: 1.3, w: 4.7, h: 0.3, fontSize: 13, bold: true, valign: "middle" });
    s.addChart(pres.charts.BAR, [{ name: "과징금(억원)", labels: ["국민", "하나", "농협", "신한", "SC제일", "우리"], values: [1723.34, 924.75, 401.75, 387.45, 179.0, 3.31] }], {
      ...chartBase(), x: 4.65, y: 1.6, w: 4.85, h: 3.3, chartColors: [hx("loss")], valAxisMaxVal: 2100, valAxisMinVal: 0, dataLabelFormatCode: "#,##0",
      objectName: "은행별 과징금 차트",
    });
    s.addNotes(NOTES[5] + "\n\n[시각화] 합계는 3,746.32억원(과징금 3,712.42억 + 과태료 33.90억, 슬라이드는 억 단위 반올림). '첫 사례'는 금융위 표현대로 \"다수 금융소비자와 관련되고 주요 금융기관이 대상이 된 금소법 위반 첫 사례\"라고 말함. 수협은행은 과징금 없이 과태료만 부과되어 차트에 없음. 연표는 앞 장(7번)으로 옮김.");
  }

  // ================= 9. 해외 =================
  pres.addSection({ title: SEC[3] });
  {
    const s = content(SEC[3]);
    head(s, "해외 사례", "방향은 같지만, 규율하는 지점은 다름",
      "FCA FG22/5 · SEC Regulation Best Interest · 일본증권업협회 자율규제 개정(2023) · 단계 구분은 팀 작성");
    const gx = 3.75, gw = 5.75, st = ["상품 설계", "설명·표시", "권유", "사후 지원"], sw = gw / 4;
    st.forEach((t, i) => T(s, t, { x: gx + i * sw, y: 1.3, w: sw, h: 0.3, fontSize: 11, bold: true, align: "center", valign: "middle", color: c("muted") }));
    const rowsD = [
      ["UK", "FCA Consumer Duty", "영국 · 2023", 0, 4, "고객에게 좋은 결과", "right"],
      ["US", "Regulation Best Interest", "미국 · 2020", 2, 1, "최선이익", "neutral"],
      ["JP", "복잡한 구조채 표시", "일본 · 2023", 1, 1, "경고 표시", "warn"],
    ];
    rowsD.forEach((r, i) => {
      const y = 1.7 + i * 0.9;
      card(s, 0.5, y, 9, 0.78);
      s.addText(r[0], { shape: pres.ShapeType.ellipse, x: 0.68, y: y + 0.15, w: 0.48, h: 0.48, fill: { color: c(r[6]) }, line: { type: "none" },
        fontSize: 10, bold: true, color: c("onAccent"), align: "center", valign: "middle", margin: 0, objectName: nm("Badge") });
      T(s, r[1], { x: 1.3, y: y + 0.14, w: 2.4, h: 0.26, fontSize: 12.5, bold: true, valign: "middle" });
      T(s, r[2], { x: 1.3, y: y + 0.42, w: 2.4, h: 0.22, fontSize: 10, color: c("muted") });
      for (let k = 1; k < 4; k++) vline(s, gx + k * sw, y + 0.12, 0.54, c("faint"), 0.75, true);
      pill(s, r[5], gx + r[3] * sw + 0.08, y + 0.21, r[4] * sw - 0.16, c(r[6]), c("onAccent"), { h: 0.36, fontSize: 11 });
    });
    card(s, 0.5, 4.4, 9, 0.5, "strong");
    pill(s, "한국", 0.7, 4.52, 0.62, c("loss"), c("onAccent"), { h: 0.26, fontSize: 10 });
    T(s, "최선이익 원칙을 담은 금소법 개정안, 국회 계류 중", { x: 1.5, y: 4.4, w: 7.8, h: 0.5, fontSize: 12.5, valign: "middle", color: c("onStrong") });
    s.addNotes(NOTES[6] + "\n\n[수정 반영] 세 제도는 소비자 보호라는 방향은 같지만 규율 대상과 법적 의무가 다름. 세 나라가 모두 같은 '결과책임'을 채택했다고 묶지 않음. 영국: 상품·서비스, 가격·가치, 소비자 이해, 소비자 지원의 네 영역. 판매·갱신 중인 상품은 2023.7.31, 판매 종료 상품은 2024.7.31부터 적용. '좋은 결과'는 투자수익 보장이 아님. 미국: 브로커딜러가 소매고객에게 권유할 때 적용되며 공시·주의·이해상충 관리·준수체계의 네 의무로 구성, 준수기한 2020.6.30. 일본: 일본증권업협회(JSDA) 2023년 자율규제 개정으로 상품명 앞 '복잡한 구조채' 표시와 부적합 고객 유형 표시 예시를 제시. 슬라이드의 단계 구분(상품 설계~사후 지원)은 비교를 위한 팀의 해석임. 한국의 개정안 계류는 2026.11.4 금융위 보도자료 기준.");
  }

  // ================= 10. 문제점 =================
  {
    const s = content(SEC[3]);
    head(s, "주요 문제점", "문제는 상품보다 판매 구조에 있었음",
      "강의자료 4장 3절, 6장 2절, 8장 1·2절 · 금융위원회·금융감독원(2025.2) · 금감원 현장검사 결과(2024.3.11)");
    const cw = 2.1, gap = 0.2;
    const xOf = (i) => 0.5 + i * (cw + gap);
    T(s, "판매자 측", { x: xOf(0), y: 1.25, w: 1.2, h: 0.22, fontSize: 10, bold: true, color: c("loss"), valign: "middle" });
    line(s, xOf(0) + 0.8, 1.36, 3 * cw + 2 * gap - 0.8, c("loss"), 1);
    T(s, "소비자 측", { x: xOf(3), y: 1.25, w: 1.2, h: 0.22, fontSize: 10, bold: true, color: c("neutral"), valign: "middle" });
    line(s, xOf(3) + 0.8, 1.36, cw - 0.8, c("neutral"), 1);
    const P = [
      [fa.FaUniversity, "예금 창구에서\n고위험 상품 권유", "만기 고객에게 같은 창구에서", "8장 2절 예금자보호"],
      [fa.FaClipboardCheck, "형식적인\n적합성 평가", "\"가입하고 싶어요\" 답변 유도", "8장 1절 6대 판매원칙"],
      [fa.FaChartLine, "판매실적\n중심의 KPI", "고위험 상품에 유리한 평가", "6장 2절 · 4장 3절"],
      [fa.FaSignature, "판매직원에 대한\n높은 의존", "직원 설명만 믿고 서명", "8장 1절 소비자의 책무"],
    ];
    for (let i = 0; i < 4; i++) {
      const x = xOf(i), y = 1.6, h = 3.27, col = i < 3 ? "loss" : "neutral";
      card(s, x, y, cw, h);
      T(s, "0" + (i + 1), { x: x + 0.2, y: y + 0.22, w: 0.8, h: 0.54, fontSize: 28, bold: true, color: c("faint"), valign: "middle" });
      await iconDot(s, P[i][0], x + cw - 0.2 - 0.54, y + 0.22, 0.54, c(col));
      T(s, P[i][1], { x: x + 0.2, y: y + 1.1, w: cw - 0.4, h: 0.66, fontSize: 15, bold: true, valign: "middle" });
      T(s, P[i][2], { x: x + 0.2, y: y + 1.9, w: cw - 0.4, h: 0.5, fontSize: 10.5, color: c("muted"), valign: "top" });
      pill(s, P[i][3], x + 0.2, y + h - 0.5, cw - 0.4, c("pillBg"), c(col), { h: 0.28 });
    }
    s.addNotes(NOTES[7] + "\n\n[시각화] 카드 아래 칸은 연결되는 강의자료 위치임. 03번은 6장 2절 이해상충과 4장 3절 수익성과 공공성."
      + "\n\n[슬라이드에서 줄인 설명] 01 예·적금 만기 고객이 같은 창구에서 ELS를 권유받음. 02 \"이 상품에 가입하고 싶어요\"라고 말하도록 직원이 유도. 03 고위험 상품 판매에 유리하게 KPI를 설계하고 공격적 영업 지속. 04 위험을 스스로 판단하기보다 직원 설명만 믿고 서명.");
  }

  // ================= 11. 개선방안 =================
  {
    const s = content(SEC[3]);
    head(s, "개선방안", "네 주체가 각자 바꿔야 할 것",
      "강의자료 6장 4절, 8장 1절 주요내용 5 · 금융위원회·금융감독원(2025.2) · 금융위원회 보도자료(2026.11.4)");
    const tag = (t, r) => ({ text: "   " + t, options: { color: c(r), bold: true, fontSize: 9.5 } });
    const Q = [
      [fa.FaLandmark, "정부·금융당국", "neutral", [[{ text: "최선이익 원칙 법제화" }, tag("국회 계류", "warn")], [{ text: "신속하고 예측 가능한 제재" }, tag("팀 제안", "right")]]],
      [fa.FaUniversity, "금융기관 (판매사)", "loss", [[{ text: "거점점포·전담직원만 판매" }, tag("2025.2 대책", "muted")], [{ text: "고객 손익을 반영한 KPI" }, tag("2025.2 대책", "muted")]]],
      [fa.FaIndustry, "기업 (상품 발행·설계)", "warn", [[{ text: "설명서 맨 위에 손실 사례" }], [{ text: "상품명 앞 '고난도' 표시" }]]],
      [fa.FaUser, "소비자", "right", [[{ text: "예금자보호 여부부터 확인" }], [{ text: "이해 못 한 계약엔 서명하지 않음" }]]],
    ];
    const cw = 4.4, chh = 1.7;
    for (let i = 0; i < 4; i++) {
      const x = 0.5 + (i % 2) * (cw + 0.2), y = 1.3 + Math.floor(i / 2) * (chh + 0.17);
      card(s, x, y, cw, chh);
      await iconDot(s, Q[i][0], x + 0.22, y + 0.2, 0.5, c(Q[i][2]));
      T(s, Q[i][1], { x: x + 0.88, y: y + 0.2, w: cw - 1.05, h: 0.5, fontSize: 15, bold: true, valign: "middle" });
      Q[i][3].forEach((it, k) => {
        const ry = y + 0.88 + k * 0.34;
        dot(s, x + 0.3, ry + 0.12, 0.07, c(Q[i][2]));
        T(s, it.map((r) => ({ text: r.text, options: { ...(r.options || {}) } })), { x: x + 0.5, y: ry, w: cw - 0.7, h: 0.3, fontSize: 12.5, valign: "middle" });
      });
    }
    s.addNotes(NOTES[8] + "\n\n[수정 반영] 항목별 상태를 색 글씨로 구분함. 금소법 개정안은 국회 계류 중(2026.11.4 금융위 발표). 거점점포·전담직원, KPI는 2025.2 대책의 과제이며 모든 기관의 이행 완료 여부까지 확인한 것은 아님. 전담직원 경력 요건(3년 이상)은 후속 규정 확인이 필요해 슬라이드에서 뺌."
      + "\n\n[슬라이드에서 뺀 항목] 정부·금융당국: 투자성 상품 설명 가이드라인(2027년 1분기 시행 예정). 금융기관: 책무구조도로 경영진 책임 명확화. 기업: 부적합 소비자 유형 기재, 전문용어 대신 쉬운 표현. 소비자: 가족 지정인 확인 서비스(65세 이상 중 희망자 대상).");
  }

  // ================= 12. 결론 =================
  pres.addSection({ title: SEC[4] });
  {
    const s = pres.addSlide({ masterName: "Closing", sectionTitle: SEC[4] });
    s.addText("결론", { placeholder: "eyebrow" });
    s.addText("권리를 늘리는 것보다,\n있는 권리가 작동하게 하는 것이 중요함", { placeholder: "title" });
    const K = [
      [fa.FaShieldAlt, "결과", "구제는\n자율배상 중심", "right"],
      [fa.FaUniversity, "원인", "같은 창구 · 형식적 평가\n실적 중심 KPI", "loss"],
      [fa.FaBolt, "과제", "원칙 중심 규제\n신속한 제재", "warn"],
    ];
    const cw = 2.8, gap = 0.2;
    for (let i = 0; i < 3; i++) {
      const x = 0.6 + i * (cw + gap), y = 2.5;
      box(s, x, y, cw, 1.75, c("closeCard"));
      await iconDot(s, K[i][0], x + 0.25, y + 0.25, 0.52, c(K[i][3]));
      T(s, "0" + (i + 1) + "  " + K[i][1], { x: x + 0.92, y: y + 0.25, w: cw - 1.1, h: 0.52, fontSize: 13, bold: true, color: c("closeFg"), valign: "middle" });
      T(s, K[i][2], { x: x + 0.25, y: y + 0.98, w: cw - 0.5, h: 0.6, fontSize: 13.5, bold: true, color: c("closeFg") });
    }
    s.addText("출처   박운학, 「금융경영의 이해」 강의자료 4·6·8장(2026) · 금융위원회·금융감독원, 「홍콩 H지수 ELS 현황 및 대책」(2025.2) · 금융위원회, 「홍콩 H지수 ELS 관련 은행·증권사 검사 결과 조치」(2026.11.4) · 금감원 분쟁조정위원회 대표사례(2024.5) · 금융위원회 위법계약해지권 안내 · 법제처 찾기쉬운 생활법령정보 · FCA FG22/5 · SEC Regulation Best Interest · 일본증권업협회(2023)", { placeholder: "source" });
    s.addNotes(NOTES[9] + "\n\n[세 가지 요약] 01 금소법상 권리는 있었으나 실제 구제는 자율배상 중심으로 진행됨. 02 원인은 판매 구조: 같은 창구, 형식적 적합성 평가, 실적 중심 KPI. 03 권리가 작동하려면 원칙 중심 규제와 신속한 제재가 필요함.");
  }

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
}

(async () => {
  const only = process.argv[3];
  for (const k of Object.keys(VARIANTS)) if (!only || only === k) await build(VARIANTS[k]);
})();
