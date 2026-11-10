// 홍콩 ELS · 소비자 권리 발표 덱 생성 스크립트 (pptxgenjs)
// 실행: PPTX_SKILL_DIR=<pptx skill dir> node build_deck.js [out.pptx]
const path = require("path");
const fs = require("fs");
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
const { applyTheme } = require(path.join(process.env.PPTX_SKILL_DIR, "scripts/apply_theme.js"));

const OUT = process.argv[2] || "deck.pptx";
const NOTES = JSON.parse(fs.readFileSync(path.join(__dirname, "notes_orig.json"), "utf8"));

const THEME = {
  name: "ELS Consumer Rights",
  headFontFace: "Pretendard",
  bodyFontFace: "Pretendard",
  colors: {
    dk1: "0E1726", lt1: "FFFFFF", dk2: "1F2A44", lt2: "F2F4F8",
    accent1: "E5484D", // 손실·제재 (red)
    accent2: "0E9384", // 권리·보호 (teal)
    accent3: "D9942B", // 진행 중·주의 (amber)
    accent4: "5B667A", // 보조 텍스트 (slate)
    accent5: "D5DAE3", // 선·비활성
    accent6: "2B3857", // 어두운 면
    hlink: "0E9384", folHlink: "5B667A",
  },
};
const H = THEME.colors;

async function icon(Comp, hex, size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Comp, { color: "#" + hex, size }));
  const buf = await sharp(Buffer.from(svg)).resize(size, size, { fit: "contain", background: { r: 0, g: 0, b: 0, alpha: 0 } }).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

(async () => {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9"; // 10 x 5.625
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  pres.title = "홍콩 ELS 사태로 알아보는 소비자 권리";
  const C = pres.SchemeColor;
  const CF = "+mn-lt";

  // ---------- layouts ----------
  const ph = (name, type, o) => ({ placeholder: { options: { name, type, margin: 0, ...o }, text: "" } });
  pres.defineSlideMaster({
    title: "Cover", background: { color: C.text1 },
    objects: [
      ph("eyebrow", "body", { x: 0.6, y: 1.45, w: 5.2, h: 0.3, fontSize: 12, bold: true, color: C.accent2, charSpacing: 2 }),
      ph("title", "title", { x: 0.6, y: 1.85, w: 5.3, h: 1.35, fontSize: 31, bold: true, color: C.background1, valign: "top", align: "left" }),
      ph("sub", "body", { x: 0.6, y: 3.4, w: 5.2, h: 0.5, fontSize: 14, color: C.accent5, valign: "top" }),
    ],
  });
  pres.defineSlideMaster({
    title: "Content", background: { color: C.background1 },
    objects: [
      ph("eyebrow", "body", { x: 0.5, y: 0.32, w: 6, h: 0.22, fontSize: 10, bold: true, color: C.accent1, charSpacing: 2 }),
      ph("title", "title", { x: 0.5, y: 0.55, w: 9, h: 0.55, fontSize: 26, bold: true, color: C.text1, valign: "middle", align: "left" }),
      ph("source", "body", { x: 0.5, y: 5.12, w: 8.3, h: 0.28, fontSize: 8, color: C.accent4, valign: "middle" }),
    ],
    slideNumber: { x: 9.0, y: 5.12, w: 0.5, h: 0.28, fontSize: 8, color: C.accent4, align: "right" },
  });
  pres.defineSlideMaster({
    title: "Closing", background: { color: C.text1 },
    objects: [
      ph("eyebrow", "body", { x: 0.6, y: 0.5, w: 6, h: 0.25, fontSize: 10, bold: true, color: C.accent2, charSpacing: 2 }),
      ph("title", "title", { x: 0.6, y: 0.85, w: 8.8, h: 1.2, fontSize: 28, bold: true, color: C.background1, valign: "top", align: "left" }),
      ph("source", "body", { x: 0.6, y: 4.75, w: 8.8, h: 0.55, fontSize: 8, color: C.accent4, valign: "top" }),
    ],
  });

  // ---------- helpers ----------
  let n = 0;
  const nm = (p) => `${p} ${++n}`;
  const T = (s, text, o) => s.addText(text, { isTextBox: true, margin: 0, valign: "top", color: C.text1, objectName: nm("Text"), ...o });
  const card = (s, x, y, w, h, fill, o = {}) =>
    s.addShape(pres.ShapeType.roundRect, { x, y, w, h, rectRadius: 0.08, fill: { color: fill }, line: { type: "none" }, objectName: nm("Card"), ...o });
  const dot = (s, x, y, d, fill, o = {}) =>
    s.addShape(pres.ShapeType.ellipse, { x, y, w: d, h: d, fill: { color: fill }, line: { type: "none" }, objectName: nm("Dot"), ...o });
  const pill = (s, text, x, y, w, fill, color, o = {}) =>
    s.addText(text, { shape: pres.ShapeType.roundRect, rectRadius: 0.11, x, y, w, h: 0.22, fill: { color: fill }, line: { type: "none" },
      fontSize: 8.5, bold: true, color, align: "center", valign: "middle", margin: 0, objectName: nm("Pill"), ...o });
  const iconDot = async (s, Comp, x, y, d, fill, hex = "FFFFFF") => {
    dot(s, x, y, d, fill);
    const k = d * 0.48;
    s.addImage({ data: await icon(Comp, hex), x: x + (d - k) / 2, y: y + (d - k) / 2, w: k, h: k, objectName: nm("Icon") });
  };
  const line = (s, x, y, w, color, pt = 1.5) =>
    s.addShape(pres.ShapeType.line, { x, y, w, h: 0, line: { color, width: pt }, objectName: nm("Line") });
  const chartBase = () => ({
    catAxisLabelColor: H.dk1, catAxisLabelFontFace: CF, catAxisLabelFontSize: 10,
    valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" },
    catAxisLineShow: false, valAxisLineShow: false,
    showValue: true, dataLabelColor: H.dk1, dataLabelFontFace: CF, dataLabelFontSize: 10, dataLabelFontBold: true,
    showLegend: false,
  });
  const head = (s, eyebrow, title, source) => {
    s.addText(eyebrow, { placeholder: "eyebrow" });
    s.addText(title, { placeholder: "title" });
    if (source) s.addText(source, { placeholder: "source" });
  };

  // ================= 1. 표지 =================
  pres.addSection({ title: "도입" });
  {
    const s = pres.addSlide({ masterName: "Cover", sectionTitle: "도입" });
    s.addText("주제 5 · 소비자의 권리", { placeholder: "eyebrow" });
    s.addText("홍콩 ELS 사태로 알아보는\n소비자 권리", { placeholder: "title" });
    s.addText("금융소비자보호법이 준 권리는 실제로 작동했는가", { placeholder: "sub" });
    s.addChart(pres.charts.DOUGHNUT, [{ name: "손실확정 계좌 원금", labels: ["손실", "잔여"], values: [44.2, 55.8] }], {
      x: 6.05, y: 0.95, w: 3.4, h: 3.4, holeSize: 72, chartColors: [H.accent1, H.accent6],
      showLegend: false, showValue: false, showPercent: false, showLabel: false, showTitle: false,
      dataBorder: { pt: 0, color: H.dk1 }, objectName: "손실률 도넛",
    });
    T(s, "44.2%", { x: 6.75, y: 2.1, w: 2.0, h: 0.6, fontSize: 34, bold: true, color: C.background1, align: "center", valign: "middle" });
    T(s, "누적 손실률", { x: 6.75, y: 2.72, w: 2.0, h: 0.25, fontSize: 11, color: C.accent5, align: "center" });
    T(s, "원금 10.4조원 중 4.6조원 손실 · 5대 판매은행, 2024년 말", { x: 5.75, y: 4.5, w: 4.0, h: 0.25, fontSize: 9, color: C.accent4, align: "center" });
    s.addNotes(NOTES[0]);
  }

  // ================= 2. 배경 =================
  pres.addSection({ title: "배경과 권리" });
  {
    const s = pres.addSlide({ masterName: "Content", sectionTitle: "배경과 권리" });
    head(s, "배경", "법이 생긴 뒤에도 더 큰 손실이 발생함",
      "강의자료 8장 1절 · 금융위원회·금융감독원, 「홍콩 H지수 ELS 현황 및 대책」(2025.2) · 원 넓이는 손실액 비례(팀 작성)");
    const ly = 2.9;
    line(s, 0.9, ly, 8.2, C.accent5, 2);
    const cols = [1.9, 4.9, 7.9];
    const yrs = ["2019", "2021.3", "2024"];
    const evs = ["DLF 사태", "금융소비자보호법 시행", "홍콩 H지수 ELS 사태"];
    cols.forEach((cx, i) => {
      T(s, yrs[i], { x: cx - 1.3, y: 1.25, w: 2.6, h: 0.34, fontSize: 20, bold: true, align: "center", color: i === 1 ? C.accent2 : C.text1 });
      T(s, evs[i], { x: cx - 1.3, y: 1.6, w: 2.6, h: 0.25, fontSize: 12, align: "center", color: C.accent4 });
    });
    // DLF: 4,468억 / ELS: 4.6조 → 넓이 비 10.3배, 지름 비 3.21배
    const dE = 1.8, dD = dE / Math.sqrt(46000 / 4468);
    dot(s, cols[0] - dD / 2, ly - dD / 2, dD, C.accent1);
    T(s, "손실 4,468억원", { x: cols[0] - 1.3, y: 3.35, w: 2.6, h: 0.28, fontSize: 14, bold: true, align: "center" });
    T(s, "투자자 3,243명 · 판매 7,950억원", { x: cols[0] - 1.3, y: 3.65, w: 2.6, h: 0.22, fontSize: 10, align: "center", color: C.accent4 });
    await iconDot(s, fa.FaShieldAlt, cols[1] - 0.36, ly - 0.36, 0.72, C.accent2);
    T(s, "6대 판매원칙 도입", { x: cols[1] - 1.3, y: 3.35, w: 2.6, h: 0.28, fontSize: 14, bold: true, align: "center" });
    T(s, "청약철회권 · 위법계약해지권 · 입증책임 전환", { x: cols[1] - 1.7, y: 3.65, w: 3.4, h: 0.22, fontSize: 10, align: "center", color: C.accent4 });
    dot(s, cols[2] - dE / 2, ly - dE / 2, dE, C.accent1);
    T(s, "4.6조원", { x: cols[2] - 0.9, y: ly - 0.42, w: 1.8, h: 0.5, fontSize: 28, bold: true, color: C.background1, align: "center", valign: "middle" });
    T(s, "손실률 44.2%", { x: cols[2] - 0.9, y: ly + 0.1, w: 1.8, h: 0.25, fontSize: 11, color: C.background1, align: "center" });
    T(s, "손실확정 17.0만 계좌", { x: cols[2] - 1.3, y: 3.9, w: 2.6, h: 0.22, fontSize: 10, align: "center", color: C.accent4 });
    card(s, 0.5, 4.35, 9, 0.55, C.background2);
    T(s, [
      { text: "손실 규모 약 10배", options: { bold: true, color: C.accent1 } },
      { text: "   판매규제는 강화됐지만 현장의 소비자 보호는 충실히 작동하지 않음 (금감원 검사 결과)", options: {} },
    ], { x: 0.75, y: 4.35, w: 8.5, h: 0.55, fontSize: 12, valign: "middle" });
    s.addNotes(NOTES[1] + "\n\n[시각화] 원의 넓이는 손실액에 비례함. 4.6조원 ÷ 4,468억원 ≈ 10.3배(팀 계산). DLF는 전체 투자자, ELS는 5대 판매은행 손실확정 계좌 기준이라 집계 범위가 다르다는 점을 질문이 나오면 밝힘.");
  }

  // ================= 3. ELS =================
  {
    const s = pres.addSlide({ masterName: "Content", sectionTitle: "배경과 권리" });
    head(s, "ELS란", "은행 창구에서 팔린 '예금 같은' 투자상품",
      "강의자료 8장 1·2절 · 금융위원회·금융감독원(2025.2), 5대 판매은행 2024년 말 기준");
    const xa = 1.75, wa = 1.85, xb = 3.75, wb = 1.95;
    card(s, xa, 1.3, wa, 3.0, C.background2);
    card(s, xb, 1.3, wb, 3.0, C.text1);
    T(s, "정기예금", { x: xa, y: 1.4, w: wa, h: 0.4, fontSize: 14, bold: true, align: "center", valign: "middle" });
    T(s, "ELS", { x: xb, y: 1.4, w: wb, h: 0.4, fontSize: 14, bold: true, align: "center", valign: "middle", color: C.background1 });
    const rows = [
      ["판매 창구", "은행", "은행", null],
      ["원금", "보장", "손실 가능", true],
      ["예금자보호", "대상", "대상 아님", true],
      ["수익", "확정 이자", "지수 연계", true],
    ];
    const ok = await icon(fa.FaCheckCircle, H.accent2), no = await icon(fa.FaTimesCircle, H.accent1), same = await icon(fa.FaEquals, H.accent4), sameL = await icon(fa.FaEquals, H.accent5);
    rows.forEach((r, i) => {
      const y = 1.9 + i * 0.58;
      if (i > 0) { line(s, 0.5, y, 1.25 + wa, C.accent5, 0.75); }
      T(s, r[0], { x: 0.5, y, w: 1.15, h: 0.58, fontSize: 11, color: C.accent4, valign: "middle" });
      const ia = r[3] ? ok : same, ib = r[3] ? no : same;
      s.addImage({ data: ia, x: xa + 0.22, y: y + 0.19, w: 0.2, h: 0.2, objectName: nm("Icon") });
      T(s, r[1], { x: xa + 0.52, y, w: wa - 0.6, h: 0.58, fontSize: 13, bold: true, valign: "middle" });
      s.addImage({ data: r[3] ? ib : sameL, x: xb + 0.22, y: y + 0.19, w: 0.2, h: 0.2, objectName: nm("Icon") });
      T(s, r[2], { x: xb + 0.52, y, w: wb - 0.6, h: 0.58, fontSize: 13, bold: true, valign: "middle", color: C.background1 });
    });
    T(s, "은행은 신탁(ELT)·펀드(ELF) 형태로 판매함. 대표적 구조에서는 만기평가일의 기초지수와 상환조건에 따라 원금손실이 발생할 수 있음", {
      x: 0.5, y: 4.42, w: 5.2, h: 0.45, fontSize: 10, color: C.accent4 });
    // 오른쪽: 손실 규모
    const rx = 6.1, rw = 3.4;
    card(s, rx, 1.3, rw, 3.57, C.background2);
    T(s, "손실확정 계좌", { x: rx + 0.3, y: 1.5, w: rw - 0.6, h: 0.22, fontSize: 10, color: C.accent4 });
    T(s, "17.0만", { x: rx + 0.3, y: 1.72, w: rw - 0.6, h: 0.55, fontSize: 32, bold: true, valign: "middle" });
    T(s, "손실액", { x: rx + 0.3, y: 2.45, w: rw - 0.6, h: 0.22, fontSize: 10, color: C.accent4 });
    T(s, "4.6조원", { x: rx + 0.3, y: 2.67, w: rw - 0.6, h: 0.55, fontSize: 32, bold: true, color: C.accent1, valign: "middle" });
    const bw = rw - 0.6;
    s.addShape(pres.ShapeType.rect, { x: rx + 0.3, y: 3.45, w: bw, h: 0.3, fill: { color: C.accent5 }, line: { type: "none" }, objectName: "원금 막대" });
    s.addShape(pres.ShapeType.rect, { x: rx + 0.3, y: 3.45, w: bw * 0.442, h: 0.3, fill: { color: C.accent1 }, line: { type: "none" }, objectName: "손실 막대" });
    T(s, "44.2%", { x: rx + 0.3, y: 3.45, w: bw * 0.442, h: 0.3, fontSize: 10, bold: true, color: C.background1, align: "center", valign: "middle" });
    T(s, "원금 10.4조원", { x: rx + 0.3, y: 3.82, w: bw, h: 0.22, fontSize: 10, color: C.accent4, align: "right" });
    T(s, "2021년 초 H지수 고점 판매분의 3년 만기가 2024년에 도래", { x: rx + 0.3, y: 4.2, w: bw, h: 0.45, fontSize: 10, color: C.text1 });
    s.addNotes(NOTES[2] + "\n\n[수정 반영] 손실 조건은 \"대표적인 구조에서는 만기평가일의 기초지수와 정해진 상환조건 등에 따라 원금손실이 발생할 수 있다\"로 설명함. 장중 하락만으로 손실이 확정된다고 말하지 않음.");
  }

  // ================= 4. 권리 =================
  {
    const s = pres.addSlide({ masterName: "Content", sectionTitle: "배경과 권리" });
    head(s, "금융소비자보호법", "금소법이 소비자에게 준 여섯 가지 장치",
      "강의자료 8장 1절 주요내용 2~4 · 금소법 제28·42·44·46·47·57조 · 법제처 찾기쉬운 생활법령정보");
    const G = { a: ["권리 강화", C.accent2], b: ["분쟁 시 권익 확보", C.text2], c: ["처벌 강화", C.accent1] };
    const items = [
      [fa.FaUndoAlt, "청약철회권", "대상 투자성 상품·자문은 원칙적으로 7일 안에 철회", "a"],
      [fa.FaFileContract, "위법계약해지권", "판매원칙 위반 계약을 수수료·위약금 없이 해지", "a"],
      [fa.FaHandshake, "조정이탈금지", "2,000만원 이하 분쟁조정 중 금융회사의 제소 제한", "b"],
      [fa.FaFolderOpen, "자료열람요구권", "분쟁조정·소송을 위해 자료 열람·사본 요구", "b"],
      [fa.FaBalanceScale, "입증책임 전환", "설명의무 위반 시 금융회사가 고의·과실 없음을 입증", "b"],
      [fa.FaGavel, "징벌적 과징금", "판매원칙 위반 시 관련 수입 등의 50% 이내", "c"],
    ];
    const cw = 2.85, gap = 0.225, chh = 1.65;
    for (let i = 0; i < items.length; i++) {
      const [ic, name, desc, g] = items[i];
      const x = 0.5 + (i % 3) * (cw + gap), y = 1.3 + Math.floor(i / 3) * (chh + 0.2);
      card(s, x, y, cw, chh, C.background2);
      await iconDot(s, ic, x + 0.22, y + 0.2, 0.52, G[g][1]);
      T(s, G[g][0], { x: x + 0.95, y: y + 0.2, w: cw - 1.15, h: 0.52, fontSize: 9, bold: true, color: G[g][1], align: "right", valign: "middle" });
      T(s, name, { x: x + 0.22, y: y + 0.85, w: cw - 0.44, h: 0.3, fontSize: 15, bold: true, valign: "middle" });
      T(s, desc, { x: x + 0.22, y: y + 1.17, w: cw - 0.44, h: 0.4, fontSize: 10.5, color: C.accent4 });
    }
    s.addNotes(NOTES[3] + "\n\n[수정 반영] ① 청약철회는 \"대상 투자성 상품·자문은 원칙적으로 7일\"(법 제46조, 시행령 제37조)로 한정함. 철회기간 중 자금 운용에 동의한 경우 등 예외가 있어 모든 ELS 거래에 자동 적용되지 않음. ② '소송이탈금지'는 '조정이탈금지'로 표기함(법 제42조, 조정안 제시 전이 원칙이며 예외 존재). ③ 위법계약해지권의 기간은 위반을 안 날부터 1년·계약일부터 5년 이내. ④ 입증책임 전환은 설명의무 위반 손해배상에 한정됨(법 제44조 제2항). ⑤ 과징금은 적합성·적정성 위반이 제외됨(슬라이드에서는 생략, 질문 시 답변).");
  }

  // ================= 5. 실제 구제 =================
  pres.addSection({ title: "실제 결과" });
  {
    const s = pres.addSlide({ masterName: "Content", sectionTitle: "실제 결과" });
    head(s, "실제 구제", "실제 구제는 자율배상 중심으로 진행됨",
      "금융위원회·금융감독원(2025.2), 5대 판매은행 2024년 말 기준 · 금감원 분쟁조정위원회 대표사례(2024.5.13 의결)");
    card(s, 0.5, 1.3, 4.3, 3.57, C.background2);
    T(s, "평균 배상비율", { x: 0.75, y: 1.48, w: 3.8, h: 0.22, fontSize: 10, color: C.accent4 });
    T(s, "31.4%", { x: 0.75, y: 1.7, w: 3.8, h: 0.75, fontSize: 46, bold: true, color: C.accent1, valign: "middle" });
    s.addChart(pres.charts.BAR, [
      { name: "배상액 (조원)", labels: ["손실 4.1조원"], values: [1.3] },
      { name: "소비자 부담 (조원)", labels: ["손실 4.1조원"], values: [2.8] },
    ], {
      x: 0.62, y: 2.5, w: 4.06, h: 1.05, barDir: "bar", barGrouping: "stacked", barGapWidthPct: 30,
      chartColors: [H.accent1, H.accent5], catAxisHidden: true, valAxisHidden: true, valAxisMaxVal: 4.1,
      valGridLine: { style: "none" }, catGridLine: { style: "none" },
      showValue: true, dataLabelPosition: "ctr", dataLabelFormatCode: '0.0', dataLabelColor: H.dk1, dataLabelFontFace: CF, dataLabelFontSize: 11, dataLabelFontBold: true,
      showLegend: true, legendPos: "b", legendFontFace: CF, legendFontSize: 9, legendColor: H.accent4,
      objectName: "배상 구성 차트",
    });
    T(s, "배상 동의 15.9만 계좌 · 손실 약 4.1조원 기준", { x: 0.75, y: 3.6, w: 3.8, h: 0.22, fontSize: 9.5, color: C.accent4 });
    pill(s, "금감원 분쟁조정기준 2024.3.11", 0.75, 4.12, 2.05, C.text2, C.background1, { h: 0.34, fontSize: 9.5 });
    T(s, "→", { x: 2.8, y: 4.12, w: 0.35, h: 0.34, fontSize: 14, bold: true, align: "center", valign: "middle", color: C.accent4 });
    pill(s, "은행 자율배상", 3.15, 4.12, 1.4, C.accent1, C.background1, { h: 0.34, fontSize: 9.5 });
    // 오른쪽
    T(s, "대표사례 배상비율은 30~65%", { x: 5.1, y: 1.3, w: 4.4, h: 0.3, fontSize: 14, bold: true, valign: "middle" });
    T(s, "판매사 책임에 투자자 개별 사정을 반영", { x: 5.1, y: 1.6, w: 4.4, h: 0.22, fontSize: 10, color: C.accent4 });
    s.addChart(pres.charts.BAR, [{ name: "배상비율(%)", labels: ["농협", "국민", "신한", "SC제일", "하나"], values: [65, 60, 55, 55, 30] }], {
      ...chartBase(), x: 5.0, y: 1.85, w: 4.5, h: 1.75, barDir: "bar", catAxisOrientation: "maxMin", barGapWidthPct: 45,
      chartColors: [H.dk2], valAxisMaxVal: 100, valAxisMinVal: 0, dataLabelPosition: "outEnd", dataLabelFormatCode: '0"%"',
      objectName: "대표사례 배상비율 차트",
    });
    card(s, 5.1, 3.77, 4.4, 1.1, C.text1);
    await iconDot(s, fa.FaExclamation, 5.3, 4.06, 0.52, C.accent3);
    T(s, "위법계약해지권의 한계", { x: 6.02, y: 3.93, w: 3.3, h: 0.25, fontSize: 12, bold: true, color: C.background1, valign: "middle" });
    T(s, "해지는 장래를 향해서만 효력이 있어, 이미 확정된 손실을 되돌리는 장치가 아님", { x: 6.02, y: 4.2, w: 3.3, h: 0.5, fontSize: 10.5, color: C.accent5 });
    s.addNotes(NOTES[4].replace("피해 구제는 강의에서 배운 권리 행사가 아니라 금감원 분쟁조정기준과 은행 자율배상으로 이루어짐", "피해 구제는 금감원 분쟁조정기준에 따른 은행 자율배상 중심으로 진행됨")
      + "\n\n[수정 반영] 제목을 '자율배상 중심으로 진행됨'으로 바꿈. 분쟁조정·손해배상도 권리구제 수단이므로 권리 행사 전체를 부정하지 않음. 31.4%는 배상 동의 15.9만 계좌(배상진행 16.9만건의 93.8%) 기준이며, 전체 피해자에게 똑같이 적용된 확정 부담률이 아님."
      + "\n\n[대표사례] 2024.5.13 분조위 의결(보도자료 5.14): 농협 65% / 국민 60% / 신한 55% / SC제일 55% / 하나 30%. 이 5건은 2021.3.24 이전 판매분이라 '금소법 시행 후 판매 사례'라고 말하면 안 됨. 배상비율은 판매사 책임에 투자자의 경험·가입 목적 등 개별 사정을 반영한 것이므로, 낮은 배상비율만으로 보호 실패가 입증되는 것은 아님(시장손실과 판매절차 위반 책임을 구분).");
  }

  // ================= 6. 제재 =================
  {
    const s = pres.addSlide({ masterName: "Content", sectionTitle: "실제 결과" });
    head(s, "제재 현황", "최종 제재 3,746억원, 첫 검사 후 약 2년 10개월",
      "강의자료 8장 1절 주요내용 4 · 금융위원회·금융감독원(2025.2) · 금융위원회 보도자료(2026.11.4, 제19차 정례회의 의결)");
    const xs = [1.2, 3.05, 4.9, 6.75, 8.6], ty = 1.5;
    line(s, xs[0], ty, xs[4] - xs[0], C.accent5, 2);
    const tl = [["2024.1~3", "금감원 현장검사"], ["2024.3.11", "분쟁조정기준 마련"], ["2024.5.13", "대표사례 배상비율 의결"], ["2025.2", "재발방지 대책 발표"], ["2026.11.4", "과징금·과태료 최종 의결"]];
    tl.forEach((t, i) => {
      const last = i === 4, d = last ? 0.26 : 0.16;
      dot(s, xs[i] - d / 2, ty - d / 2, d, last ? C.accent1 : C.text2);
      T(s, t[0], { x: xs[i] - 0.95, y: 1.72, w: 1.9, h: 0.24, fontSize: 11, bold: true, align: "center", color: last ? C.accent1 : C.text1 });
      T(s, t[1], { x: xs[i] - 0.95, y: 1.96, w: 1.9, h: 0.22, fontSize: 9.5, align: "center", color: C.accent4 });
    });
    card(s, 0.5, 2.55, 3.9, 2.32, C.text1);
    T(s, "은행 7곳 · 증권사 6곳", { x: 0.78, y: 2.75, w: 3.4, h: 0.22, fontSize: 10, color: C.accent5 });
    T(s, "3,746억원", { x: 0.78, y: 2.98, w: 3.4, h: 0.7, fontSize: 40, bold: true, color: C.background1, valign: "middle" });
    T(s, "과징금 3,712.42억 + 과태료 33.90억", { x: 0.78, y: 3.72, w: 3.4, h: 0.24, fontSize: 11, color: C.background1 });
    pill(s, "금융위 최대 규모 금전 제재", 0.78, 4.12, 1.95, C.accent1, C.background1, { h: 0.28, fontSize: 9 });
    T(s, "주요 금융기관 대상 금소법 위반 첫 사례", { x: 0.78, y: 4.48, w: 3.4, h: 0.22, fontSize: 9.5, color: C.accent5 });
    T(s, "은행별 과징금 (억원)", { x: 4.75, y: 2.55, w: 4.7, h: 0.25, fontSize: 11, bold: true, valign: "middle" });
    s.addChart(pres.charts.BAR, [{ name: "과징금(억원)", labels: ["국민", "하나", "농협", "신한", "SC제일", "우리"], values: [1723.34, 924.75, 401.75, 387.45, 179.0, 3.31] }], {
      ...chartBase(), x: 4.65, y: 2.8, w: 4.85, h: 2.1, barDir: "bar", catAxisOrientation: "maxMin", barGapWidthPct: 45,
      chartColors: [H.accent1], valAxisMaxVal: 2100, valAxisMinVal: 0, dataLabelPosition: "outEnd", dataLabelFormatCode: "#,##0",
      objectName: "은행별 과징금 차트",
    });
    s.addNotes(NOTES[5] + "\n\n[시각화] 첫 현장검사(2024.1)부터 최종 의결(2026.11.4)까지 약 2년 10개월(팀 계산). 합계는 3,746.32억원. '첫 사례'는 금융위 표현대로 \"다수 금융소비자와 관련되고 주요 금융기관이 대상이 된 금소법 위반 첫 사례\"라고 말함. 수협은행은 과징금 없이 과태료만 부과되어 차트에 없음.");
  }

  // ================= 7. 해외 =================
  pres.addSection({ title: "원인과 대안" });
  {
    const s = pres.addSlide({ masterName: "Content", sectionTitle: "원인과 대안" });
    head(s, "해외 사례", "방향은 같지만, 규율하는 지점은 다름",
      "FCA FG22/5 · SEC Regulation Best Interest · 일본증권업협회 자율규제 개정(2023) · 금융위원회·금융감독원(2025.2) 참고1·2 · 단계 구분은 팀 작성");
    const gx = 3.75, gw = 5.75, st = ["상품 설계", "설명·표시", "권유", "사후 지원"], sw = gw / 4;
    st.forEach((t, i) => {
      T(s, t, { x: gx + i * sw, y: 1.3, w: sw, h: 0.3, fontSize: 10.5, bold: true, align: "center", valign: "middle", color: C.accent4 });
    });
    const rowsD = [
      ["UK", "FCA Consumer Duty", "영국 · 2023.7.31 적용", 0, 4, "고객에게 좋은 결과: 상품·가격·이해·지원", C.accent2],
      ["US", "SEC Regulation Best Interest", "미국 · 준수기한 2020.6.30", 2, 1, "고객 최선이익", C.text2],
      ["JP", "복잡한 구조채 표시", "일본 · 2023 자율규제 개정", 1, 1, "눈에 띄는 경고", C.accent3],
    ];
    rowsD.forEach((r, i) => {
      const y = 1.7 + i * 0.9;
      card(s, 0.5, y, 9, 0.78, C.background2);
      s.addText(r[0], { shape: pres.ShapeType.ellipse, x: 0.68, y: y + 0.15, w: 0.48, h: 0.48, fill: { color: r[6] }, line: { type: "none" },
        fontSize: 10, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0, objectName: nm("Badge") });
      T(s, r[1], { x: 1.3, y: y + 0.14, w: 2.4, h: 0.26, fontSize: 12, bold: true, valign: "middle" });
      T(s, r[2], { x: 1.3, y: y + 0.42, w: 2.4, h: 0.22, fontSize: 9.5, color: C.accent4 });
      for (let k = 1; k < 4; k++) s.addShape(pres.ShapeType.line, { x: gx + k * sw, y: y + 0.12, w: 0, h: 0.54, line: { color: C.accent5, width: 0.75, dashType: "dash" }, objectName: nm("Grid") });
      s.addText(r[5], { shape: pres.ShapeType.roundRect, rectRadius: 0.17, x: gx + r[3] * sw + 0.08, y: y + 0.22, w: r[4] * sw - 0.16, h: 0.34,
        fill: { color: r[6] }, line: { type: "none" }, fontSize: 10, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0, objectName: nm("Span") });
    });
    card(s, 0.5, 4.42, 9, 0.48, C.text1);
    T(s, [
      { text: "한국", options: { bold: true, color: C.accent3 } },
      { text: "   최선이익 원칙을 담은 금소법 개정안이 국회에 계류 중 (2026.11.4 금융위)", options: { color: C.background1 } },
    ], { x: 0.75, y: 4.42, w: 8.5, h: 0.48, fontSize: 12, valign: "middle" });
    s.addNotes(NOTES[6] + "\n\n[수정 반영] 세 제도는 소비자 보호라는 방향은 같지만 규율 대상과 법적 의무가 다름. 세 나라가 모두 같은 '결과책임'을 채택했다고 묶지 않음. 영국: 판매·갱신 중인 상품은 2023.7.31, 판매 종료 상품은 2024.7.31부터 적용. '좋은 결과'는 투자수익 보장이 아님. 미국: 브로커딜러가 소매고객에게 권유할 때 적용되며 공시·주의·이해상충 관리·준수체계의 네 의무로 구성, 준수기한 2020.6.30. 일본: 일본증권업협회(JSDA) 2023년 자율규제 개정으로 상품명 앞 '복잡한 구조채' 표시와 부적합 고객 유형 표시 예시를 제시. 슬라이드의 단계 구분(상품 설계~사후 지원)은 비교를 위한 팀의 해석임.");
  }

  // ================= 8. 문제점 =================
  {
    const s = pres.addSlide({ masterName: "Content", sectionTitle: "원인과 대안" });
    head(s, "주요 문제점", "문제는 상품보다 판매 구조에 있었음",
      "강의자료 4장 3절, 6장 2절, 8장 1·2절 · 금융위원회·금융감독원(2025.2) · 금감원 현장검사 결과(2024.3.11)");
    const cw = 2.1, gap = 0.2;
    const xOf = (i) => 0.5 + i * (cw + gap);
    T(s, "판매자 측", { x: xOf(0), y: 1.25, w: 1.2, h: 0.22, fontSize: 10, bold: true, color: C.accent1, valign: "middle" });
    line(s, xOf(0) + 0.8, 1.36, 3 * cw + 2 * gap - 0.8, C.accent1, 1);
    T(s, "소비자 측", { x: xOf(3), y: 1.25, w: 1.2, h: 0.22, fontSize: 10, bold: true, color: C.text2, valign: "middle" });
    line(s, xOf(3) + 0.8, 1.36, cw - 0.8, C.text2, 1);
    const P = [
      [fa.FaUniversity, "예금 창구에서\n고위험 상품 권유", "예·적금 만기 고객이 같은 창구에서 ELS를 권유받음", "8장 2절 예금자보호"],
      [fa.FaClipboardCheck, "형식적인\n적합성 평가", "\"가입하고 싶어요\"라고 말하도록 직원이 유도", "8장 1절 6대 판매원칙"],
      [fa.FaChartLine, "판매실적\n중심의 KPI", "고위험 상품 판매에 유리한 KPI로 공격적 영업 지속", "6장 2절 · 4장 3절"],
      [fa.FaSignature, "판매직원에 대한\n높은 의존", "위험을 따져 보기보다 직원 설명만 믿고 서명", "8장 1절 소비자의 책무"],
    ];
    for (let i = 0; i < 4; i++) {
      const x = xOf(i), y = 1.6, h = 3.27, col = i < 3 ? C.accent1 : C.text2;
      card(s, x, y, cw, h, C.background2);
      T(s, "0" + (i + 1), { x: x + 0.2, y: y + 0.22, w: 0.8, h: 0.5, fontSize: 26, bold: true, color: C.accent5, valign: "middle" });
      await iconDot(s, P[i][0], x + cw - 0.2 - 0.52, y + 0.2, 0.52, col);
      T(s, P[i][1], { x: x + 0.2, y: y + 0.95, w: cw - 0.4, h: 0.62, fontSize: 14, bold: true, valign: "middle" });
      T(s, P[i][2], { x: x + 0.2, y: y + 1.7, w: cw - 0.4, h: 0.75, fontSize: 11.5, color: C.accent4 });
      pill(s, P[i][3], x + 0.2, y + h - 0.5, cw - 0.4, C.background1, col, { h: 0.28 });
    }
    s.addNotes(NOTES[7] + "\n\n[시각화] 카드 아래 칸은 연결되는 강의자료 위치임. 03번은 6장 2절 이해상충과 4장 3절 수익성과 공공성.");
  }

  // ================= 9. 개선방안 =================
  {
    const s = pres.addSlide({ masterName: "Content", sectionTitle: "원인과 대안" });
    head(s, "개선방안", "네 주체가 각자 바꿔야 할 것",
      "강의자료 6장 4절, 8장 1절 주요내용 5 · 금융위원회·금융감독원(2025.2) · 금융위원회 보도자료(2026.11.4)");
    const tag = (t, c) => ({ text: "  " + t, options: { color: c, bold: true, fontSize: 9 } });
    const Q = [
      [fa.FaLandmark, "정부·금융당국", C.text2, [
        [{ text: "최선이익 원칙을 담은 금소법 개정" }, tag("국회 계류", C.accent3)],
        [{ text: "투자성 상품 설명 가이드라인" }, tag("2027년 1분기 예정", C.accent3)],
        [{ text: "신속하고 예측 가능한 제재" }, tag("팀 제안", C.accent2)],
      ]],
      [fa.FaUniversity, "금융기관 (판매사)", C.accent1, [
        [{ text: "거점점포·전담직원만 ELS 판매" }, tag("2025.2 대책", C.accent4)],
        [{ text: "고객 손익을 반영한 KPI 재설계" }, tag("2025.2 대책", C.accent4)],
        [{ text: "책무구조도로 경영진 책임 명확화" }],
      ]],
      [fa.FaIndustry, "기업 (상품 발행·설계)", C.accent3, [
        [{ text: "설명서 맨 위에 부적합 유형·손실 사례" }],
        [{ text: "전문용어 대신 쉬운 표현" }],
        [{ text: "상품명 앞 '고난도 금융투자상품' 표시" }],
      ]],
      [fa.FaUser, "소비자", C.accent2, [
        [{ text: "예금자보호 여부부터 확인" }],
        [{ text: "이해하지 못한 계약에는 서명하지 않음" }],
        [{ text: "가족 지정인 확인 서비스" }, tag("65세 이상 희망자", C.accent4)],
      ]],
    ];
    const cw = 4.4, chh = 1.7;
    for (let i = 0; i < 4; i++) {
      const x = 0.5 + (i % 2) * (cw + 0.2), y = 1.3 + Math.floor(i / 2) * (chh + 0.17);
      card(s, x, y, cw, chh, C.background2);
      await iconDot(s, Q[i][0], x + 0.22, y + 0.18, 0.46, Q[i][2]);
      T(s, Q[i][1], { x: x + 0.82, y: y + 0.18, w: cw - 1.0, h: 0.46, fontSize: 14, bold: true, valign: "middle" });
      Q[i][3].forEach((it, k) => {
        const ry = y + 0.8 + k * 0.27;
        dot(s, x + 0.28, ry + 0.1, 0.06, Q[i][2]);
        T(s, it.map((r) => ({ text: r.text, options: { ...(r.options || {}) } })), { x: x + 0.45, y: ry, w: cw - 0.65, h: 0.26, fontSize: 11, valign: "middle" });
      });
    }
    s.addNotes(NOTES[8] + "\n\n[수정 반영] 항목별 상태를 색 글씨로 구분함. 금소법 개정안은 국회 계류 중, 설명 가이드라인은 2027년 1분기 시행 예정(둘 다 2026.11.4 금융위 발표). 거점점포·전담직원, KPI, 상품명 표시는 2025.2 대책의 과제이며 모든 기관의 이행 완료 여부까지 확인한 것은 아님. 전담직원 경력 요건(3년 이상)은 후속 규정 확인이 필요해 슬라이드에서 뺌. 가족 지정인 확인은 65세 이상 중 희망자 대상.");
  }

  // ================= 10. 결론 =================
  pres.addSection({ title: "결론" });
  {
    const s = pres.addSlide({ masterName: "Closing", sectionTitle: "결론" });
    s.addText("결론", { placeholder: "eyebrow" });
    s.addText("권리를 늘리는 것보다,\n있는 권리가 작동하게 하는 것이 중요함", { placeholder: "title" });
    const K = [
      [fa.FaShieldAlt, "결과", "권리는 있었으나, 실제 구제는 자율배상 중심으로 진행됨", C.accent2],
      [fa.FaUniversity, "원인", "판매 구조: 같은 창구, 형식적 적합성 평가, 실적 중심 KPI", C.accent1],
      [fa.FaBolt, "과제", "원칙 중심 규제와 신속한 제재가 필요함", C.accent3],
    ];
    const cw = 2.8, gap = 0.2;
    for (let i = 0; i < 3; i++) {
      const x = 0.6 + i * (cw + gap), y = 2.45;
      card(s, x, y, cw, 1.95, C.accent6);
      await iconDot(s, K[i][0], x + 0.25, y + 0.25, 0.5, K[i][3]);
      T(s, "0" + (i + 1) + "  " + K[i][1], { x: x + 0.9, y: y + 0.25, w: cw - 1.1, h: 0.5, fontSize: 13, bold: true, color: C.background1, valign: "middle" });
      T(s, K[i][2], { x: x + 0.25, y: y + 0.95, w: cw - 0.5, h: 0.85, fontSize: 12.5, color: C.background1 });
    }
    s.addText("출처   박운학, 「금융경영의 이해」 강의자료 4·6·8장(2026) · 금융위원회·금융감독원, 「홍콩 H지수 ELS 현황 및 대책」(2025.2) · 금융위원회, 「홍콩 H지수 ELS 관련 은행·증권사 검사 결과 조치」(2026.11.4) · 금감원 분쟁조정위원회 대표사례(2024.5) · 금융위원회 위법계약해지권 안내 · 법제처 찾기쉬운 생활법령정보 · FCA FG22/5 · SEC Regulation Best Interest · 일본증권업협회(2023)", { placeholder: "source" });
    s.addNotes(NOTES[9]);
  }

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
