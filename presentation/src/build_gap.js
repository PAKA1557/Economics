// 홍콩 ELS · 소비자 권리 발표 덱 — 시안 D "제도와 현실의 간극" (12장)
// 시각 규칙: 제도·보호 장치 = 가는 윤곽선 / 실제 손실·부담 = 채운 면과 굵은 숫자 / 작동하지 않은 연결 = 점선
// 실행: PPTX_SKILL_DIR=<pptx skill dir> NODE_PATH=<node_modules> node build_gap.js <out.pptx>
const path = require("path");
const fs = require("fs");
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
const { applyTheme } = require(path.join(process.env.PPTX_SKILL_DIR, "scripts/apply_theme.js"));

const OUT = process.argv[2] || "deck.pptx";
const NOTES = JSON.parse(fs.readFileSync(path.join(__dirname, "notes_v3.json"), "utf8"));

const THEME = {
  name: "ELS Gap", headFontFace: "Pretendard", bodyFontFace: "Pretendard",
  colors: {
    dk1: "161A21", lt1: "FFFFFF", dk2: "161A21", lt2: "ECEEF1", // lt2 = 회색 배경
    accent1: "BF3A30", // 실제 손실·부담 (채운 면)
    accent2: "161A21", accent3: "8B929C",
    accent4: "5E6570", // 보조 텍스트
    accent5: "C4C8CF", // 가는 선
    accent6: "DADDE2", hlink: "161A21", folHlink: "5E6570",
  },
};
const H = THEME.colors;
const BG = "bg2", INK = "tx1", LOSS = "accent1", MUTED = "accent4", HAIR = "accent5", WHITE = "bg1";

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
  const CF = "+mn-lt";
  const L = 0.6, R = 9.4, CW = R - L;

  const ph = (name, type, o) => ({ placeholder: { options: { name, type, margin: 0, ...o }, text: "" } });
  pres.defineSlideMaster({
    title: "Cover", background: { color: BG },
    objects: [
      ph("eyebrow", "body", { x: L, y: 0.7, w: 6, h: 0.28, fontSize: 11, bold: true, color: MUTED, charSpacing: 2 }),
      ph("title", "title", { x: L, y: 1.2, w: CW, h: 1.75, fontSize: 44, bold: true, color: INK, valign: "top", align: "left" }),
      ph("label", "body", { x: L, y: 3.55, w: 3, h: 0.25, fontSize: 10.5, bold: true, color: MUTED, charSpacing: 2 }),
      ph("question", "body", { x: L, y: 3.85, w: 6.2, h: 0.6, fontSize: 22, bold: true, color: INK, valign: "middle" }),
    ],
  });
  pres.defineSlideMaster({
    title: "Content", background: { color: BG },
    objects: [
      ph("eyebrow", "body", { x: L, y: 0.34, w: 6, h: 0.22, fontSize: 10, bold: true, color: MUTED, charSpacing: 2 }),
      ph("title", "title", { x: L, y: 0.58, w: CW, h: 0.55, fontSize: 26, bold: true, color: INK, valign: "middle", align: "left" }),
      ph("source", "body", { x: L, y: 5.12, w: 8.2, h: 0.28, fontSize: 8, color: MUTED, valign: "middle" }),
    ],
    slideNumber: { x: 8.9, y: 5.12, w: 0.5, h: 0.28, fontSize: 8, color: MUTED, align: "right" },
  });
  pres.defineSlideMaster({
    title: "Closing", background: { color: BG },
    objects: [
      ph("eyebrow", "body", { x: L, y: 0.7, w: 6, h: 0.28, fontSize: 11, bold: true, color: MUTED, charSpacing: 2 }),
      ph("title", "title", { x: L, y: 1.15, w: CW, h: 1.3, fontSize: 32, bold: true, color: INK, valign: "top", align: "left" }),
      ph("source", "body", { x: L, y: 4.85, w: CW, h: 0.5, fontSize: 7.5, color: MUTED, valign: "top" }),
    ],
  });

  // ---------- helpers ----------
  let n = 0;
  const nm = (p) => `${p} ${++n}`;
  const NOFILL = () => ({ color: BG, transparency: 100 });
  const T = (s, text, o) => s.addText(text, { isTextBox: true, margin: 0, valign: "top", color: INK, objectName: nm("Text"), ...o });
  const hair = (s, x, y, w, color = HAIR, pt = 0.75, extra = {}) =>
    s.addShape(pres.ShapeType.line, { x, y, w, h: 0, line: { color, width: pt, ...extra }, objectName: nm("Line") });
  const vhair = (s, x, y, h, color = HAIR, pt = 0.75, extra = {}) =>
    s.addShape(pres.ShapeType.line, { x, y, w: 0, h, line: { color, width: pt, ...extra }, objectName: nm("Line") });
  // 제도 = 윤곽선 원 / 현실 = 채운 원
  const ring = (s, x, y, d, o = {}) =>
    s.addShape(pres.ShapeType.ellipse, { x, y, w: d, h: d, fill: { color: BG }, line: { color: INK, width: 1.25, ...(o.dash ? { dashType: "dash" } : {}) }, objectName: nm("제도 원") });
  const disc = (s, x, y, d, color = LOSS) =>
    s.addShape(pres.ShapeType.ellipse, { x, y, w: d, h: d, fill: { color }, line: { type: "none" }, objectName: nm("현실 원") });
  const ringIcon = async (s, Comp, x, y, d) => {
    ring(s, x, y, d);
    const k = d * 0.46;
    s.addImage({ data: await icon(Comp, H.dk1), x: x + (d - k) / 2, y: y + (d - k) / 2, w: k, h: k, objectName: nm("Icon") });
  };
  const outBox = (s, text, x, y, w, h, o = {}) =>
    s.addText(text, { shape: pres.ShapeType.rect, x, y, w, h, fill: NOFILL(), line: { color: INK, width: 1.25, ...(o.dash ? { dashType: "dash" } : {}) },
      fontSize: o.fontSize || 11.5, bold: true, color: INK, align: "center", valign: "middle", margin: 0, objectName: nm("제도 상자") });
  const head = (s, eyebrow, title, source) => {
    s.addText(eyebrow, { placeholder: "eyebrow" });
    s.addText(title, { placeholder: "title" });
    if (source) s.addText(source, { placeholder: "source" });
  };
  const motif = (s, x, y, connected) => { // 제도 ○ ─ ─ ● 현실
    ring(s, x, y, 0.5);
    hair(s, x + 0.5, y + 0.25, connected ? 1.2 : 0.95, INK, 1.25, connected ? {} : { dashType: "dash" });
    disc(s, x + 1.7, y, 0.5, connected ? INK : LOSS);
    T(s, "제도", { x: x - 0.25, y: y + 0.6, w: 1.0, h: 0.22, fontSize: 10, color: MUTED, align: "center" });
    T(s, "현실", { x: x + 1.45, y: y + 0.6, w: 1.0, h: 0.22, fontSize: 10, color: MUTED, align: "center" });
  };
  const SEC = ["도입", "배경과 권리", "실제 결과", "원인과 대안", "결론"];
  const content = (sec) => pres.addSlide({ masterName: "Content", sectionTitle: sec });

  // ================= 1. 표지 =================
  pres.addSection({ title: SEC[0] });
  {
    const s = pres.addSlide({ masterName: "Cover", sectionTitle: SEC[0] });
    s.addText("주제 5 · 소비자의 권리", { placeholder: "eyebrow" });
    s.addText("홍콩 ELS 사태로 알아보는\n소비자 권리", { placeholder: "title" });
    hair(s, L, 3.3, CW, INK, 1);
    s.addText("핵심 질문", { placeholder: "label" });
    s.addText("법이 준 권리는 실제로 작동했는가", { placeholder: "question" });
    motif(s, 7.2, 3.75, false);
    s.addNotes(NOTES[0] + "\n\n[표지] 수치(손실률 44.2% 등)는 표지에서 빼고 2·3번 장에서 설명함. 오른쪽 아래 도식은 발표 전체의 시각 규칙임: 윤곽선 원은 제도·보호 장치, 채운 원은 실제 손실·부담, 점선은 작동하지 않은 연결.");
  }

  // ================= 2. 배경 =================
  pres.addSection({ title: SEC[1] });
  {
    const s = content(SEC[1]);
    head(s, "배경", "법이 생긴 뒤에도 더 큰 손실이 발생함",
      "강의자료 8장 1절 · 금융위원회·금융감독원(2025.2) · 원 넓이는 손실액 비례(팀 작성)");
    const ly = 2.95, cols = [1.9, 4.9, 7.9];
    const dE = 1.9, dD = dE / Math.sqrt(46000 / 4468), dR = 0.76;
    hair(s, cols[0] + dD / 2, ly, cols[1] - dR / 2 - cols[0] - dD / 2, INK, 1.25);
    hair(s, cols[1] + dR / 2 + 0.12, ly, cols[2] - dE / 2 - cols[1] - dR / 2 - 0.3, INK, 1.25, { dashType: "dash" });
    T(s, "닿지 않은 보호", { x: 5.4, y: ly - 0.32, w: 1.5, h: 0.22, fontSize: 9.5, color: MUTED, align: "center" });
    const yrs = ["2019", "2021.3", "2024"], evs = ["DLF 사태", "금융소비자보호법 시행", "홍콩 H지수 ELS 사태"];
    cols.forEach((cx, i) => {
      T(s, yrs[i], { x: cx - 1.3, y: 1.3, w: 2.6, h: 0.36, fontSize: 22, bold: true, align: "center" });
      T(s, evs[i], { x: cx - 1.3, y: 1.68, w: 2.6, h: 0.25, fontSize: 12, align: "center", color: MUTED });
    });
    disc(s, cols[0] - dD / 2, ly - dD / 2, dD);
    T(s, "손실 4,468억원", { x: cols[0] - 1.3, y: 3.45, w: 2.6, h: 0.3, fontSize: 15, bold: true, align: "center" });
    await ringIcon(s, fa.FaShieldAlt, cols[1] - dR / 2, ly - dR / 2, dR);
    T(s, "6대 판매원칙 도입", { x: cols[1] - 1.3, y: 3.45, w: 2.6, h: 0.3, fontSize: 15, align: "center" });
    disc(s, cols[2] - dE / 2, ly - dE / 2, dE);
    T(s, "4.6조원", { x: cols[2] - 0.9, y: ly - 0.45, w: 1.8, h: 0.5, fontSize: 30, bold: true, color: WHITE, align: "center", valign: "middle" });
    T(s, "손실률 44.2%", { x: cols[2] - 0.9, y: ly + 0.1, w: 1.8, h: 0.25, fontSize: 11, color: WHITE, align: "center" });
    hair(s, L, 4.3, CW);
    T(s, [{ text: "손실 약 10배.", options: { bold: true } }, { text: "  규제는 강화됐지만 현장의 소비자 보호는 작동하지 않음", options: {} }],
      { x: L, y: 4.42, w: CW, h: 0.4, fontSize: 13.5, valign: "middle" });
    s.addNotes(NOTES[1] + "\n\n[시각 규칙] 금소법은 윤곽선 원(제도), 두 사태의 손실은 채운 원(현실). 금소법과 ELS 사태 사이의 점선은 법이 생겼지만 보호가 현장에 닿지 않았다는 뜻임.");
  }

  // ================= 3. ELS =================
  {
    const s = content(SEC[1]);
    head(s, "ELS란", "은행 창구에서 팔린 '예금 같은' 투자상품",
      "강의자료 8장 1·2절 · 금융위원회·금융감독원(2025.2), 5대 판매은행 2024년 말 기준");
    const xa = 2.1, xb = 3.95, rh = 0.66, y0 = 1.9;
    T(s, "정기예금", { x: xa, y: 1.4, w: 1.7, h: 0.38, fontSize: 14, bold: true, valign: "middle" });
    T(s, "ELS", { x: xb, y: 1.4, w: 1.7, h: 0.38, fontSize: 14, bold: true, valign: "middle" });
    hair(s, L, y0, 5.1, INK, 1);
    const rows = [["판매 창구", "은행", "은행", false], ["원금", "보장", "손실 가능", true], ["예금자보호", "대상", "대상 아님", true], ["수익", "확정 이자", "지수 연계", true]];
    rows.forEach((r, i) => {
      const y = y0 + i * rh;
      if (i > 0) hair(s, L, y, 5.1);
      T(s, r[0], { x: L, y, w: 1.4, h: rh, fontSize: 12, color: MUTED, valign: "middle" });
      if (r[3]) { ring(s, xa, y + rh / 2 - 0.08, 0.16); disc(s, xb, y + rh / 2 - 0.08, 0.16); }
      const dx = r[3] ? 0.28 : 0;
      T(s, r[1], { x: xa + dx, y, w: 1.6 - dx, h: rh, fontSize: 14, valign: "middle" });
      T(s, r[2], { x: xb + dx, y, w: 1.7 - dx, h: rh, fontSize: 14, bold: r[3], valign: "middle" });
    });
    hair(s, L, y0 + 4 * rh, 5.1);
    vhair(s, 6.05, 1.4, 3.3);
    const rx = 6.4, rw = R - rx;
    T(s, "손실확정 계좌", { x: rx, y: 1.4, w: rw, h: 0.22, fontSize: 10.5, color: MUTED });
    T(s, "17.0만", { x: rx, y: 1.62, w: rw, h: 0.6, fontSize: 34, bold: true, valign: "middle" });
    T(s, "손실액", { x: rx, y: 2.5, w: rw, h: 0.22, fontSize: 10.5, color: MUTED });
    T(s, "4.6조원", { x: rx, y: 2.72, w: rw, h: 0.6, fontSize: 34, bold: true, color: LOSS, valign: "middle" });
    s.addShape(pres.ShapeType.rect, { x: rx, y: 3.75, w: rw, h: 0.36, fill: NOFILL(), line: { color: INK, width: 1.25 }, objectName: "원금 윤곽" });
    s.addShape(pres.ShapeType.rect, { x: rx, y: 3.75, w: rw * 0.442, h: 0.36, fill: { color: LOSS }, line: { type: "none" }, objectName: "손실 면" });
    T(s, "44.2%", { x: rx, y: 3.75, w: rw * 0.442, h: 0.36, fontSize: 11, bold: true, color: WHITE, align: "center", valign: "middle" });
    T(s, "원금 10.4조원", { x: rx, y: 4.2, w: rw, h: 0.22, fontSize: 10.5, color: MUTED, align: "right" });
    s.addNotes(NOTES[2]);
  }

  // ================= 4. 권리 =================
  {
    const s = content(SEC[1]);
    head(s, "금융소비자보호법", "금소법이 소비자에게 준 여섯 가지 장치",
      "강의자료 8장 1절 주요내용 2~4 · 금소법 제28·42·44·46·47·57조");
    const groups = [
      ["권리 강화", [[fa.FaUndoAlt, "청약철회권", "대상 상품은 7일 안에 철회"], [fa.FaFileContract, "위법계약해지권", "위반 계약을 위약금 없이 해지"]]],
      ["분쟁 시 권익 확보", [[fa.FaHandshake, "조정이탈금지", "소액분쟁 조정 중 제소 제한"], [fa.FaFolderOpen, "자료열람요구권", "분쟁 시 자료 열람·사본 요구"], [fa.FaBalanceScale, "입증책임 전환", "금융회사가 무과실을 입증"]]],
      ["처벌 강화", [[fa.FaGavel, "징벌적 과징금", "관련 수입의 50% 이내"]]],
    ];
    const gw = 2.7, gg = 0.35;
    for (let g = 0; g < 3; g++) {
      const x = L + g * (gw + gg);
      T(s, groups[g][0], { x, y: 1.38, w: gw, h: 0.25, fontSize: 11, bold: true, color: MUTED, valign: "middle" });
      hair(s, x, 1.7, gw, INK, 1);
      for (let i = 0; i < groups[g][1].length; i++) {
        const [ic, name, desc] = groups[g][1][i];
        const y = 1.92 + i * 0.98;
        await ringIcon(s, ic, x, y, 0.56);
        T(s, name, { x: x + 0.74, y: y - 0.02, w: gw - 0.74, h: 0.32, fontSize: 15, bold: true, valign: "middle" });
        T(s, desc, { x: x + 0.74, y: y + 0.31, w: gw - 0.74, h: 0.25, fontSize: 10.5, color: MUTED, valign: "middle" });
      }
    }
    s.addNotes(NOTES[3] + "\n\n[시각 규칙] 여섯 장치는 모두 제도이므로 윤곽선 원으로만 그림. 강의자료 주요내용 2(권리 강화)·3(분쟁 시 권익 확보)·4(처벌 강화)의 세 묶음으로 배치함.");
  }

  // ================= 5. 실제 구제 =================
  pres.addSection({ title: SEC[2] });
  {
    const s = content(SEC[2]);
    head(s, "실제 구제", "실제 구제는 자율배상 중심으로 진행됨",
      "금융위원회·금융감독원(2025.2), 5대 판매은행 2024년 말 기준 · 금융위원회 위법계약해지권 안내 · 법제처");
    const part = 1.3 / 4.1, by = 2.0, bh = 0.85;
    T(s, "31.4%", { x: L, y: 1.25, w: 2.8, h: 0.55, fontSize: 32, bold: true, valign: "middle" });
    T(s, "평균 배상비율", { x: L + 1.75, y: 1.42, w: 1.6, h: 0.3, fontSize: 11, color: MUTED, valign: "middle" });
    T(s, "손실 약 4.1조원 · 배상 동의 15.9만 계좌", { x: 5.4, y: 1.42, w: 4.0, h: 0.3, fontSize: 10.5, color: MUTED, align: "right", valign: "middle" });
    s.addShape(pres.ShapeType.rect, { x: L, y: by, w: CW * part, h: bh, fill: NOFILL(), line: { color: INK, width: 1.25 }, objectName: "배상 윤곽" });
    s.addShape(pres.ShapeType.rect, { x: L + CW * part, y: by, w: CW * (1 - part), h: bh, fill: { color: LOSS }, line: { type: "none" }, objectName: "소비자 부담 면" });
    T(s, "배상 1.3조원", { x: L, y: by, w: CW * part, h: bh, fontSize: 14, align: "center", valign: "middle" });
    T(s, "소비자 부담 2.8조원", { x: L + CW * part, y: by, w: CW * (1 - part), h: bh, fontSize: 20, bold: true, color: WHITE, align: "center", valign: "middle" });
    // 경로 A: 권리 → (끊어진 선) → 확정 손실 회복
    const ya = 3.42, yb = 4.22, bhx = 0.44;
    T(s, "법이 준 권리", { x: L, y: ya - 0.24, w: 2, h: 0.2, fontSize: 9.5, color: MUTED });
    outBox(s, "위법계약해지권", L, ya, 2.0, bhx);
    hair(s, L + 2.1, ya + bhx / 2, 3.6, INK, 1.25, { dashType: "dash" });
    T(s, "장래효 · 확정된 손실은 되돌리지 못함", { x: L + 2.1, y: ya - 0.06, w: 3.6, h: 0.22, fontSize: 9.5, color: MUTED, align: "center" });
    T(s, "확정 손실 회복", { x: 6.95, y: ya, w: 2.45, h: bhx, fontSize: 12.5, color: MUTED, valign: "middle" });
    // 경로 B: 분쟁조정기준 → 자율배상 → 배상
    T(s, "실제 구제 경로", { x: L, y: yb - 0.24, w: 2, h: 0.2, fontSize: 9.5, color: MUTED });
    outBox(s, "금감원 분쟁조정기준", L, yb, 2.0, bhx);
    hair(s, L + 2.0, yb + bhx / 2, 0.95, INK, 1.25, { endArrowType: "triangle" });
    outBox(s, "은행 자율배상", L + 3.0, yb, 2.0, bhx);
    hair(s, L + 5.0, yb + bhx / 2, 1.2, INK, 1.25, { endArrowType: "triangle" });
    T(s, "배상 1.3조원", { x: 6.95, y: yb, w: 2.45, h: bhx, fontSize: 12.5, bold: true, valign: "middle" });
    s.addNotes(NOTES[4].replace("\n\n[구성] 위법계약해지권의 한계와 대표사례는 다음 장(6번)으로 옮김.", "")
      + "\n\n[시각 규칙] 긴 막대 하나가 손실 약 4.1조원 전체임. 윤곽선 부분이 배상(1.3조원), 채운 부분이 소비자 부담(2.8조원, 팀 계산 4.1 - 1.3). 아래 두 줄 가운데 점선은 위법계약해지권이 확정 손실 회복으로 이어지지 못했다는 뜻이고, 실선은 실제로 작동한 경로(분쟁조정기준 → 자율배상)임."
      + "\n\n[위법계약해지권] 해지 효과는 장래를 향해 발생함. 만기에 이미 확정된 손실을 돌려받는 장치가 아니며 손해배상청구와 구별됨. 적용 요건: 계속적 거래이면서 해지 시 재산상 불이익이 있는 상품(법 제47조, 시행령 제38조). 'ELS에는 적용되지 않았다'고 단정하지 않고 구조적 한계까지만 말함. 분쟁조정·손해배상도 권리구제 수단이라는 점을 함께 말함.");
  }

  // ================= 6. 대표사례 =================
  {
    const s = content(SEC[2]);
    head(s, "대표사례", "배상비율은 사례에 따라 30~65%",
      "금감원 분쟁조정위원회 대표사례(2024.5.13 의결)");
    const banks = ["농협 65%", "국민 60%", "신한 55%", "SC제일 55%", "하나 30%"], v = [65, 60, 55, 55, 30];
    s.addChart(pres.charts.BAR, [
      { name: "배상", labels: banks, values: v },
      { name: "소비자 부담", labels: banks, values: v.map((x) => 100 - x) },
    ], {
      x: 0.45, y: 1.3, w: 5.6, h: 3.1, barDir: "bar", barGrouping: "stacked", catAxisOrientation: "maxMin", barGapWidthPct: 55,
      chartColors: [H.lt2, H.accent1], dataBorder: { pt: 1, color: H.dk1 },
      catAxisLabelColor: H.dk1, catAxisLabelFontFace: CF, catAxisLabelFontSize: 10.5, catAxisLineShow: false,
      valAxisHidden: true, valAxisMaxVal: 100, valAxisMinVal: 0, valAxisLineShow: false, valGridLine: { style: "none" }, catGridLine: { style: "none" },
      showValue: false, showLegend: false, objectName: "대표사례 배상비율 차트",
    });
    s.addShape(pres.ShapeType.rect, { x: 1.75, y: 4.55, w: 0.2, h: 0.2, fill: NOFILL(), line: { color: INK, width: 1 }, objectName: "범례 배상" });
    T(s, "배상", { x: 2.03, y: 4.5, w: 0.7, h: 0.3, fontSize: 10.5, color: MUTED, valign: "middle" });
    s.addShape(pres.ShapeType.rect, { x: 2.85, y: 4.55, w: 0.2, h: 0.2, fill: { color: LOSS }, line: { type: "none" }, objectName: "범례 부담" });
    T(s, "소비자 부담", { x: 3.13, y: 4.5, w: 1.2, h: 0.3, fontSize: 10.5, color: MUTED, valign: "middle" });
    const rx = 6.5, rw = R - rx;
    hair(s, rx, 1.45, rw, INK, 1);
    T(s, "국민은행 사례", { x: rx, y: 1.6, w: rw, h: 0.24, fontSize: 10.5, color: MUTED, valign: "middle" });
    T(s, "60%", { x: rx, y: 1.9, w: rw, h: 0.8, fontSize: 44, bold: true, valign: "middle" });
    T(s, "예금하러 온 고객에게\nELT 권유", { x: rx, y: 2.85, w: rw, h: 0.62, fontSize: 14.5, bold: true, valign: "top" });
    hair(s, rx, 3.75, rw);
    T(s, "판매사 책임에\n투자자 개별 사정을 반영", { x: rx, y: 3.88, w: rw, h: 0.5, fontSize: 10.5, color: MUTED });
    s.addNotes(NOTES[5].replace(/\n\n\[위법계약해지권\][\s\S]*$/, "")
      + "\n\n[시각 규칙] 5번 장과 같은 규칙임. 윤곽선 부분이 배상비율, 채운 부분이 나머지(소비자 부담)임. 위법계약해지권의 한계는 5번 장으로 옮김.");
  }

  // ================= 7. 제재 경과 =================
  {
    const s = content(SEC[2]);
    head(s, "제재 경과", "첫 검사에서 최종 의결까지 약 2년 10개월",
      "금융위원회·금융감독원(2025.2) · 금융위원회 보도자료(2026.11.4, 제19차 정례회의 의결) · 기간은 팀 계산");
    const xs = [1.3, 3.15, 5.0, 6.85, 8.7], ty = 3.15;
    hair(s, xs[0], 2.0, xs[4] - xs[0], INK, 1);
    vhair(s, xs[0], 1.9, 0.2, INK, 1);
    vhair(s, xs[4], 1.9, 0.2, INK, 1);
    s.addText("약 2년 10개월", { x: 3.9, y: 1.72, w: 2.2, h: 0.56, fill: { color: BG }, line: { type: "none" }, fontSize: 20, bold: true, color: INK,
      align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: "기간" });
    hair(s, xs[0], ty, xs[4] - xs[0], INK, 1.25);
    const tl = [["2024.1~3", "금감원 현장검사", "검사"], ["2024.3.11", "분쟁조정기준", "구제"], ["2024.5.13", "대표사례 의결", "구제"], ["2025.2", "재발방지 대책", "대책"], ["2026.11.4", "과징금 최종 의결", "제재"]];
    tl.forEach((t, i) => {
      const last = i === 4, d = last ? 0.4 : 0.26;
      if (last) disc(s, xs[i] - d / 2, ty - d / 2, d, INK); else ring(s, xs[i] - d / 2, ty - d / 2, d);
      T(s, t[2], { x: xs[i] - 0.6, y: 2.6, w: 1.2, h: 0.22, fontSize: 10, bold: true, color: MUTED, align: "center" });
      T(s, t[0], { x: xs[i] - 0.9, y: 3.55, w: 1.8, h: 0.32, fontSize: 15, bold: true, align: "center" });
      T(s, t[1], { x: xs[i] - 0.9, y: 3.9, w: 1.8, h: 0.25, fontSize: 11, align: "center", color: MUTED });
    });
    s.addNotes(NOTES[6]);
  }

  // ================= 8. 제재 규모 =================
  {
    const s = content(SEC[2]);
    head(s, "제재 규모", "최종 제재는 약 3,746억원",
      "강의자료 8장 1절 주요내용 4 · 금융위원회 보도자료(2026.11.4, 제19차 정례회의 의결)");
    T(s, "은행 7곳 · 증권사 6곳", { x: L, y: 1.45, w: 3.8, h: 0.25, fontSize: 11.5, color: MUTED });
    T(s, "3,746억원", { x: L, y: 1.8, w: 3.9, h: 1.0, fontSize: 50, bold: true, valign: "middle" });
    T(s, "과징금 3,712억 + 과태료 34억", { x: L, y: 2.9, w: 3.8, h: 0.28, fontSize: 12.5, valign: "middle" });
    hair(s, L, 3.5, 3.6);
    T(s, "금융위 최대 규모 금전 제재", { x: L, y: 3.62, w: 3.8, h: 0.28, fontSize: 11, color: MUTED, valign: "middle" });
    vhair(s, 4.6, 1.45, 3.35);
    T(s, "은행별 과징금 (억원)", { x: 4.95, y: 1.4, w: 4.4, h: 0.3, fontSize: 12, bold: true, valign: "middle" });
    s.addChart(pres.charts.BAR, [{ name: "과징금(억원)", labels: ["국민", "하나", "농협", "신한", "SC제일", "우리"], values: [1723.34, 924.75, 401.75, 387.45, 179.0, 3.31] }], {
      x: 4.8, y: 1.7, w: 4.7, h: 3.15, barDir: "bar", catAxisOrientation: "maxMin", barGapWidthPct: 55, chartColors: [H.dk1],
      catAxisLabelColor: H.dk1, catAxisLabelFontFace: CF, catAxisLabelFontSize: 11, catAxisLineShow: false,
      valAxisHidden: true, valAxisMaxVal: 2100, valAxisMinVal: 0, valAxisLineShow: false, valGridLine: { style: "none" }, catGridLine: { style: "none" },
      showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "#,##0", dataLabelColor: H.dk1, dataLabelFontFace: CF, dataLabelFontSize: 11, dataLabelFontBold: true,
      showLegend: false, objectName: "은행별 과징금 차트",
    });
    s.addNotes(NOTES[7]);
  }

  // ================= 9. 해외 =================
  pres.addSection({ title: SEC[3] });
  {
    const s = content(SEC[3]);
    head(s, "해외 사례", "방향은 같지만, 규율하는 지점은 다름",
      "FCA FG22/5 · SEC Regulation Best Interest · 일본증권업협회(2023) · 금융위원회 보도자료(2026.11.4) · 단계 구분은 팀 작성");
    const gx = 3.8, gw = R - gx, st = ["상품 설계", "설명·표시", "권유", "사후 지원"], sw = gw / 4, rh = 0.74, y0 = 1.72;
    st.forEach((t, i) => T(s, t, { x: gx + i * sw, y: 1.32, w: sw, h: 0.3, fontSize: 11, bold: true, align: "center", valign: "middle", color: MUTED }));
    hair(s, L, y0, CW, INK, 1);
    const rowsD = [
      ["UK", "FCA Consumer Duty", "영국 · 2023", 0, 4, "고객에게 좋은 결과", false],
      ["US", "Regulation Best Interest", "미국 · 2020", 2, 1, "최선이익", false],
      ["JP", "복잡한 구조채 표시", "일본 · 2023", 1, 1, "경고 표시", false],
      ["KR", "최선이익 원칙 법제화", "한국 · 국회 계류 중", 0, 4, "아직 시행 전", true],
    ];
    rowsD.forEach((r, i) => {
      const y = y0 + i * rh;
      hair(s, L, y + rh, CW);
      s.addText(r[0], { shape: pres.ShapeType.ellipse, x: L, y: y + 0.15, w: 0.44, h: 0.44, fill: NOFILL(), line: { color: INK, width: 1.25, ...(r[6] ? { dashType: "dash" } : {}) },
        fontSize: 9.5, bold: true, color: INK, align: "center", valign: "middle", margin: 0, objectName: nm("Badge") });
      T(s, r[1], { x: L + 0.6, y: y + 0.12, w: 2.55, h: 0.26, fontSize: 12.5, bold: true, valign: "middle" });
      T(s, r[2], { x: L + 0.6, y: y + 0.4, w: 2.55, h: 0.22, fontSize: 10, color: MUTED });
      s.addText(r[5], { shape: pres.ShapeType.roundRect, rectRadius: 0.17, x: gx + r[3] * sw + 0.08, y: y + 0.2, w: r[4] * sw - 0.16, h: 0.34, fill: NOFILL(),
        line: { color: INK, width: 1.25, ...(r[6] ? { dashType: "dash" } : {}) }, fontSize: 11, bold: !r[6], color: r[6] ? MUTED : INK, align: "center", valign: "middle", margin: 0, objectName: nm("Span") });
    });
    s.addNotes(NOTES[8] + "\n\n[시각 규칙] 네 줄 모두 제도이므로 윤곽선으로 그림. 한국 줄만 점선인 것은 개정안이 국회에 계류 중이라 아직 시행되지 않았다는 뜻임.");
  }

  // ================= 10. 문제점 =================
  {
    const s = content(SEC[3]);
    head(s, "주요 문제점", "문제는 상품보다 판매 구조에 있었음",
      "강의자료 4장 3절, 6장 2절, 8장 1·2절 · 금융위원회·금융감독원(2025.2) · 금감원 현장검사 결과(2024.3.11)");
    const stg = [
      { x: L, w: 2.3, side: "판매자", t: "판매실적 압박", items: [["판매실적 중심의 KPI", "고위험 상품에 유리한 평가", "6장 2절 · 4장 3절"]] },
      { x: 3.4, w: 2.75, side: "판매자", t: "권유·평가 과정", items: [["예금 창구에서 고위험 상품 권유", "만기 고객에게 같은 창구에서", "8장 2절 예금자보호"], ["형식적인 적합성 평가", "\"가입하고 싶어요\" 답변 유도", "8장 1절 6대 판매원칙"]] },
      { x: 6.65, w: 2.75, side: "소비자", t: "소비자 판단", items: [["판매직원에 대한 높은 의존", "직원 설명만 믿고 서명", "8장 1절 소비자의 책무"]] },
    ];
    stg.forEach((g, i) => {
      T(s, g.side, { x: g.x, y: 1.3, w: 2, h: 0.2, fontSize: 9.5, bold: true, color: MUTED });
      disc(s, g.x, 1.6, 0.36, INK);
      T(s, String(i + 1), { x: g.x, y: 1.6, w: 0.36, h: 0.36, fontSize: 12, bold: true, color: WHITE, align: "center", valign: "middle" });
      T(s, g.t, { x: g.x + 0.48, y: 1.6, w: g.w - 0.48, h: 0.36, fontSize: 16, bold: true, valign: "middle" });
      hair(s, g.x, 2.12, g.w, INK, 1);
      if (i < 2) hair(s, g.x + g.w + 0.08, 1.78, stg[i + 1].x - g.x - g.w - 0.16, INK, 1.25, { endArrowType: "triangle" });
      g.items.forEach((it, k) => {
        const y = 2.28 + k * 0.98;
        if (k > 0) hair(s, g.x, y - 0.1, g.w);
        T(s, it[0], { x: g.x, y, w: g.w, h: 0.28, fontSize: 12.5, bold: true, valign: "middle" });
        T(s, it[1], { x: g.x, y: y + 0.29, w: g.w, h: 0.24, fontSize: 10.5, color: MUTED, valign: "middle" });
        T(s, "강의자료 " + it[2], { x: g.x, y: y + 0.54, w: g.w, h: 0.22, fontSize: 9, color: MUTED, valign: "middle" });
      });
    });
    s.addShape(pres.ShapeType.rect, { x: L, y: 4.38, w: CW, h: 0.5, fill: { color: LOSS }, line: { type: "none" }, objectName: "결과 면" });
    T(s, [{ text: "결과", options: { bold: false } }, { text: "    17.0만 계좌 · 손실 4.6조원", options: { bold: true } }],
      { x: L + 0.25, y: 4.38, w: CW - 0.5, h: 0.5, fontSize: 13.5, color: WHITE, valign: "middle" });
    s.addNotes(NOTES[9].replace(/\n\n\[시각화\][^\n]*/, "")
      + "\n\n[구성] 네 문제를 나열하지 않고 연결 순서로 배치함: ① 판매실적 압박(원래 03 KPI) → ② 권유·평가 과정(원래 01 창구 권유, 02 형식적 적합성 평가) → ③ 소비자 판단(원래 04 직원 의존). 각 항목 아래 작은 글씨는 연결되는 강의자료 위치임(KPI는 6장 2절 이해상충, 4장 3절 수익성과 공공성). 맨 아래 채운 띠는 그 결과인 손실 규모(5대 판매은행 손실확정 계좌, 2024년 말).");
  }

  // ================= 11. 개선방안 =================
  {
    const s = content(SEC[3]);
    head(s, "개선방안", "네 주체가 각자 바꿔야 할 것",
      "강의자료 6장 4절, 8장 1절 주요내용 5 · 금융위원회·금융감독원(2025.2) · 금융위원회 보도자료(2026.11.4)");
    const Q = [
      [fa.FaLandmark, "정부·금융당국", ["최선이익 원칙 법제화", "국회 계류"], ["신속하고 예측 가능한 제재", "팀 제안"]],
      [fa.FaUniversity, "금융기관 (판매사)", ["거점점포·전담직원만 판매", "2025.2 대책"], ["고객 손익을 반영한 KPI", "2025.2 대책"]],
      [fa.FaIndustry, "기업 (발행·설계)", ["설명서 맨 위에 손실 사례", ""], ["상품명 앞 '고난도' 표시", ""]],
      [fa.FaUser, "소비자", ["예금자보호 여부부터 확인", ""], ["이해 못 한 계약엔 서명하지 않음", ""]],
    ];
    const y0 = 1.4, rh = 0.84, c1 = 3.5, c2 = 6.55;
    hair(s, L, y0, CW, INK, 1);
    for (let i = 0; i < 4; i++) {
      const y = y0 + i * rh;
      hair(s, L, y + rh, CW);
      await ringIcon(s, Q[i][0], L, y + rh / 2 - 0.24, 0.48);
      T(s, Q[i][1], { x: L + 0.65, y, w: 2.2, h: rh, fontSize: 14, bold: true, valign: "middle" });
      [[c1, Q[i][2]], [c2, Q[i][3]]].forEach(([x, it]) => {
        const has = !!it[1];
        T(s, it[0], { x, y: has ? y + 0.16 : y, w: 2.85, h: has ? 0.3 : rh, fontSize: 12.5, valign: "middle" });
        if (has) T(s, it[1], { x, y: y + 0.47, w: 2.85, h: 0.22, fontSize: 9.5, color: MUTED, valign: "middle" });
      });
    }
    s.addNotes(NOTES[10].replace("항목별 상태를 색 글씨로 구분함.", "항목별 상태를 과제 아래 작은 글씨로 구분함."));
  }

  // ================= 12. 결론 =================
  pres.addSection({ title: SEC[4] });
  {
    const s = pres.addSlide({ masterName: "Closing", sectionTitle: SEC[4] });
    s.addText("결론", { placeholder: "eyebrow" });
    s.addText("권리를 늘리는 것보다,\n있는 권리가 작동하게 하는 것이 중요함", { placeholder: "title" });
    hair(s, L, 2.95, CW, INK, 1);
    T(s, "실행 과제", { x: L, y: 3.12, w: 3, h: 0.25, fontSize: 10.5, bold: true, color: MUTED, charSpacing: 2 });
    const K = [["01", "원칙 중심 규제", "조문 준수에서 소비자 최선이익으로"], ["02", "신속한 제재", "첫 검사부터 의결까지 2년 10개월"]];
    K.forEach((k, i) => {
      const x = L + i * 3.3;
      T(s, k[0], { x, y: 3.55, w: 0.6, h: 0.4, fontSize: 18, bold: true, color: MUTED, valign: "middle" });
      T(s, k[1], { x: x + 0.6, y: 3.55, w: 2.6, h: 0.4, fontSize: 18, bold: true, valign: "middle" });
      T(s, k[2], { x: x + 0.6, y: 3.98, w: 2.6, h: 0.26, fontSize: 10.5, color: MUTED, valign: "middle" });
    });
    motif(s, 7.2, 3.5, true);
    s.addText("출처   박운학, 「금융경영의 이해」 강의자료 4·6·8장(2026) · 금융위원회·금융감독원, 「홍콩 H지수 ELS 현황 및 대책」(2025.2) · 금융위원회, 「홍콩 H지수 ELS 관련 은행·증권사 검사 결과 조치」(2026.11.4) · 금감원 분쟁조정위원회 대표사례(2024.5) · 금융위원회 위법계약해지권 안내 · 법제처 찾기쉬운 생활법령정보 · FCA FG22/5 · SEC Regulation Best Interest · 일본증권업협회(2023)", { placeholder: "source" });
    s.addNotes(NOTES[11] + "\n\n[구성] 결론 한 문장을 크게 두고 실행 과제를 두 가지로 줄임. 01 원칙 중심 규제(9번 장: 최선이익 원칙 법제화), 02 신속한 제재(7번 장: 2년 10개월). 원인(판매 구조)은 10번 장에서 이미 다뤘으므로 말로 한 번 짚고 넘어감. 오른쪽 도식은 표지의 점선이 실선으로 이어진 모습으로, 제도가 현실에 닿아야 한다는 뜻임.");
  }

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
