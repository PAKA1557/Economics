// 홍콩 ELS · 소비자 권리 발표 덱 — 밝은 배경 PowerPoint 판 (pptxgenjs)
// 좌표는 1920x1080 px 기준으로 쓰고 인치로 바꾼다. 애니메이션은 add_anim.py 가 뒤에서 넣는다.
// 도형 이름 "aN ..." = N번째 클릭에 등장, "a0 ..." = 슬라이드 진입 시 자동 등장. 접미 "|wipe" "|wheel" = 효과.
const path = require("path");
const fs = require("fs");
const pptxgen = require("pptxgenjs");

const OUT = process.argv[2] || "deck.pptx";
const COVER = process.argv[3]; // 표지 이미지(선택)
const NOTES = JSON.parse(fs.readFileSync(path.join(__dirname, "notes_light.json"), "utf8"));

const BG = "EFF0F2", INK = "1B1E24", MUTED = "5E6672", HAIR = "C3C8D0", SOFT = "D9DDE3", MID = "B4BAC4", WHITE = "FFFFFF";
const ACC = "EE4E22", ACC_SOFT = "F6B29D", ACC_MID = "F28462";
const FONT = "Noto Sans KR";
const K = 13.333 / 1920; // px → inch
const P = (px) => px * K;
const PT = (px) => px * 0.5; // px → pt

(async () => {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";
  pres.theme = { headFontFace: FONT, bodyFontFace: FONT };
  pres.title = "홍콩 ELS 사태로 알아보는 소비자 권리";

  let n = 0;
  const nm = (tag, label) => `${tag || "s"} ${label || "obj"} ${++n}`;
  const mk = () => { const s = pres.addSlide(); s.background = { color: BG }; return s; };
  const T = (s, text, x, y, w, h, o = {}) => s.addText(text, {
    x: P(x), y: P(y), w: P(w), h: P(h), isTextBox: true, margin: 0, fontFace: FONT, fontSize: PT(o.size || 32), bold: !!o.bold,
    color: o.color || INK, align: o.align || "left", valign: o.valign || "top", charSpacing: o.spacing, lineSpacingMultiple: o.lh,
    objectName: nm(o.a, "text"), fit: "none", wrap: o.wrap === false ? false : true,
  });
  const rect = (s, x, y, w, h, fill, o = {}) => s.addShape(pres.ShapeType.rect, {
    x: P(x), y: P(y), w: P(w), h: P(h), fill: fill ? { color: fill } : { color: BG, transparency: 100 },
    line: o.line ? { color: o.line, width: o.lw || 2.25, dashType: o.dash ? "dash" : "solid" } : { type: "none" }, objectName: nm(o.a, "rect"),
  });
  const circle = (s, cx, cy, d, fill, o = {}) => s.addShape(pres.ShapeType.ellipse, {
    x: P(cx - d / 2), y: P(cy - d / 2), w: P(d), h: P(d), fill: fill ? { color: fill } : { color: BG },
    line: o.line ? { color: o.line, width: o.lw || 2.25, dashType: o.dash ? "dash" : "solid" } : { type: "none" }, objectName: nm(o.a, "circle"),
  });
  const line = (s, x1, y1, x2, y2, color, o = {}) => s.addShape(pres.ShapeType.line, {
    x: P(Math.min(x1, x2)), y: P(Math.min(y1, y2)), w: P(Math.abs(x2 - x1)), h: P(Math.abs(y2 - y1)),
    line: { color, width: o.lw || 2.25, dashType: o.dash ? "dash" : "solid" }, objectName: nm(o.a, "line"),
  });
  // 곡선 띠: 왼쪽 x1 의 [a1,b1] 에서 오른쪽 x2 의 [a2,b2] 로
  const band = (s, x1, a1, b1, x2, a2, b2, fill, o = {}) => {
    const top = Math.min(a1, a2), bot = Math.max(b1, b2), w = x2 - x1, h = bot - top, m = w / 2;
    const q = (x, y) => ({ x: P(x - x1), y: P(y - top) });
    s.addShape(pres.ShapeType.custGeom, {
      x: P(x1), y: P(top), w: P(w), h: P(h), fill: { color: fill }, line: { type: "none" }, objectName: nm(o.a, "band"),
      points: [
        { ...q(x1, a1), moveTo: true },
        { ...q(x2, a2), curve: { type: "cubic", x1: P(m), y1: P(a1 - top), x2: P(m), y2: P(a2 - top) } },
        q(x2, b2),
        { ...q(x1, b1), curve: { type: "cubic", x1: P(m), y1: P(b2 - top), x2: P(m), y2: P(b1 - top) } },
        { close: true },
      ],
    });
  };
  const poly = (s, x, y, w, h, pts, fill, o = {}) => s.addShape(pres.ShapeType.custGeom, {
    x: P(x), y: P(y), w: P(w), h: P(h), fill: fill ? { color: fill } : { color: BG, transparency: 100 },
    line: o.line ? { color: o.line, width: o.lw || 2.25, dashType: o.dash ? "dash" : "solid" } : { type: "none" }, objectName: nm(o.a, "poly"),
    points: [...pts.map((p, i) => ({ x: P(p[0]), y: P(p[1]), ...(i === 0 ? { moveTo: true } : {}) })), { close: true }],
  });
  const title = (s, t) => T(s, t, 128, 92, 1400, 64, { size: 44, bold: true, valign: "middle" });
  const src = (s, t) => T(s, "참고문헌 " + t, 128, 996, 900, 34, { size: 22, color: MUTED, valign: "middle" });
  const chip = (s, t, x, y, w, h, o = {}) => s.addText(t, {
    shape: pres.ShapeType.roundRect, rectRadius: P(h / 2), x: P(x), y: P(y), w: P(w), h: P(h), margin: 0, fontFace: FONT, fontSize: PT(o.size || 30), bold: true,
    color: o.fill ? WHITE : (o.color || INK), fill: o.fill ? { color: o.fill } : { color: BG, transparency: 100 },
    line: o.fill ? { type: "none" } : { color: o.color || INK, width: 2.25, dashType: o.dash ? "dash" : "solid" }, align: "center", valign: "middle", objectName: nm(o.a, "chip"),
  });

  // ============ 1. 표지 ============
  {
    const s = mk();
    if (COVER) s.addImage({ path: COVER, x: P(1040), y: 0, w: P(880), h: P(1080), sizing: { type: "cover", w: P(880), h: P(1080) }, objectName: "표지 이미지" });
    T(s, "홍콩 ELS 사태로\n알아보는 소비자 권리", 128, 300, 880, 300, { size: 104, bold: true, lh: 1.12 });
    line(s, 128, 660, 868, 660, INK, { lw: 3 });
    T(s, "금융경영의 이해 · 주제 5 소비자의 권리", 128, 690, 880, 50, { size: 32, color: MUTED, valign: "middle" });
    s.addNotes(NOTES.cover);
  }

  // ============ 2. 배경 ============
  {
    const s = mk(); title(s, "배경");
    const ly = 560, cx = [330, 900, 1500], dE = 470, dD = dE / Math.sqrt(46000 / 4468), dR = 170;
    const head = (i, yr, ev, a) => { T(s, yr, cx[i] - 260, 196, 520, 76, { size: 64, bold: true, align: "center", a }); T(s, ev, cx[i] - 260, 276, 520, 44, { size: 32, color: MUTED, align: "center", a }); };
    head(0, "2019", "DLF 사태");
    circle(s, cx[0], ly, dD, ACC);
    T(s, "손실 4,468억원", cx[0] - 260, 680, 520, 56, { size: 44, bold: true, align: "center" });
    T(s, "투자자 3,243명", cx[0] - 260, 742, 520, 44, { size: 32, color: MUTED, align: "center" });
    line(s, cx[0] + dD / 2, ly, cx[1] - dR / 2, ly, INK, { lw: 3, a: "a1" });
    head(1, "2021.3", "금융소비자보호법 시행", "a1");
    circle(s, cx[1], ly, dR, null, { line: INK, lw: 3, a: "a1" });
    T(s, "법", cx[1] - 60, ly - 46, 120, 92, { size: 68, bold: true, align: "center", valign: "middle", a: "a1" });
    T(s, "6대 판매원칙", cx[1] - 260, 680, 520, 56, { size: 44, bold: true, align: "center", a: "a1" });
    chip(s, "청약철회권", cx[1] - 250, 752, 240, 60, { a: "a1" });
    chip(s, "위법계약해지권", cx[1] + 10, 752, 260, 60, { a: "a1" });
    line(s, cx[1] + dR / 2 + 14, ly, cx[2] - dE / 2 - 14, ly, INK, { lw: 3, dash: true, a: "a2" });
    head(2, "2024", "홍콩 H지수 ELS 사태", "a2");
    circle(s, cx[2], ly, dE, ACC, { a: "a2" });
    T(s, "4.6조원", cx[2] - 230, ly - 86, 460, 120, { size: 104, bold: true, color: WHITE, align: "center", valign: "middle", a: "a2" });
    T(s, "손실률 44.2%", cx[2] - 230, ly + 40, 460, 50, { size: 36, bold: true, color: WHITE, align: "center", a: "a2" });
    T(s, "17.0만 계좌", cx[2] - 260, 826, 520, 56, { size: 44, bold: true, align: "center", a: "a2" });
    src(s, "[1]"); s.addNotes(NOTES.background);
  }

  // ============ 3. ELS란 ============
  {
    const s = mk(); title(s, "ELS란");
    rect(s, 128, 440, 26, 240, INK);
    T(s, "은행 창구", 128, 360, 420, 64, { size: 48, bold: true, valign: "middle" });
    band(s, 154, 440, 560, 820, 250, 370, SOFT, { a: "a1|wipe" });
    rect(s, 820, 250, 26, 120, null, { line: INK, lw: 3, a: "a1" });
    T(s, "정기예금", 890, 232, 700, 70, { size: 56, bold: true, valign: "middle", a: "a1" });
    chip(s, "원금 보장", 890, 316, 220, 64, { a: "a1" }); chip(s, "예금자보호", 1126, 316, 240, 64, { a: "a1" }); chip(s, "확정 이자", 1382, 316, 220, 64, { a: "a1" });
    band(s, 154, 560, 680, 820, 730, 850, ACC_SOFT, { a: "a2|wipe" });
    rect(s, 820, 730, 26, 120, ACC, { a: "a2" });
    T(s, "ELS (주가연계증권)", 890, 712, 800, 70, { size: 56, bold: true, color: ACC, valign: "middle", a: "a2" });
    chip(s, "손실 가능", 890, 796, 220, 64, { fill: ACC, a: "a2" }); chip(s, "보호 대상 아님", 1126, 796, 290, 64, { fill: ACC, a: "a2" }); chip(s, "지수 연계", 1432, 796, 220, 64, { fill: ACC, a: "a2" });
    src(s, "[1] [9]"); s.addNotes(NOTES.product);
  }

  // ============ 4. 손실 규모 ============
  {
    const s = mk(); title(s, "손실 규모");
    const cx = 440, cy = 560, d = 560;
    const th = 78;
    s.addShape(pres.ShapeType.ellipse, { x: P(cx - d / 2 + th / 2), y: P(cy - d / 2 + th / 2), w: P(d - th), h: P(d - th), fill: { color: BG, transparency: 100 }, line: { color: SOFT, width: PT(th) }, objectName: nm("", "ring") });
    s.addShape(pres.ShapeType.blockArc, { x: P(cx - d / 2), y: P(cy - d / 2), w: P(d), h: P(d), fill: { color: ACC }, line: { type: "none" }, angleRange: [270, 69.12], arcThicknessRatio: (th / d) * 2, objectName: nm("a0|wheel", "arc") });
    T(s, "44.2%", cx - 200, cy - 76, 400, 110, { size: 96, bold: true, align: "center", valign: "middle" });
    T(s, "원금 대비 손실률", cx - 200, cy + 40, 400, 44, { size: 32, color: MUTED, align: "center" });
    const rows = [["손실이 확정된 계좌", "17.0만 계좌", INK], ["그 계좌에 들어간 원금", "10.4조원", INK], ["손실액", "4.6조원", ACC]];
    rows.forEach((r, i) => {
      const y = 250 + i * 210, a = "a" + (i + 1);
      if (i > 0) line(s, 900, y - 22, 1792, y - 22, HAIR, { lw: 1.5, a });
      T(s, r[0], 900, y, 860, 46, { size: 34, color: MUTED, a });
      T(s, r[1], 900, y + 44, 860, 124, { size: 104, bold: true, color: r[2], valign: "middle", a });
    });
    src(s, "[1]"); s.addNotes(NOTES.scale);
  }

  // ============ 5. 금융소비자의 법적권리 ============
  {
    const s = mk(); title(s, "금융소비자의 법적권리");
    const G = [["권리 강화", [["청약철회권", "7일 안에 계약 철회"], ["위법계약해지권", "위반 계약은 위약금 없이 해지"]]],
      ["분쟁 시 권익 확보", [["조정이탈금지", "소액분쟁 조정 중 제소 제한"], ["자료열람요구권", "분쟁 자료 열람·사본 요구"], ["입증책임 전환", "금융회사가 무과실을 입증"]]],
      ["처벌 강화", [["징벌적 과징금", "관련 수입의 50% 이내"]]]];
    G.forEach((g, i) => {
      const x = 128 + i * 572, w = 520;
      line(s, x, 250, x + w, 250, INK, { lw: 3 });
      T(s, g[0], x, 270, w, 46, { size: 32, bold: true, color: MUTED, valign: "middle" });
      g[1].forEach((it, k) => {
        const y = 350 + k * 180;
        T(s, it[0], x, y, w, 70, { size: 52, bold: true, valign: "middle" });
        T(s, it[1], x, y + 72, w, 48, { size: 32, color: MUTED, valign: "middle" });
      });
    });
    // 오른쪽 아래: 과징금 상한
    T(s, "50%", 1272, 560, 520, 220, { size: 200, bold: true, color: ACC, valign: "middle" });
    T(s, "과징금 상한", 1272, 790, 520, 48, { size: 34, color: MUTED, valign: "middle" });
    // 왼쪽 아래: 철회 기간
    T(s, "7일", 128, 740, 520, 150, { size: 128, bold: true, valign: "middle" });
    T(s, "청약철회 기간", 128, 890, 520, 48, { size: 34, color: MUTED, valign: "middle" });
    src(s, "[4] [9]"); s.addNotes(NOTES.rights);
  }

  // ============ 6. 실제 구제 ============
  {
    const s = mk(); title(s, "실제 구제");
    const sc = 36, x1 = 128, x2 = 780, x3 = 1320, nw = 20, top = 250;
    const y1 = top + 20, h1 = 10.4 * sc; // 원금
    const yK = top, hK = 5.8 * sc, yL = top + hK + 44, hL = 4.6 * sc; // 남은 원금 / 손실
    const hB = 1.3 * sc, hC = 2.8 * sc, hE = 0.5 * sc;
    const yB = yL - 70, yC = yB + hB + 40, yE = yC + hC + 40;
    T(s, "원금 10.4조원", x1, top - 46, 500, 56, { size: 40, bold: true, valign: "middle" });
    rect(s, x1, y1, nw, h1, INK);
    band(s, x1 + nw, y1, y1 + hK, x2, yK, yK + hK, SOFT, { a: "a1|wipe" });
    band(s, x1 + nw, y1 + hK, y1 + h1, x2, yL, yL + hL, ACC_SOFT, { a: "a1|wipe" });
    rect(s, x2, yK, nw, hK, MID, { a: "a1" });
    rect(s, x2, yL, nw, hL, ACC, { a: "a1" });
    T(s, "남은 원금 5.8조원", x2 + 44, yK + 20, 480, 56, { size: 40, bold: true, color: MUTED, valign: "middle", a: "a1" });
    T(s, "손실 4.6조원", x2 - 330, yL + hL + 26, 680, 70, { size: 56, bold: true, color: ACC, align: "center", valign: "middle", a: "a1" });
    band(s, x2 + nw, yL, yL + hB, x3, yB, yB + hB, SOFT, { a: "a2|wipe" });
    band(s, x2 + nw, yL + hB, yL + hB + hC, x3, yC, yC + hC, ACC_MID, { a: "a2|wipe" });
    band(s, x2 + nw, yL + hB + hC, yL + hL, x3, yE, yE + hE, SOFT, { a: "a2|wipe" });
    rect(s, x3, yB, nw, hB, null, { line: INK, lw: 3, a: "a2" });
    rect(s, x3, yC, nw, hC, ACC, { a: "a2" });
    rect(s, x3, yE, nw, hE, MID, { a: "a2" });
    T(s, "배상 1.3조원", x3 + 44, yB - 6, 430, 58, { size: 44, bold: true, valign: "middle", a: "a2" });
    T(s, "소비자 부담 2.8조원", x3 + 44, yC + hC / 2 - 40, 440, 80, { size: 46, bold: true, color: ACC, valign: "middle", a: "a2" });
    T(s, "그 밖 0.5조원", x3 + 44, yE - 16, 430, 50, { size: 32, color: MUTED, valign: "middle", a: "a2" });
    // 아래: 핵심 수치 세 개
    const ky = 800;
    line(s, 128, ky - 24, 1792, ky - 24, HAIR, { lw: 1.5, a: "a3" });
    [["31.4%", "평균 배상비율", ACC], ["15.9만", "배상에 동의한 계좌", INK], ["93.8%", "배상 진행 계좌 중 동의율", INK]].forEach((k, i) => {
      const x = 128 + i * 572;
      T(s, k[0], x, ky, 300, 110, { size: 88, bold: true, color: k[2], valign: "middle", a: "a3" });
      T(s, k[1], x + 4, ky + 112, 520, 46, { size: 32, color: MUTED, valign: "middle", a: "a3" });
    });
    src(s, "[1] [5]"); s.addNotes(NOTES.relief);
  }

  // ============ 7. 대표사례 ============
  {
    const s = mk(); title(s, "대표사례");
    rect(s, 128, 220, 620, 700, ACC);
    T(s, "국민은행 사례", 176, 268, 520, 46, { size: 32, bold: true, color: WHITE, valign: "middle" });
    T(s, "예금하러 온\n고객에게\nELS 권유", 176, 340, 540, 260, { size: 64, bold: true, color: WHITE, lh: 1.15 });
    T(s, "배상비율", 176, 690, 520, 46, { size: 32, color: WHITE, valign: "middle" });
    T(s, "60%", 176, 730, 540, 170, { size: 150, bold: true, color: WHITE, valign: "middle" });
    T(s, "대표사례 5건의 배상비율", 860, 226, 900, 48, { size: 34, bold: true, color: MUTED, valign: "middle", a: "a1" });
    [["농협", 65], ["국민", 60], ["신한", 55], ["SC제일", 55], ["하나", 30]].forEach((r, i) => {
      const y = 320 + i * 112, bx = 1080, bw = 500, bh = 44;
      T(s, r[0], 860, y - 12, 210, 68, { size: 46, bold: true, valign: "middle", a: "a1" });
      rect(s, bx + bw * r[1] / 100, y, bw * (1 - r[1] / 100), bh, ACC, { a: "a1" });
      rect(s, bx, y, bw * r[1] / 100, bh, null, { line: INK, lw: 3, a: "a1" });
      T(s, r[1] + "%", bx + bw + 28, y - 14, 180, 72, { size: 54, bold: true, valign: "middle", a: "a1" });
    });
    rect(s, 860, 884, 30, 30, null, { line: INK, lw: 2.25, a: "a1" }); T(s, "배상", 904, 876, 120, 46, { size: 32, color: MUTED, valign: "middle", a: "a1" });
    rect(s, 1040, 884, 30, 30, ACC, { a: "a1" }); T(s, "소비자 부담", 1084, 876, 300, 46, { size: 32, color: MUTED, valign: "middle", a: "a1" });
    src(s, "[3]"); s.addNotes(NOTES.cases);
  }

  // ============ 8. 제재 경과 ============
  {
    const s = mk(); title(s, "제재 경과");
    T(s, "2년 10개월", 128, 240, 1300, 210, { size: 176, bold: true, valign: "middle" });
    const xs = [228, 594, 960, 1326, 1692], ty = 700;
    // 기간 표시
    line(s, xs[0], 560, xs[4], 560, ACC, { lw: 3 }); line(s, xs[0], 540, xs[0], 580, ACC, { lw: 3 }); line(s, xs[4], 540, xs[4], 580, ACC, { lw: 3 });
    line(s, xs[0], ty, xs[4], ty, MID, { lw: 3 });
    const tl = [["2024.1", "금감원 현장검사", "a1"], ["2024.3", "분쟁조정기준", "a2"], ["2024.5", "대표사례 의결", "a2"], ["2025.2", "재발방지 대책", "a3"], ["2026.11.4", "제재 최종 의결", "a4"]];
    tl.forEach((t, i) => {
      const last = i === 4, al = i === 0 ? "left" : last ? "right" : "center", lx = i === 0 ? xs[i] - 30 : last ? xs[i] - 370 : xs[i] - 200;
      if (last) circle(s, xs[i], ty, 70, ACC, { a: t[2] }); else circle(s, xs[i], ty, 46, null, { line: INK, lw: 3, a: t[2] });
      T(s, t[0], lx, 764, 400, 64, { size: 48, bold: true, color: last ? ACC : INK, align: al, valign: "middle", a: t[2] });
      T(s, t[1], lx, 832, 400, 46, { size: 32, color: MUTED, align: al, valign: "middle", a: t[2] });
    });
    src(s, "[1] [2] [3]"); s.addNotes(NOTES.timeline);
  }

  // ============ 9. 제재 규모 ============
  {
    const s = mk(); title(s, "제재 규모");
    const x1 = 128, x2 = 1060, nw = 20, tot = 3746.32, H = 384, sc = H / tot, top = 340;
    const items = [["국민", 1723.34, 64], ["하나", 924.75, 50], ["농협", 401.75, 38], ["신한", 387.45, 38], ["SC제일", 179.0, 34], ["그 외", 130.03, 32]];
    T(s, "3,746억원", x1, 196, 800, 120, { size: 104, bold: true, valign: "middle" });
    rect(s, x1, top, nw, H, INK);
    let yl = top, yr = top;
    const fills = ["9AA1AD", "AEB4BE", "C0C5CD", "CDD1D8", "D9DDE3", "E2E5E9"];
    items.forEach((it, i) => {
      const h = it[1] * sc;
      band(s, x1 + nw, yl, yl + h, x2, yr, yr + h, fills[i], { a: "a1|wipe" });
      rect(s, x2, yr, nw, h, i < 5 ? INK : MID, { a: "a1" });
      const lh = Math.max(h, it[2] * 1.3);
      T(s, `${it[0]} ${Math.round(it[1]).toLocaleString("en-US")}억`, x2 + 44, yr + h / 2 - lh / 2, 520, lh, { size: it[2], bold: i < 5, color: i < 5 ? INK : MUTED, valign: "middle", a: "a1" });
      yl += h; yr += h + 20;
    });
    chip(s, "금융위 최대 규모 금전 제재", x1, 800, 520, 72, { fill: ACC, size: 32, a: "a2" });
    T(s, "은행 7곳 · 증권사 6곳", x1, 886, 520, 50, { size: 32, color: MUTED, valign: "middle", a: "a2" });
    src(s, "[2]"); s.addNotes(NOTES.sanction);
  }

  // ============ 10. 해외 사례 ============
  {
    const s = mk(); title(s, "해외 사례");
    const gx = 792, sw = 250, st = ["상품 설계", "설명·표시", "권유", "사후 지원"], y0 = 290, rh = 150;
    st.forEach((t, i) => T(s, t, gx + i * sw, 226, sw, 48, { size: 32, bold: true, color: MUTED, align: "center", valign: "middle" }));
    line(s, 128, y0, 1792, y0, INK, { lw: 3 });
    const R = [["영국 Consumer Duty", "2023년 적용", 0, 4, "고객에게 좋은 결과", false], ["미국 Regulation BI", "2020년 적용", 2, 1, "최선이익", false],
      ["일본 복잡한 구조채 표시", "2023년 자율규제", 1, 1, "경고 표시", false], ["한국 최선이익 원칙", "국회 계류 중", 0, 4, "아직 시행 전", true]];
    R.forEach((r, i) => {
      const y = y0 + i * rh, col = r[5] ? ACC : INK;
      line(s, 128, y + rh, 1792, y + rh, HAIR, { lw: 1.5 });
      T(s, r[0], 128, y + 22, 650, 62, { size: 46, bold: true, color: col, valign: "middle" });
      T(s, r[1], 128, y + 84, 650, 44, { size: 32, color: r[5] ? ACC : MUTED, valign: "middle" });
      chip(s, r[4], gx + r[2] * sw + 14, y + 40, r[3] * sw - 28, 70, { size: 34, color: col, dash: r[5] });
    });
    src(s, "[1] [2] [6] [7] [8]"); s.addNotes(NOTES.abroad);
  }

  // ============ 11. 주요 문제점 ============
  {
    const s = mk(); title(s, "주요 문제점");
    const y = 250, h = 640, w = 416, tip = 46;
    const C = [["1 · 판매자", "판매실적 압박", "고위험 상품에\n유리한 KPI", null], ["2 · 판매자", "권유와 평가", "창구에서 ELS 권유\n형식적 적합성 평가", "a1"], ["3 · 소비자", "소비자 판단", "직원 설명만\n믿고 서명", "a2"]];
    C.forEach((c, i) => {
      const x = 128 + i * w, pts = i === 0 ? [[0, 0], [w - tip - 6, 0], [w - 6, h / 2], [w - tip - 6, h], [0, h]] : [[0, 0], [w - tip - 6, 0], [w - 6, h / 2], [w - tip - 6, h], [0, h], [tip, h / 2]];
      poly(s, x, y, w - 6, h, pts, WHITE, { line: INK, lw: 3, dash: i === 2, a: c[3] });
      const tx = x + (i === 0 ? 44 : 84);
      T(s, c[0], tx, y + 70, 290, 46, { size: 32, bold: true, color: MUTED, valign: "middle", a: c[3] });
      T(s, c[1], tx, y + 128, 300, 66, { size: 46, bold: true, valign: "middle", a: c[3] });
      T(s, c[2], tx, y + 250, 290, 110, { size: 34, color: INK, lh: 1.25, a: c[3] });
    });
    const x = 128 + 3 * w;
    poly(s, x, y, w, h, [[0, 0], [w, 0], [w, h], [0, h], [tip, h / 2]], ACC, { a: "a3" });
    T(s, "결과", x + 84, y + 70, 290, 46, { size: 32, bold: true, color: WHITE, valign: "middle", a: "a3" });
    T(s, "4.6조원", x + 84, y + 124, 320, 100, { size: 76, bold: true, color: WHITE, valign: "middle", a: "a3" });
    T(s, "17.0만 계좌의\n손실", x + 84, y + 250, 300, 110, { size: 34, bold: true, color: WHITE, lh: 1.25, a: "a3" });
    src(s, "[1] [9] [10]"); s.addNotes(NOTES.structure);
  }

  // ============ 12. 개선방안 ============
  {
    const s = mk(); title(s, "개선방안");
    const y0 = 230, rh = 172;
    line(s, 128, y0, 1792, y0, INK, { lw: 3 });
    [["정부·금융당국", "최선이익 원칙 법제화", "신속하고 예측 가능한 제재"], ["금융기관", "거점점포·전담직원만 판매", "고객 손익을 반영한 KPI"],
      ["상품 설계 기업", "설명서 맨 위에 손실 사례", "상품명 앞 ‘고난도’ 표시"], ["소비자", "예금자보호 여부부터 확인", "이해 못 한 계약엔 서명 안 함"]].forEach((r, i) => {
      const y = y0 + i * rh;
      line(s, 128, y + rh, 1792, y + rh, HAIR, { lw: 1.5 });
      T(s, r[0], 128, y, 440, rh, { size: 48, bold: true, valign: "middle" });
      circle(s, 606, y + rh / 2, 16, ACC); T(s, r[1], 632, y, 560, rh, { size: 38, valign: "middle" });
      circle(s, 1206, y + rh / 2, 16, ACC); T(s, r[2], 1232, y, 560, rh, { size: 38, valign: "middle" });
    });
    src(s, "[1] [2] [9]"); s.addNotes(NOTES.reform);
  }

  // ============ 13. 결론 ============
  {
    const s = mk(); title(s, "결론");
    [["01", "원칙 중심 규제", "조문 준수에서 소비자 최선이익으로"], ["02", "신속한 제재", "제재 확정까지 2년 10개월"]].forEach((k, i) => {
      const y = 320 + i * 290;
      T(s, k[0], 128, y, 200, 150, { size: 120, bold: true, color: ACC, valign: "middle" });
      T(s, k[1], 340, y, 800, 96, { size: 80, bold: true, valign: "middle" });
      T(s, k[2], 344, y + 100, 800, 50, { size: 36, color: MUTED, valign: "middle" });
    });
    // 오른쪽: 제도와 현실이 이어진 모습
    const cy = 520;
    circle(s, 1330, cy, 190, null, { line: INK, lw: 3 });
    line(s, 1425, cy, 1545, cy, INK, { lw: 3 });
    circle(s, 1640, cy, 190, INK);
    T(s, "제도", 1235, cy - 30, 190, 60, { size: 40, bold: true, align: "center", valign: "middle" });
    T(s, "현실", 1545, cy - 30, 190, 60, { size: 40, bold: true, color: WHITE, align: "center", valign: "middle" });
    T(s, "있는 권리가 작동하도록", 1180, 680, 612, 50, { size: 36, color: MUTED, align: "center", valign: "middle" });
    s.addNotes(NOTES.conclusion);
  }

  // ============ 14. 참고문헌 ============
  {
    const s = mk(); title(s, "참고문헌");
    const L = ["[1] 금융위원회·금융감독원, 「홍콩 H지수 ELS 현황 및 대책」, 2025.2.", "[2] 금융위원회, 「홍콩 H지수 ELS 관련 은행·증권사 검사 결과 조치」(제19차 정례회의 의결), 2026.11.4. fsc.go.kr/no010101/88082",
      "[3] 금융감독원 분쟁조정위원회, 홍콩 H지수 ELS 대표사례 분쟁조정 결과(보도자료), 2024.5.14.", "[4] 「금융소비자 보호에 관한 법률」 및 같은 법 시행령, 국가법령정보센터.", "[5] 금융위원회, 위법계약해지권 안내. fsc.go.kr/po010105/75631"];
    const Rr = ["[6] Financial Conduct Authority, FG22/5 Final non-Handbook Guidance for firms on the Consumer Duty, 2022.", "[7] U.S. Securities and Exchange Commission, Regulation Best Interest: The Broker-Dealer Standard of Conduct, 2019.",
      "[8] 일본증권업협회(JSDA), 복잡한 구조채 판매에 관한 자율규제 개정, 2023.", "[9] 박운학, 「금융경영의 이해」 강의자료 4·6·8장, 2026.", "[10] 금융감독원, 홍콩 H지수 ELS 현장검사 결과, 2024.3.11."];
    const col = (arr, x) => s.addText(arr.map((t, i) => ({ text: t, options: { breakLine: i < arr.length - 1, paraSpaceAfter: 14 } })), { x: P(x), y: P(230), w: P(790), h: P(520), isTextBox: true, margin: 0, fontFace: FONT, fontSize: 14, color: INK, valign: "top", objectName: nm("", "refs") });
    col(L, 128); col(Rr, 1002);
    line(s, 128, 800, 1792, 800, HAIR, { lw: 1.5 });
    T(s, "수치는 5대 판매은행의 2024년 말 기준이며, 5.8조원·2.8조원·0.5조원·2년 10개월·130억원은 위 자료의 수치로 팀이 계산한 값임. 표지 사진은 AI로 생성한 이미지이며 실제 장소의 사진이 아님.", 128, 824, 1664, 90, { size: 25, color: MUTED, lh: 1.3 });
    s.addNotes(NOTES.references);
  }

  await pres.writeFile({ fileName: OUT });
  console.log("wrote", OUT);
})();
