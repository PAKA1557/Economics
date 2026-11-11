"""사용자 PowerPoint 복사본에 요청된 수정만 적용한다 (원본은 건드리지 않음).
사용: python3 revise.py <원본.pptx> <새 파일.pptx>"""
import copy, re, sys
from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Emu, Pt

SRC, DST = sys.argv[1], sys.argv[2]
U = 6350  # 1px(1920 기준) = 6350 EMU
E = lambda px: Emu(int(round(px * U)))
INK, MUTED, HAIR, SOFT, MID, ACC, WHITE, BG = "1B1E24", "5E6672", "C3C8D0", "D9DDE3", "B4BAC4", "EE4E22", "FFFFFF", "EFF0F2"
FONT = "Noto Sans KR"
NS = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main", "a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
P = "{%s}" % NS["p"]

prs = Presentation(SRC)
S = lambda n: prs.slides[n - 1]
rgb = lambda h: RGBColor.from_string(h)


def find(slide, text):
    return [s for s in slide.shapes if s.has_text_frame and s.text_frame.text.strip().startswith(text)]


def remove(shape):
    shape._element.getparent().remove(shape._element)


def text(slide, s, x, y, w, h, size, color=INK, bold=False, align="l", anchor="ctr", name="text"):
    tb = slide.shapes.add_textbox(E(x), E(y), E(w), E(h)); tb.name = name
    tf = tb.text_frame; tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = {"ctr": MSO_ANCHOR.MIDDLE, "t": MSO_ANCHOR.TOP}[anchor]
    lines = s.split("\n")
    for i, ln in enumerate(lines):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}[align]
        r = para.add_run(); r.text = ln
        f = r.font; f.size = Pt(size / 2); f.bold = bold; f.color.rgb = rgb(color); f.name = FONT
        rPr = r._r.get_or_add_rPr()
        for tag in ("ea", "cs"):
            el = etree.SubElement(rPr, "{%s}%s" % (NS["a"], tag)); el.set("typeface", FONT)
    return tb


def shape(slide, kind, x, y, w, h, fill=None, line=None, lw=2.25, dash=False, name="shape"):
    sp = slide.shapes.add_shape(kind, E(x), E(y), E(w), E(h)); sp.name = name
    sp.shadow.inherit = False
    if fill: sp.fill.solid(); sp.fill.fore_color.rgb = rgb(fill)
    else: sp.fill.background()
    if line:
        sp.line.color.rgb = rgb(line); sp.line.width = Pt(lw)
        if dash:
            ln = sp.line._get_or_add_ln(); d = etree.SubElement(ln, "{%s}prstDash" % NS["a"]); d.set("val", "dash")
    else: sp.line.fill.background()
    return sp


def hline(slide, x1, x2, y, color, lw=1.5, dash=False, name="line"):
    c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, E(x1), E(y), E(x2), E(y)); c.name = name
    c.line.color.rgb = rgb(color); c.line.width = Pt(lw)
    if dash:
        ln = c.line._get_or_add_ln(); d = etree.SubElement(ln, "{%s}prstDash" % NS["a"]); d.set("val", "dash")
    return c


# ---------- 애니메이션 XML ----------
EFF = {"fade": ("10", "0", "fade", 450), "wipe": ("22", "8", "wipe(left)", 800), "wheel": ("21", "1", "wheel(1)", 1500)}


def build_timing(slide, steps, extra_after=None):
    """steps: [(mode, [(shape, effect), ...]), ...]; mode = 'click' | 'auto'. extra_after: {step_index: xml 조각(afterEffect par)}"""
    cid = [2]
    def nid():
        cid[0] += 1; return cid[0]
    pars, bld = [], []
    for si, (mode, targets) in enumerate(steps):
        effs = []
        for i, (sh, eff) in enumerate(targets):
            spid = sh.shape_id; pid, sub, flt, dur = EFF[eff]
            node = ("afterEffect" if mode == "auto" else "clickEffect") if i == 0 else "withEffect"
            decel = ' decel="100000"' if eff == "wheel" else ""
            effs.append(f'<p:par><p:cTn id="{nid()}" presetID="{pid}" presetClass="entr" presetSubtype="{sub}" fill="hold" grpId="0" nodeType="{node}"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
                        f'<p:set><p:cBhvr><p:cTn id="{nid()}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>'
                        f'<p:animEffect transition="in" filter="{flt}"><p:cBhvr><p:cTn id="{nid()}" dur="{dur}"{decel}/><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl></p:cBhvr></p:animEffect>'
                        f'</p:childTnLst></p:cTn></p:par>')
            if sh._element.tag == P + "sp": bld.append(f'<p:bldP spid="{spid}" grpId="0" animBg="1"/>')
        inner = f'<p:par><p:cTn id="{nid()}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>{"".join(effs)}</p:childTnLst></p:cTn></p:par>'
        if extra_after and si in extra_after:
            for delay, spid, by, dur in extra_after[si]:  # 강조: 커졌다 돌아오기
                inner += (f'<p:par><p:cTn id="{nid()}" fill="hold"><p:stCondLst><p:cond delay="{delay}"/></p:stCondLst><p:childTnLst>'
                          f'<p:par><p:cTn id="{nid()}" presetID="6" presetClass="emph" presetSubtype="0" autoRev="1" fill="hold" grpId="1" nodeType="afterEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
                          f'<p:animScale><p:cBhvr><p:cTn id="{nid()}" dur="{dur}" decel="100000" fill="hold"/><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl></p:cBhvr><p:by x="{by}" y="{by}"/></p:animScale>'
                          f'</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>')
                bld.append(f'<p:bldP spid="{spid}" grpId="1" animBg="1"/>')
        cond = '<p:cond delay="indefinite"/>' + ('<p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond>' if mode == "auto" else "")
        pars.append(f'<p:par><p:cTn id="{nid()}" fill="hold"><p:stCondLst>{cond}</p:stCondLst><p:childTnLst>{inner}</p:childTnLst></p:cTn></p:par>')
    xml = ('<p:timing xmlns:p="%s"><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>'
           '<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>' % NS["p"] + "".join(pars) +
           '</p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
           '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>'
           '</p:childTnLst></p:cTn></p:par></p:tnLst><p:bldLst>' + "".join(bld) + '</p:bldLst></p:timing>')
    sld = slide._element
    for old in sld.findall(P + "timing"): sld.remove(old)
    new = etree.fromstring(xml)
    ext = sld.find(P + "extLst")
    if ext is not None: ext.addprevious(new)
    else: sld.append(new)


def group(slide, shapes, name):
    g = slide.shapes.add_group_shape(shapes); g.name = name; return g


def note(slide, s):
    tf = slide.notes_slide.notes_text_frame
    tf.text = tf.text.rstrip() + "\n\n" + s


# ============ 참고문헌 줄: 2~12장 오른쪽 아래, 오른쪽 정렬 (높이 유지) ============
for n in range(2, 13):
    for sh in find(S(n), "참고문헌"):
        sh.left = E(1792) - sh.width
        for para in sh.text_frame.paragraphs: para.alignment = PP_ALIGN.RIGHT

# ============ 4장: 도넛 애니메이션 (감속 + 끝에서 튕김) ============
s = S(4)
arc = [x for x in s.shapes if x.name.startswith("arc")][0]
T4 = {x.name: x for x in s.shapes}
rows = [[T4["text 47"], T4["text 48"]], [T4["line 49"], T4["text 50"], T4["text 51"]], [T4["line 52"], T4["text 53"], T4["text 54"]]]
build_timing(s, [("auto", [(arc, "wheel")])] + [("click", [(x, "fade") for x in r]) for r in rows],
             extra_after={0: [(1500, arc.shape_id, 106000, 170), (1850, arc.shape_id, 102500, 130)]})
note(s, "[변경] 도넛의 주황색 부분은 처음에 빠르게, 끝으로 갈수록 느리게 채워지고, 다 채워진 뒤 주황색 부분만 두 번 살짝 커졌다 돌아옴(튕김). 회색 고리와 숫자는 움직이지 않음.")

# ============ 6장: 격자(와플) 차트 ============
s = S(6)
keep = {"s text 80", "line 98", "text 99", "text 100", "text 101", "text 102", "text 103", "text 104", "s text 105"}
for sh in list(s.shapes):
    if sh.name not in keep: remove(sh)
K6 = {x.name: x for x in s.shapes}
gx, gy, cell, pitch = 128, 304, 56, 66.5  # 25열 x 4행 = 100칸, 한 칸 약 0.1조원
pos = lambda i: (gx + (i // 4) * pitch, gy + (i % 4) * pitch)
text(s, "원금 10.4조원", 128, 196, 420, 60, 44, bold=True, name="원금")
text(s, "한 칸 = 약 0.1조원", 560, 204, 500, 46, 30, color=MUTED, name="단위")
for i in range(100):
    x, y = pos(i); shape(s, MSO_SHAPE.RECTANGLE, x, y, cell, cell, fill=SOFT, name=f"칸 {i + 1}")
loss = [shape(s, MSO_SHAPE.RECTANGLE, *pos(i), cell, cell, fill="F6B29D", name=f"손실 칸 {i + 1}") for i in range(56, 100)]
g_loss = group(s, loss, "손실 44칸")
lx0 = gx + 14 * pitch
t_loss = text(s, "손실 4.6조원 · 44.2%", lx0, 226, 700, 52, 38, color=ACC, bold=True, name="손실")
br = hline(s, lx0, gx + 24 * pitch + cell, 288, ACC, lw=3, name="손실 범위")
split = []
for i in range(56, 100):
    x, y = pos(i)
    if i < 69: split.append(shape(s, MSO_SHAPE.RECTANGLE, x + 2, y + 2, cell - 4, cell - 4, fill=BG, line=INK, lw=2.25, name="배상 칸"))
    elif i < 96: split.append(shape(s, MSO_SHAPE.RECTANGLE, x, y, cell, cell, fill=ACC, name="부담 칸"))
    else: split.append(shape(s, MSO_SHAPE.RECTANGLE, x, y, cell, cell, fill=MID, name="그 밖 칸"))
g_split = group(s, split, "손실의 구성")
# 범례 겸 수치 줄
leg = []
for i, (lab, val, fill, line, col) in enumerate((("남은 원금", "5.8조원", SOFT, None, INK), ("배상", "1.3조원", BG, INK, INK), ("소비자 부담", "2.8조원", ACC, None, ACC), ("그 밖", "0.5조원", MID, None, INK))):
    x = 128 + i * 416; y = 606
    sw = shape(s, MSO_SHAPE.RECTANGLE, x, y + 8, 44, 44, fill=fill, line=line, name=f"범례 {lab}")
    a = text(s, lab, x + 62, y, 340, 46, 32, color=MUTED, name=f"범례 이름 {lab}")
    b_ = text(s, val, x + 62, y + 46, 340, 76, 60, color=col, bold=True, name=f"범례 값 {lab}")
    leg.append((sw, a, b_))
kp = [K6[k] for k in ("line 98", "text 99", "text 100", "text 101", "text 102", "text 103", "text 104")]
build_timing(s, [("click", [(g_loss, "wipe"), (t_loss, "fade"), (br, "wipe")]), ("click", [(g_split, "fade")] + [(z, "fade") for tri in leg[1:] for z in tri]), ("click", [(x, "fade") for x in kp])])
note(s, "[변경] 띠 차트를 격자 차트로 바꿈. 네모 100칸이 원금 10.4조원이고 한 칸이 약 0.1조원. 회색 56칸 = 남은 원금 5.8조원, 윤곽선 13칸 = 배상 1.3조원, 주황 27칸 = 소비자 부담 2.8조원, 진회색 4칸 = 그 밖 0.5조원. 칸 수는 반올림한 값이라 실제 비율(12.5·26.9·4.8칸)과 1칸 안팎의 차이가 있음. [화면 순서] 클릭 1: 손실 44칸이 물듦 → 클릭 2: 배상·소비자 부담·그 밖으로 갈림 → 클릭 3: 아래 수치 세 개.")

# ============ 9장: 넓이 비례 원 ============
s = S(9)
keep = {"s text 160", "s text 161", "chip 181", "text 182", "s text 183"}
for sh in list(s.shapes):
    if sh.name not in keep: remove(sh)
K9 = {x.name: x for x in s.shapes}
items = [("국민", 1723.34, ACC), ("하나", 924.75, INK), ("농협", 401.75, "3A404B"), ("신한", 387.45, "59606C"), ("SC제일", 179.0, "7A828F"), ("그 외", 130.03, MID)]
k = 420 / items[0][1] ** 0.5
ds = [k * v ** 0.5 for _, v, _ in items]
gap = (1664 - sum(ds)) / (len(ds) - 1); cy = 560; x = 128
circ = []
for (nm_, v, col), d in zip(items, ds):
    c = shape(s, MSO_SHAPE.OVAL, x, cy - d / 2, d, d, fill=col, name=f"{nm_} 원")
    amt = f"{round(v):,}억"
    if d > 300:
        a = text(s, nm_, x, cy - 74, d, 60, 44, color=WHITE, bold=True, align="c", name=f"{nm_} 이름")
        b = text(s, amt, x, cy - 16, d, 96, 72 if d > 400 else 60, color=WHITE, bold=True, align="c", name=f"{nm_} 금액")
    else:
        w = max(d, 240); lx = x + d / 2 - w / 2
        a = text(s, nm_, lx, cy + d / 2 + 14, w, 48, 34, bold=True, align="c", name=f"{nm_} 이름")
        b = text(s, amt, lx, cy + d / 2 + 60, w, 48, 34, color=MUTED, align="c", name=f"{nm_} 금액")
    circ.append((c, a, b)); x += d + gap
steps = [("click", [(c, "fade"), (a, "fade"), (b, "fade")]) for c, a, b in circ[:2]]
steps.append(("click", [(z, "fade") for tri in circ[2:] for z in tri]))
steps.append(("click", [(K9["chip 181"], "fade"), (K9["text 182"], "fade")]))
build_timing(s, steps)
note(s, "[변경] 띠 차트를 넓이가 과징금에 비례하는 원으로 바꿈. 국민 1,723억 / 하나 925억 / 농협 402억 / 신한 387억 / SC제일 179억 / 그 외 130억(우리은행 과징금, 증권사 과징금·과태료, 은행 과태료를 합친 팀 계산값). [화면 순서] 클릭 1: 국민 → 클릭 2: 하나 → 클릭 3: 나머지 → 클릭 4: ‘금융위 최대 규모 금전 제재’.")

# ============ 12·13장 제목 ============
ref_rPr = find(S(11), "주요 문제점")[0].text_frame.paragraphs[0].runs[0]._r.rPr
for n, old, new in ((12, "Strategic", "개선방안"), (13, "Conclusion", "결론")):
    t = find(S(n), old)[0]; r = t.text_frame.paragraphs[0].runs[0]
    r.text = new
    r._r.remove(r._r.rPr); r._r.insert(0, copy.deepcopy(ref_rPr))

# ============ 13장: 화면 넓히기 + 핵심 수치 + 선 애니메이션 ============
s = S(13)
K = {x.name: x for x in s.shapes}
def mv(name, x=None, y=None, w=None, h=None):
    sh = K[name]
    if x is not None: sh.left = E(x)
    if y is not None: sh.top = E(y)
    if w is not None: sh.width = E(w)
    if h is not None: sh.height = E(h)
mv("s text 253", y=230); mv("s text 254", y=230); mv("s text 255", y=332)
mv("s text 256", y=470); mv("s text 257", y=470); mv("s text 258", y=572)
D = 240; cyy = 400
mv("s circle 259", x=1190, y=cyy - D / 2, w=D, h=D); mv("s text 262", x=1190, y=cyy - 32, w=D, h=64)
mv("s circle 261", x=1552, y=cyy - D / 2, w=D, h=D); mv("s text 263", x=1552, y=cyy - 32, w=D, h=64)
remove(K["s line 260"])
dashed = hline(s, 1430, 1552, cyy, INK, lw=3, dash=True, name="점선")
solid = hline(s, 1430, 1552, cyy, INK, lw=4.5, name="실선")
mv("s text 264", x=1150, y=560, w=682)
hline(s, 128, 1792, 716, HAIR, lw=1.5, name="구분선")
for i, (num, lab, col) in enumerate((("44.2%", "원금 대비 손실률", ACC), ("31.4%", "평균 배상비율", INK), ("2년 10개월", "제재 확정까지", INK))):
    x = 128 + i * 572
    text(s, num, x, 744, 560, 120, 92, color=col, bold=True, name=f"수치 {i + 1}")
    text(s, lab, x + 4, 868, 520, 48, 32, color=MUTED, name=f"수치 설명 {i + 1}")
build_timing(s, [("click", [(solid, "wipe")])])
note(s, "[변경] 과제 두 줄과 오른쪽 그림을 키우고, 아래에 발표 전체를 되짚는 수치 세 개를 넣음: 원금 대비 손실률 44.2%, 평균 배상비율 31.4%(배상 동의 15.9만 계좌 기준), 제재 확정까지 2년 10개월(팀 계산). 오른쪽 그림은 처음에 점선으로 보이다가 클릭하면 실선으로 이어짐.")

prs.save(DST)
import zipfile, shutil, os
tmp = DST + '.tmp'; zin = zipfile.ZipFile(DST); zout = zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED)
for it in zin.infolist():
    d = zin.read(it.filename)
    if it.filename == '[Content_Types].xml' and b'Extension="jpg"' not in d:
        d = d.replace(b'<Default ', b'<Default Extension="jpg" ContentType="image/jpeg"/><Default ', 1)
    zout.writestr(it, d)
zout.close(); zin.close(); os.replace(tmp, DST)
print("saved", DST)
