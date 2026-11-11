"""3차 수정: '수정2' 복사본에 지시된 장만 고친다. 사용: python3 revise3.py <수정2.pptx> <새 파일.pptx>"""
import copy, re, sys, os, zipfile
from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Emu, Pt

SRC, DST = sys.argv[1], sys.argv[2]
U = 6350
E = lambda px: Emu(int(round(px * U)))
INK, MUTED, HAIR, SOFT, MID, ACC, WHITE, BG = "1B1E24", "5E6672", "C3C8D0", "D9DDE3", "B4BAC4", "EE4E22", "FFFFFF", "EFF0F2"
ACC_SOFT = "F6B29D"
FONT = "Noto Sans KR"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
PN = "http://schemas.openxmlformats.org/presentationml/2006/main"
P = "{%s}" % PN
rgb = lambda h: RGBColor.from_string(h)

prs = Presentation(SRC)
slides = list(prs.slides)
S = {n + 1: s for n, s in enumerate(slides)}  # 수정2 기준 쪽 번호


def find(slide, text):
    return [s for s in slide.shapes if s.has_text_frame and s.text_frame.text.strip().startswith(text)]


def remove(shape):
    shape._element.getparent().remove(shape._element)


def clear(slide, keep_text=()):
    for sh in list(slide.shapes):
        t = sh.text_frame.text.strip() if sh.has_text_frame else ""
        if not any(t.startswith(k) for k in keep_text): remove(sh)
    for old in slide._element.findall(P + "timing"): slide._element.remove(old)


def text(slide, s, x, y, w, h, size, color=INK, bold=False, align="l", anchor="ctr", name="text", lh=None):
    tb = slide.shapes.add_textbox(E(x), E(y), E(w), E(h)); tb.name = name
    tf = tb.text_frame; tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = {"ctr": MSO_ANCHOR.MIDDLE, "t": MSO_ANCHOR.TOP}[anchor]
    for i, ln in enumerate(s.split("\n")):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}[align]
        if lh: para.line_spacing = lh
        r = para.add_run(); r.text = ln
        f = r.font; f.size = Pt(size / 2); f.bold = bold; f.color.rgb = rgb(color); f.name = FONT
        rPr = r._r.get_or_add_rPr()
        for tag in ("ea", "cs"):
            el = etree.SubElement(rPr, "{%s}%s" % (A, tag)); el.set("typeface", FONT)
    return tb


def shape(slide, kind, x, y, w, h, fill=None, line=None, lw=2.25, dash=False, name="shape"):
    sp = slide.shapes.add_shape(kind, E(x), E(y), E(w), E(h)); sp.name = name
    sp.shadow.inherit = False
    if fill: sp.fill.solid(); sp.fill.fore_color.rgb = rgb(fill)
    else: sp.fill.background()
    if line:
        sp.line.color.rgb = rgb(line); sp.line.width = Pt(lw)
        if dash: etree.SubElement(sp.line._get_or_add_ln(), "{%s}prstDash" % A).set("val", "dash")
    else: sp.line.fill.background()
    return sp


def conn(slide, x1, y1, x2, y2, color, lw=1.5, dash=False, name="line"):
    c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, E(x1), E(y1), E(x2), E(y2)); c.name = name
    c.line.color.rgb = rgb(color); c.line.width = Pt(lw)
    if dash: etree.SubElement(c.line._get_or_add_ln(), "{%s}prstDash" % A).set("val", "dash")
    return c


def group(slide, shapes, name):
    g = slide.shapes.add_group_shape(shapes); g.name = name; return g


def note(slide, s):
    ns = slide.notes_slide
    tf = ns.notes_text_frame
    if tf is None:  # 새 슬라이드: 기존 노트의 본문 틀을 복사해 넣는다
        src = [p_ for p_ in S[7].notes_slide.placeholders if p_.placeholder_format.type is not None and "BODY" in str(p_.placeholder_format.type)][0]
        ns.shapes._spTree.append(copy.deepcopy(src._element))
        tf = ns.shapes[-1].text_frame; tf.text = ""
    tf.text = (tf.text.rstrip() + "\n\n" + s).strip()


def set_timing(slide, inner_pars, bld):
    xml = ('<p:timing xmlns:p="%s"><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>'
           '<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>' % PN + "".join(inner_pars) +
           '</p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
           '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>'
           '</p:childTnLst></p:cTn></p:par></p:tnLst><p:bldLst>' + "".join(bld) + '</p:bldLst></p:timing>')
    sld = slide._element
    for old in sld.findall(P + "timing"): sld.remove(old)
    new = etree.fromstring(xml); ext = sld.find(P + "extLst")
    if ext is not None: ext.addprevious(new)
    else: sld.append(new)


class Seq:
    """클릭 단계별 등장(나타나기/닦아내기)과 회전 강조를 만드는 작은 도우미"""
    EFF = {"fade": ("10", "0", "fade", 450), "wipe": ("22", "8", "wipe(left)", 800)}
    def __init__(self): self.i = 2; self.pars = []; self.bld = []
    def nid(self): self.i += 1; return self.i
    def entr(self, sh, eff, node):
        pid, sub, flt, dur = self.EFF[eff]; spid = sh.shape_id
        if sh._element.tag == P + "sp": self.bld.append(f'<p:bldP spid="{spid}" grpId="0" animBg="1"/>')
        return (f'<p:par><p:cTn id="{self.nid()}" presetID="{pid}" presetClass="entr" presetSubtype="{sub}" fill="hold" grpId="0" nodeType="{node}"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
                f'<p:set><p:cBhvr><p:cTn id="{self.nid()}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>'
                f'<p:animEffect transition="in" filter="{flt}"><p:cBhvr><p:cTn id="{self.nid()}" dur="{dur}"/><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl></p:cBhvr></p:animEffect></p:childTnLst></p:cTn></p:par>')
    def click(self, targets):
        effs = "".join(self.entr(sh, eff, "clickEffect" if k == 0 else "withEffect") for k, (sh, eff) in enumerate(targets))
        self.pars.append(f'<p:par><p:cTn id="{self.nid()}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="{self.nid()}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>{effs}</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>')
    def spins_on_entry(self, sh, steps):
        """steps: [(각도, 시간ms, accel%, decel%)] 를 슬라이드 진입 시 차례로 재생"""
        spid = sh.shape_id; inner = ""; t = 0
        for k, (deg, dur, acc, dec) in enumerate(steps):
            attrs = (f' accel="{acc * 1000}"' if acc else "") + (f' decel="{dec * 1000}"' if dec else "")
            inner += (f'<p:par><p:cTn id="{self.nid()}" fill="hold"><p:stCondLst><p:cond delay="{t}"/></p:stCondLst><p:childTnLst>'
                      f'<p:par><p:cTn id="{self.nid()}" presetID="8" presetClass="emph" presetSubtype="0" fill="hold" grpId="{k}" nodeType="afterEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
                      f'<p:animRot by="{int(round(deg * 60000))}"><p:cBhvr><p:cTn id="{self.nid()}" dur="{dur}"{attrs} fill="hold"/><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl><p:attrNameLst><p:attrName>r</p:attrName></p:attrNameLst></p:cBhvr></p:animRot>'
                      f'</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>')
            self.bld.append(f'<p:bldP spid="{spid}" grpId="{k}" animBg="1"/>'); t += dur
        self.pars.append(f'<p:par><p:cTn id="{self.nid()}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/><p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond></p:stCondLst><p:childTnLst>{inner}</p:childTnLst></p:cTn></p:par>')
    def apply(self, slide): set_timing(slide, self.pars, self.bld)


def new_slide_after(ref_slide, title_text=None, source_text=None):
    """ref_slide 와 같은 배경·제목·참고문헌 줄을 가진 새 슬라이드를 바로 뒤에 넣는다"""
    s = prs.slides.add_slide(ref_slide.slide_layout)
    for sh in list(s.shapes): remove(sh)
    bg = ref_slide._element.cSld.find(P + "bg")
    if bg is not None: s._element.cSld.insert(0, copy.deepcopy(bg))
    for key, val in (("title", title_text), ("src", source_text)):
        if val is None: continue
        ref = REF_TITLE if key == "title" else REF_SRC
        el = copy.deepcopy(ref._element); s.shapes._spTree.append(el)
        sh = s.shapes[-1]; sh.text_frame.paragraphs[0].runs[0].text = val
    lst = prs.slides._sldIdLst; ids = list(lst)
    me = ids[-1]; pos = [i for i, x in enumerate(ids) if prs.slides.get(int(x.get("id"))) is ref_slide][0]
    lst.remove(me); lst.insert(pos + 1, me)
    return s


REF_TITLE = find(S[7], "대표사례")[0]
REF_SRC = find(S[7], "참고문헌")[0]
REF_TITLE_EL = copy.deepcopy(REF_TITLE._element)

# ============ 4장: 도넛 — 회전으로 밀려 나오게 (최고 속도로 출발 → 끝에서 감속 → 벽에 부딪혀 튕김) ============
s = S[4]
ring = [x for x in s.shapes if x.name.startswith("s ring")][0]
arc = [x for x in s.shapes if x.name.startswith("arc")][0]
def set_arc(sp, a1, a2, thick_ratio):
    av = sp._element.spPr.find("{%s}prstGeom" % A).find("{%s}avLst" % A)
    for gd in list(av): av.remove(gd)
    for nm_, v in (("adj1", a1 * 60000), ("adj2", a2 * 60000), ("adj3", thick_ratio * 100000)):
        g = etree.SubElement(av, "{%s}gd" % A); g.set("name", nm_); g.set("fmla", "val %d" % round(v))
# 주황 호: 최종 위치(12시~44.2%)에 놓되 꼬리 12도를 덮개 밑에 숨김. 회색 덮개는 12시 반대쪽(94~270도)을 가림.
arc.left, arc.top, arc.width, arc.height = E(162), E(282), E(556), E(556)
set_arc(arc, 258, 69.12, 74 / 556)
cover_el = copy.deepcopy(arc._element); arc._element.addnext(cover_el)
cover = [x for x in s.shapes if x._element is cover_el][0]; cover.name = "덮개(회색 고리 조각)"
cover.shape_id  # noqa
cover_el.find(P + "nvSpPr").find(P + "cNvPr").set("id", str(max(x.shape_id for x in s.shapes) + 1))
cover.left, cover.top, cover.width, cover.height = E(160), E(280), E(560), E(560)
set_arc(cover, 94, 270, 78 / 560)
cover.fill.solid(); cover.fill.fore_color.rgb = rgb(SOFT)
T4 = {x.name: x for x in s.shapes}
q = Seq()
q.spins_on_entry(arc, [(-160.12, 1, 0, 0), (160.12, 1150, 0, 28), (-6, 130, 0, 100), (6, 190, 100, 0)])
for r in ([T4["text 47"], T4["text 48"]], [T4["line 49"], T4["text 50"], T4["text 51"]], [T4["line 52"], T4["text 53"], T4["text 54"]]):
    q.click([(x, "fade") for x in r])
q.apply(s)
note(s, "[변경 3차] 도넛 애니메이션 방식을 바꿈. 주황 호가 회색 덮개 밑에서 회전해 12시 방향부터 밀려 나옴. 처음부터 최고 속도로 달리다 끝에서 감속하고, 끝점에서 뒤로 6도 튕겼다가 제자리로 돌아옴.")

# ============ 5장: 중복 정리 ============
s = S[5]
for t in ("관련 수입의 50% 이내", "7일", "청약철회 기간"):
    for sh in find(s, t):
        if sh.text_frame.text.strip() == t: remove(sh)
cap = [x for x in find(s, "과징금 상한") if x.text_frame.text.strip() == "과징금 상한"][0]
cap.text_frame.paragraphs[0].runs[0].text = "과징금 상한 · 관련 수입 대비"
text(s, "1년 · 5년", 128, 790, 520, 80, 60, bold=True, name="해지 기한")
text(s, "위법계약해지 기한 (안 날부터 · 계약일부터)", 128, 874, 640, 44, 28, color=MUTED, name="해지 기한 설명")
note(s, "[변경 3차] 겹치던 문구를 정리함. 과징금은 큰 50% 하나로만 강조하고 위의 작은 글씨를 지움. 왼쪽 아래 ‘7일’을 지우고 위법계약해지권의 행사 기한을 작게 넣음: 위반을 안 날부터 1년, 계약일부터 5년 이내(법 제47조).")

# ============ 6장: 와플 — 손실 왼쪽, 모핑으로 손실 확대 ============
s6 = S[6]
clear(s6, keep_text=("실제 구제", "참고문헌"))
def waffle(slide, x0, y0, lw, lh_, pw, cols_gap, big):
    """손실 44칸(왼쪽, 11열) + 남은 원금 56칸(오른쪽, 14열). 칸 이름을 맞춰 모핑이 같은 칸으로 인식하게 함"""
    lp, lcw, pp, pcw, ph = lw
    cells = []
    for i in range(44):
        c, r = divmod(i, 4)
        cells.append(shape(slide, MSO_SHAPE.RECTANGLE, x0 + c * lp, y0 + r * lh_[0], lcw, lh_[1], fill=ACC_SOFT, name=f"!!손실칸{i + 1}"))
    px0 = x0 + 11 * lp + cols_gap
    for i in range(56):
        c, r = divmod(i, 4)
        cells.append(shape(slide, MSO_SHAPE.RECTANGLE, px0 + c * pp, y0 + r * lh_[0], pcw, lh_[1], fill=SOFT, name=f"!!원금칸{i + 1}"))
    return px0
# 6-1: 실제 비율
px0 = waffle(s6, 128, 330, (66.5, 56, 66.5, 56, 98), (108, 98), None, 0, False)
text(s6, "원금 10.4조원", 128, 196, 420, 60, 44, bold=True, name="!!원금 제목")
text(s6, "한 칸 = 약 0.1조원", 560, 204, 500, 46, 30, color=MUTED, name="!!단위")
text(s6, "손실 4.6조원", 128, 258, 700, 60, 44, color=ACC, bold=True, name="!!손실 라벨")
text(s6, "남은 원금 5.8조원", px0, 258, 800, 60, 44, color=MUTED, bold=True, name="!!원금 라벨")
note(s6, "[변경 3차] 손실을 왼쪽, 남은 원금을 오른쪽으로 바꿈. 이 장은 실제 비율(손실 44칸, 남은 원금 56칸)이고, 다음 장으로 넘어가면 모핑 전환으로 손실 부분이 가로로 늘어나고 남은 원금은 오른쪽에 좁게 줄어듦. 아래 수치 줄(15.9만 계좌, 동의율 93.8%)은 화면에서 빼고 말로 설명함. 31.4%는 다음 장의 배상 옆에 붙임.")
# 6-2: 손실 확대
s6b = new_slide_after(s6, "실제 구제", "참고문헌 [1] [5]")
LP, LCW, PP, PCW, RH, CH = 112, 102, 27.4, 20, 108, 98
gy = 330
px0 = waffle(s6b, 128, gy, (LP, LCW, PP, PCW, CH), (RH, CH), None, 34, True)
text(s6b, "원금 10.4조원", 128, 196, 420, 60, 44, bold=True, name="!!원금 제목")
text(s6b, "한 칸 = 약 0.1조원", 560, 204, 500, 46, 30, color=MUTED, name="!!단위")
text(s6b, "손실 4.6조원", 128, 258, 700, 60, 44, color=ACC, bold=True, name="!!손실 라벨")
text(s6b, "5.8조원", px0, 258, 386, 60, 44, color=MUTED, bold=True, name="!!원금 라벨")
split = []
for i in range(44):
    c, r = divmod(i, 4); x, y = 128 + c * LP, gy + r * RH
    if i < 13: split.append(shape(s6b, MSO_SHAPE.RECTANGLE, x + 2, y + 2, LCW - 4, CH - 4, fill=BG, line=INK, lw=2.25, name="배상 칸"))
    elif i < 40: split.append(shape(s6b, MSO_SHAPE.RECTANGLE, x, y, LCW, CH, fill=ACC, name="부담 칸"))
    else: split.append(shape(s6b, MSO_SHAPE.RECTANGLE, x, y, LCW, CH, fill=MID, name="그 밖 칸"))
g_split = group(s6b, split, "손실의 구성")
ly = gy + 4 * RH + 20
leg = []
for x, wd, lab, val, col, fill, ln_ in ((128, 430, "배상 · 평균 31.4%", "1.3조원", INK, BG, INK), (128 + 3.4 * LP, 560, "소비자 부담", "2.8조원", ACC, ACC, None), (128 + 10 * LP - 110, 330, "그 밖", "0.5조원", MUTED, MID, None)):
    sw = shape(s6b, MSO_SHAPE.RECTANGLE, x, ly + 10, 40, 40, fill=fill, line=ln_, name=f"범례 {lab}")
    a = text(s6b, lab, x + 56, ly, wd, 56, 32, color=MUTED, name=f"범례 이름 {lab}")
    b = text(s6b, val, x + 56, ly + 54, wd, 96, 72 if lab == "소비자 부담" else 56, color=col, bold=True, name=f"범례 값 {lab}")
    leg += [sw, a, b]
q = Seq(); q.click([(g_split, "fade")] + [(z, "fade") for z in leg]); q.apply(s6b)
MORPH = ('<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main" Requires="p159">'
         '<p:transition xmlns:p="%s" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" spd="slow" p14:dur="1600"><p159:morph option="byObject"/></p:transition></mc:Choice>'
         '<mc:Fallback><p:transition xmlns:p="%s" spd="slow"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent>' % (PN, PN))
tr = etree.fromstring(MORPH); s6b._element.find(P + "timing").addprevious(tr)
note(s6b, "[변경 3차] 앞 장에서 모핑 전환으로 이어지는 장. 손실 44칸이 화면의 약 4분의 3으로 늘어나고 남은 원금 56칸은 오른쪽에 좁게 줄어 ‘5.8조원’만 표시됨. 클릭하면 손실이 배상 1.3조원(윤곽선 13칸, 평균 배상비율 31.4%), 소비자 부담 2.8조원(주황 27칸), 그 밖 0.5조원(진회색 4칸)으로 갈림. 칸 수는 반올림한 값. 모핑은 PowerPoint 2019 이상·Microsoft 365에서 동작하며, 구버전에서는 서서히 바뀌는 전환으로 보임.")

# ============ 7장: 대표사례 — 소비자 부담 비율로 ============
s = S[7]
clear(s, keep_text=("참고문헌",))
s.shapes._spTree.insert(2, copy.deepcopy(REF_TITLE_EL))
text(s, "판매 과정의 문제", 380, 214, 640, 44, 30, color=MUTED, bold=True, name="머리글 1")
text(s, "소비자 부담 비율", 1060, 214, 732, 44, 30, color=MUTED, bold=True, name="머리글 2")
conn(s, 128, 270, 1792, 270, INK, lw=2.25, name="머리선")
rows = [("하나", "손실위험 설명 누락", 70), ("신한", "70대 고객의 답변 유도", 45), ("SC제일", "첫 투자 고객 사후확인 부실", 45), ("국민", "예금하러 온 고객에게 권유", 40), ("농협", "고령자 보호절차 미준수", 35)]
for i, (bank, why, v) in enumerate(rows):
    y = 270 + i * 122
    conn(s, 128, y + 122, 1792, y + 122, HAIR, lw=1.5, name="줄선")
    text(s, bank, 128, y, 240, 122, 46, bold=True, name=f"{bank}")
    text(s, why, 380, y, 660, 122, 36, name=f"{bank} 문제")
    bx, bw, bh = 1060, 520, 44
    shape(s, MSO_SHAPE.RECTANGLE, bx + bw * v / 100, y + 39, bw * (1 - v / 100), bh, line=INK, lw=2.25, name=f"{bank} 배상")
    shape(s, MSO_SHAPE.RECTANGLE, bx, y + 39, bw * v / 100, bh, fill=ACC, name=f"{bank} 부담")
    text(s, f"{v}%", bx + bw + 28, y, 180, 122, 54, bold=True, color=ACC if i == 0 else INK, name=f"{bank} 비율")
shape(s, MSO_SHAPE.RECTANGLE, 1060, 906, 28, 28, fill=ACC, name="범례 부담"); text(s, "소비자 부담", 1100, 898, 220, 44, 28, color=MUTED, name="범례 부담 글")
shape(s, MSO_SHAPE.RECTANGLE, 1340, 906, 28, 28, line=INK, lw=2.25, name="범례 배상"); text(s, "배상", 1380, 898, 160, 44, 28, color=MUTED, name="범례 배상 글")
note(s, "[변경 3차] 앞 장과 초점을 맞추려고 배상비율 대신 소비자 부담 비율(100 - 배상비율)로 다시 그림. 하나 70% / 신한 45% / SC제일 45% / 국민 40% / 농협 35%. 특정 은행만 크게 다루던 블록을 없애고 다섯 사례 각각에 판매 과정의 문제를 붙임. 부담 비율은 금감원 대표사례 배상비율(농협 65, 국민 60, 신한 55, SC제일 55, 하나 30%)에서 계산한 값이며, 5건 모두 금소법 시행(2021.3.24) 이전 판매분임.")

# ============ 새 장: 금소법 시행 전후의 유사 사례 ============
sc = new_slide_after(S[7], "유사 사례", "참고문헌 [1] [11]~[16]")
ly = 560
conn(sc, 128, ly, 1792, ly, INK, lw=2.25, name="연표선")
dx = 1452
conn(sc, dx, 220, dx, 900, ACC, lw=3, name="금소법 시행선")
text(sc, "금소법 시행 전", 128, 206, 600, 50, 34, color=MUTED, bold=True, name="시행 전")
text(sc, "시행 후", dx + 28, 206, 300, 50, 34, color=ACC, bold=True, name="시행 후")
text(sc, "2021.3", dx - 230, 206, 210, 50, 34, color=ACC, bold=True, align="r", name="시행 시점")
ev = [(210, "2008", "키코", "4개 기업", "2019년 분쟁조정\n배상 결정"), (430, "2011", "저축은행\n후순위채", "평균 42%", "불완전판매\n배상"), (650, "2013", "동양", "2만 명+", "분쟁조정 신청"),
      (870, "2019", "DLF", "4,468억원", "손실"), (1090, "2020", "라임", "1,611억원", "무역금융펀드\n전액 반환 결정"), (1310, "2020", "옵티머스", "5,146억원", "투자자 피해")]
for x, yr, nm_, num, cap in ev:
    shape(sc, MSO_SHAPE.OVAL, x - 20, ly - 20, 40, 40, fill=BG, line=INK, lw=2.25, name=f"{nm_.splitlines()[0]} 점")
    text(sc, yr, x - 105, 330, 210, 54, 42, bold=True, align="c", name=f"{yr} 연도")
    text(sc, nm_, x - 105, 388, 210, 130, 30, color=MUTED, align="c", anchor="t", name="사건")
    text(sc, num, x - 108, 606, 216, 56, 38, bold=True, align="c", name="수치")
    text(sc, cap, x - 108, 664, 216, 110, 26, color=MUTED, align="c", anchor="t", name="수치 설명")
ex = 1630
shape(sc, MSO_SHAPE.OVAL, ex - 130, ly - 130, 260, 260, fill=ACC, name="ELS 원")
text(sc, "2024", ex - 130, 330, 260, 54, 42, bold=True, color=ACC, align="c", name="ELS 연도")
text(sc, "홍콩 ELS", ex - 130, 384, 260, 44, 30, color=ACC, bold=True, align="c", name="ELS 이름")
text(sc, "4.6조원", ex - 130, ly - 52, 260, 70, 60, bold=True, color=WHITE, align="c", name="ELS 수치")
text(sc, "손실", ex - 130, ly + 18, 260, 44, 30, color=WHITE, align="c", name="ELS 설명")
note(sc, "[새 장 3차] 금소법 시행(2021.3) 전후로 나눈 유사 사례 연표. 수치는 모두 정부기관 자료에서 가져왔고, 사건마다 집계 기준이 달라 서로 크기를 비교하는 용도가 아님. "
     "키코(2008): 금감원 분조위가 2019.12.13 4개 피해기업의 분쟁조정에서 은행의 불완전판매 책임을 인정하고 손해액 일부 배상을 결정. "
     "저축은행 후순위채(2011): 2011.3월 말까지 42개사가 약 1.5조원 발행(금융위 2011.6.1). 금감원 분조위는 부산저축은행 등의 후순위채 불완전판매를 인정해 평균 42% 배상 책임 결정. "
     "동양(2013): 분쟁조정 신청 2만 명 이상, 신청 건의 67% 불완전판매 인정, 배상비율 15~50%(평균 22.9%)(금감원 2014.7.31 결정). "
     "DLF(2019): 손실 4,468억원, 투자자 3,243명. "
     "라임(2020): 금감원 분조위가 2018.11월 이후 판매된 무역금융펀드에 착오에 의한 계약취소를 결정, 최대 1,611억원(개인 500명, 법인 58개사) 반환 예상. 라임 펀드 전체의 결과가 아님. "
     "옵티머스(2020): 금융위 2021.11.24 인가·등록 취소 자료에 투자자 피해 총 5,146억원으로 기재. NH투자증권 판매분은 분조위가 투자원금 전액 반환을 결정(약 3천억원 예상). "
     "홍콩 ELS(2024): 금소법 시행 후 사건, 5대 판매은행 손실 4.6조원. 다만 손실이 난 상품의 상당수는 2021년 초 판매분임.")

# ============ 8장: 제목 삭제, 본문 6% 위로, 숫자와 주황 선 간격 벌리기 ============
s = S[8]
for sh in find(s, "제재 경과"): remove(sh)
for sh in s.shapes:
    t = sh.text_frame.text.strip() if sh.has_text_frame else ""
    if t.startswith("참고문헌"): continue
    sh.top = sh.top - E(90 if t.startswith("2년 10개월") else 65)
note(s, "[변경 3차] 주제 제목을 지우고 본문을 약 6% 위로 올림. ‘2년 10개월’과 주황 선 사이를 25px 더 벌려 선이 아래 연표에 속한 것으로 보이게 함.")

# ============ 9장: 트리맵만 ============
s = S[9]
clear(s)
X0, Y0, W, H, G = 115, 65, 1690, 950, 10
vals = [("국민", 1723.34, ACC, WHITE), ("하나", 924.75, INK, WHITE), ("농협", 401.75, "3A404B", WHITE), ("신한", 387.45, "59606C", WHITE), ("SC제일", 179.0, "7A828F", WHITE), ("기타", 130.03, MID, INK)]
tot = sum(v for _, v, _, _ in vals)
w1 = W * vals[0][1] / tot
rx, rw = X0 + w1 + G, W - w1 - G
h2 = H * vals[1][1] / (tot - vals[0][1])
by, bh = Y0 + h2 + G, H - h2 - G
rest = sum(v for _, v, _, _ in vals[2:])
w3 = (rw - 2 * G) * vals[2][1] / rest; w4 = (rw - 2 * G) * vals[3][1] / rest; w5 = rw - 2 * G - w3 - w4
h5 = (bh - G) * vals[4][1] / (vals[4][1] + vals[5][1])
boxes = [(X0, Y0, w1, H), (rx, Y0, rw, h2), (rx, by, w3, bh), (rx + w3 + G, by, w4, bh), (rx + w3 + w4 + 2 * G, by, w5, h5), (rx + w3 + w4 + 2 * G, by + h5 + G, w5, bh - h5 - G)]
sizes = [(64, 120), (52, 88), (40, 60), (40, 60), (34, 44), (30, 40)]
for (nm_, v, fill, fc), (x, y, w, h), (s1, s2) in zip(vals, boxes, sizes):
    shape(s, MSO_SHAPE.RECTANGLE, x, y, w, h, fill=fill, name=f"{nm_} 칸")
    pad = 36 if w > 300 else 22
    text(s, nm_, x + pad, y + pad - 6, w - 2 * pad, s1 * 1.3, s1, color=fc, bold=True, anchor="t", name=f"{nm_} 이름")
    text(s, f"{round(v):,}억", x + pad, y + pad - 6 + s1 * 1.3, w - 2 * pad, s2 * 1.25, s2, color=fc, bold=True, anchor="t", name=f"{nm_} 금액")
note(s, "[변경 3차] 주제 제목, 총액, 참고문헌 줄을 지우고 트리맵만 남김. 칸의 넓이가 금액에 비례함: 국민 1,723억 / 하나 925억 / 농협 402억 / 신한 387억 / SC제일 179억 / 기타 130억. 합계 3,746억원(은행 7곳·증권사 6곳, 금융위가 부과한 최대 규모 금전 제재). ‘기타’는 우리은행 과징금, 증권사 과징금·과태료, 은행 과태료를 합친 팀 계산값. 출처는 참고문헌 [2].")

# ============ 13장: Conclusion ============
s = S[13]
clear(s, keep_text=("결론",))
find(s, "결론")[0].text_frame.paragraphs[0].runs[0].text = "Conclusion"
cy, D, xa, xb = 540, 320, 430, 1490
shape(s, MSO_SHAPE.OVAL, xa - D / 2, cy - D / 2, D, D, line=INK, lw=3, name="제도 원")
shape(s, MSO_SHAPE.OVAL, xb - D / 2, cy - D / 2, D, D, fill=INK, name="현실 원")
text(s, "제도", xa - D / 2, cy - 44, D, 88, 64, bold=True, align="c", name="제도")
text(s, "현실", xb - D / 2, cy - 44, D, 88, 64, bold=True, color=WHITE, align="c", name="현실")
conn(s, xa + D / 2, cy, xb - D / 2, cy, INK, lw=3, dash=True, name="점선")
solid = conn(s, xa + D / 2, cy, xb - D / 2, cy, INK, lw=6, name="실선")
mx, mw = xa + D / 2 + 30, xb - xa - D - 60
t1 = text(s, "원칙 중심 규제", mx, cy - 150, mw, 80, 60, bold=True, align="c", name="과제 1")
t1s = text(s, "조문 준수에서 소비자 최선이익으로", mx, cy - 72, mw, 48, 30, color=MUTED, align="c", name="과제 1 설명")
t2 = text(s, "신속한 제재", mx, cy + 28, mw, 80, 60, bold=True, align="c", name="과제 2")
t2s = text(s, "위반에는 빠르고 예측 가능한 제재", mx, cy + 106, mw, 48, 30, color=MUTED, align="c", name="과제 2 설명")
q = Seq(); q.click([(solid, "wipe"), (t1, "fade"), (t1s, "fade"), (t2, "fade"), (t2s, "fade")]); q.apply(s)
note(s, "[변경 3차] 제목을 Conclusion으로 바꾸고, ELS 한 사건의 수치(손실률·배상비율·제재 기간)와 그림 밑 글씨를 모두 뺌. 제도(윤곽선)와 현실(채운 원)을 화면 가운데 크게 두고, 둘을 잇는 선 위아래에 과제 두 가지를 놓음. 처음에는 점선만 보이고, 클릭하면 실선으로 이어지며 과제가 나타남.")

# ============ 참고문헌 (2) ============
REFS = find(S[14], "참고문헌")
s14b = prs.slides.add_slide(S[14].slide_layout)
for sh in list(s14b.shapes): remove(sh)
bg = S[14]._element.cSld.find(P + "bg")
if bg is not None: s14b._element.cSld.insert(0, copy.deepcopy(bg))
tt = [x for x in REFS if x.text_frame.text.strip() == "참고문헌"][0]
s14b.shapes._spTree.append(copy.deepcopy(tt._element)); s14b.shapes[-1].text_frame.paragraphs[0].runs[0].text = "참고문헌 (2)"
extra = ["[11] 금융감독원, 「금융분쟁조정위원회, 키코 불완전판매 배상결정」, 2019.12.13.",
         "[12] 금융위원회, 「저축은행 후순위채 관련 제도개선 방안」, 2011.6.1.",
         "[13] 금융감독원, 「금융분쟁조정위원회, 부산저축은행 등에 대하여 후순위채 불완전판매를 인정하여 평균 42%의 손해배상책임을 결정」, 2011.",
         "[14] 금융감독원, 「동양증권 분쟁조정 결정…투자액 64.3% 회수」(정책브리핑), 2014.8.5.",
         "[15] 금융감독원, 「금융분쟁조정위원회, 무역금융펀드 투자원금 전액 반환 결정」, 2020.7.1.",
         "[16] 금융위원회, 「옵티머스자산운용㈜에 대한 금융투자업 인가·등록 취소 및 신탁계약 인계명령」, 2021.11.24."]
for i, t in enumerate(extra):
    text(s14b, t, 128, 230 + i * 92, 1664, 80, 28, anchor="t", name=f"문헌 {11 + i}")

prs.save(DST)
tmp = DST + ".tmp"; zin = zipfile.ZipFile(DST); zout = zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED)
for it in zin.infolist():
    d = zin.read(it.filename)
    if it.filename == "[Content_Types].xml" and b'Extension="jpg"' not in d:
        d = d.replace(b"<Default ", b'<Default Extension="jpg" ContentType="image/jpeg"/><Default ', 1)
    zout.writestr(it, d)
zout.close(); zin.close(); os.replace(tmp, DST)
print("saved", DST, len(prs.slides), "slides")
