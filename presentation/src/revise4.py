"""4차 수정: '수정3' 복사본의 4·7·8쪽만 고친다. 사용: python3 revise4.py <수정3.pptx> <새 파일.pptx>"""
import copy, sys, os, zipfile
from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Emu, Pt

SRC, DST = sys.argv[1], sys.argv[2]
E = lambda px: Emu(int(round(px * 6350)))
INK, MUTED, SOFT, MID, ACC, BG = "1B1E24", "5E6672", "D9DDE3", "B4BAC4", "EE4E22", "EFF0F2"
FONT = "Noto Sans KR"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
PN = "http://schemas.openxmlformats.org/presentationml/2006/main"
P = "{%s}" % PN
rgb = lambda h: RGBColor.from_string(h)
prs = Presentation(SRC)
S = {n + 1: s for n, s in enumerate(prs.slides)}


def remove(sh): sh._element.getparent().remove(sh._element)


def text(slide, s, x, y, w, h, size, color=INK, bold=False, align="l", anchor="ctr", name="text"):
    tb = slide.shapes.add_textbox(E(x), E(y), E(w), E(h)); tb.name = name
    tf = tb.text_frame; tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = {"ctr": MSO_ANCHOR.MIDDLE, "t": MSO_ANCHOR.TOP, "b": MSO_ANCHOR.BOTTOM}[anchor]
    para = tf.paragraphs[0]; para.alignment = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}[align]
    r = para.add_run(); r.text = s
    f = r.font; f.size = Pt(size / 2); f.bold = bold; f.color.rgb = rgb(color); f.name = FONT
    for tag in ("ea", "cs"): etree.SubElement(r._r.get_or_add_rPr(), "{%s}%s" % (A, tag)).set("typeface", FONT)
    return tb


def rect(slide, x, y, w, h, fill, name):
    sp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, E(x), E(y), E(w), E(h)); sp.name = name
    sp.shadow.inherit = False; sp.fill.solid(); sp.fill.fore_color.rgb = rgb(fill); sp.line.fill.background(); return sp


def note(slide, s):
    tf = slide.notes_slide.notes_text_frame; tf.text = (tf.text.rstrip() + "\n\n" + s).strip()


class Seq:
    def __init__(self): self.i = 2; self.pars = []; self.bld = []
    def nid(self): self.i += 1; return self.i
    def _entr(self, sh, node, pid="10", sub="0", flt="fade", dur=450, grp=0):
        spid = sh.shape_id
        if sh._element.tag == P + "sp": self.bld.append(f'<p:bldP spid="{spid}" grpId="{grp}" animBg="1"/>')
        return (f'<p:par><p:cTn id="{self.nid()}" presetID="{pid}" presetClass="entr" presetSubtype="{sub}" fill="hold" grpId="{grp}" nodeType="{node}"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
                f'<p:set><p:cBhvr><p:cTn id="{self.nid()}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>'
                f'<p:animEffect transition="in" filter="{flt}"><p:cBhvr><p:cTn id="{self.nid()}" dur="{dur}"/><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl></p:cBhvr></p:animEffect></p:childTnLst></p:cTn></p:par>')
    def _spin(self, sh, deg, dur, acc, dec, grp):
        spid = sh.shape_id; self.bld.append(f'<p:bldP spid="{spid}" grpId="{grp}" animBg="1"/>')
        at = (f' accel="{acc * 1000}"' if acc else "") + (f' decel="{dec * 1000}"' if dec else "")
        return (f'<p:par><p:cTn id="{self.nid()}" presetID="8" presetClass="emph" presetSubtype="0" fill="hold" grpId="{grp}" nodeType="afterEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
                f'<p:animRot by="{int(round(deg * 60000))}"><p:cBhvr><p:cTn id="{self.nid()}" dur="{dur}"{at} fill="hold"/><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl><p:attrNameLst><p:attrName>r</p:attrName></p:attrNameLst></p:cBhvr></p:animRot></p:childTnLst></p:cTn></p:par>')
    def click(self, shapes):
        effs = "".join(self._entr(sh, "clickEffect" if k == 0 else "withEffect") for k, sh in enumerate(shapes))
        self.pars.append(f'<p:par><p:cTn id="{self.nid()}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="{self.nid()}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>{effs}</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>')
    def on_entry(self, timed):
        """timed: [(시작 ms, xml 조각)] 을 슬라이드 진입 시 재생"""
        inner = "".join(f'<p:par><p:cTn id="{self.nid()}" fill="hold"><p:stCondLst><p:cond delay="{t}"/></p:stCondLst><p:childTnLst>{x}</p:childTnLst></p:cTn></p:par>' for t, x in timed)
        self.pars.append(f'<p:par><p:cTn id="{self.nid()}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/><p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond></p:stCondLst><p:childTnLst>{inner}</p:childTnLst></p:cTn></p:par>')
    def apply(self, slide):
        xml = ('<p:timing xmlns:p="%s"><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>'
               '<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>' % PN + "".join(self.pars) +
               '</p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
               '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>'
               '</p:childTnLst></p:cTn></p:par></p:tnLst><p:bldLst>' + "".join(self.bld) + '</p:bldLst></p:timing>')
        sld = slide._element
        for old in sld.findall(P + "timing"): sld.remove(old)
        new = etree.fromstring(xml); ext = sld.find(P + "extLst")
        if ext is not None: ext.addprevious(new)
        else: sld.append(new)


# ============ 4쪽: 도넛 — 12시에서 시계방향으로 '차오르게' (조각별 채우기) ============
s = S[4]
N4 = {x.name: x for x in s.shapes}
remove(N4["덮개(회색 고리 조각)"])
arc0 = N4["arc 44"]
def set_arc(sp, a, b, thick):
    av = sp._element.spPr.find("{%s}prstGeom" % A).find("{%s}avLst" % A)
    for gd in list(av): av.remove(gd)
    for nm_, v in (("adj1", ((270 + a) % 360) * 60000), ("adj2", ((270 + b) % 360) * 60000), ("adj3", thick * 100000)):
        g = etree.SubElement(av, "{%s}gd" % A); g.set("name", nm_); g.set("fmla", "val %d" % round(v))
END = 360 * 0.442  # 159.12도
# 45도 단위 네 조각. 바퀴살 8개 채우기는 45도마다 동시에 쓸어 가므로, 조각마다 시간을 달리 주면 속도가 변한다.
segs = [(0, 45), (45, 90), (90, 135), (135, END)]
shapes4 = []
next_id = max(x.shape_id for x in s.shapes) + 1
for k, (a, b) in enumerate(segs):
    el = arc0._element if k == 0 else copy.deepcopy(arc0._element)
    if k > 0:
        shapes4[-1]._element.addnext(el)
        el.find(P + "nvSpPr").find(P + "cNvPr").set("id", str(next_id)); next_id += 1
    sp = [x for x in s.shapes if x._element is el][0]; sp.name = f"도넛 조각 {k + 1}"
    sp.rotation = 0
    sp.left, sp.top, sp.width, sp.height = E(160), E(280), E(560), E(560)
    set_arc(sp, a - (0.8 if k else 0), b, 78 / 560)  # 이음매가 보이지 않게 0.8도 겹침
    shapes4.append(sp)
q = Seq()
# 45도를 쓸어 가는 시간(ms): 처음 두 조각은 최고 속도, 셋째부터 감속, 마지막은 끝점에 '부딪히듯' 짧게 멈춤
durs = [150, 150, 260, 520]
t = 0; timed = []
for k, (sp, d) in enumerate(zip(shapes4, durs)):
    timed.append((t, q._entr(sp, "afterEffect", pid="21", sub="8", flt="wheel(8)", dur=d)))
    t += d * (segs[k][1] - segs[k][0]) / 45
t = int(round(t))
last = shapes4[-1]
timed += [(t, q._spin(last, -7, 110, 0, 100, 1)), (t + 110, q._spin(last, 7, 150, 100, 0, 2)), (t + 260, q._spin(last, -2.5, 80, 0, 100, 3)), (t + 340, q._spin(last, 2.5, 110, 100, 0, 4))]
q.on_entry(timed)
for r in (["text 47", "text 48"], ["line 49", "text 50", "text 51"], ["line 52", "text 53", "text 54"]):
    q.click([N4[n] for n in r])
q.apply(s)
note(s, "[변경 4차] 도넛 채우기를 다시 만듦. 주황 부분을 네 조각으로 나눠 12시에서 시계방향으로 차례로 채움. 앞의 두 조각은 최고 속도, 셋째부터 감속, 마지막 조각이 끝점에 닿으면 그 조각만 뒤로 7도 튕겼다 돌아오고 한 번 더 작게 튕김.")

# ============ 7쪽: 범례를 블록 아래에 맞추고 위계를 줌 ============
s = S[7]
for sh in list(s.shapes):
    if sh.name.startswith("범례"): remove(sh)
g_split = [x for x in s.shapes if x.name == "손실의 구성"][0]
LP, gy, RH = 112, 330, 108
ly = gy + 4 * RH + 14
xs_b, xs_c, xs_e = 128, 128 + 4 * LP, 128 + 10 * LP
new = [
    text(s, "배상", xs_b, ly, 400, 50, 36, bold=True, name="배상 이름"),
    text(s, "1.3조원", xs_b, ly + 50, 400, 84, 64, bold=True, name="배상 값"),
    text(s, "평균 배상비율 31.4%", xs_b, ly + 136, 420, 44, 30, color=MUTED, name="배상 비율"),
    text(s, "소비자 부담", xs_c, ly, 600, 50, 36, color=ACC, bold=True, name="부담 이름"),
    text(s, "2.8조원", xs_c, ly + 44, 640, 140, 112, color=ACC, bold=True, name="부담 값"),
    text(s, "그 밖", xs_e, ly, 300, 50, 30, color=MUTED, name="그 밖 이름"),
    text(s, "0.5조원", xs_e, ly + 50, 300, 60, 40, color=MUTED, bold=True, name="그 밖 값"),
]
q = Seq(); q.click([g_split] + new); q.apply(s)
note(s, "[변경 4차] 아래 수치를 각 블록 바로 밑에 맞춰 놓고 크기로 위계를 줌: 소비자 부담 2.8조원이 가장 크고(주황), 배상 1.3조원이 그다음, 그 밖 0.5조원은 작게. 색 네모 범례는 없앰.")

# ============ 8쪽: 대표사례 막대 정리 ============
s = S[8]
banks = [("하나", 70), ("신한", 45), ("SC제일", 45), ("국민", 40), ("농협", 35)]
for sh in list(s.shapes):
    if sh.name.startswith("범례") or sh.name.endswith(" 배상") or sh.name.endswith(" 부담") or sh.name.endswith(" 비율"): remove(sh)
for i, (bank, v) in enumerate(banks):
    y = 270 + i * 122; bx, bw, bh = 1060, 520, 44
    rect(s, bx, y + 39, bw, bh, SOFT, f"{bank} 바탕")
    rect(s, bx, y + 39, bw * v / 100, bh, ACC, f"{bank} 부담")
    text(s, f"{v}%", bx + bw + 28, y, 180, 122, 54, color=ACC, bold=True, name=f"{bank} 비율")
note(s, "[변경 4차] 막대에서 배상 쪽 검은 테두리를 없애고 옅은 회색 바탕 위에 주황(소비자 부담)만 남김. 숫자도 주황으로 통일하고 아래 범례를 지움. 머리글 ‘소비자 부담 비율’ 하나로 읽히게 함.")

prs.save(DST)
tmp = DST + ".tmp"; zin = zipfile.ZipFile(DST); zout = zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED)
for it in zin.infolist():
    d = zin.read(it.filename)
    if it.filename == "[Content_Types].xml" and b'Extension="jpg"' not in d:
        d = d.replace(b"<Default ", b'<Default Extension="jpg" ContentType="image/jpeg"/><Default ', 1)
    zout.writestr(it, d)
zout.close(); zin.close(); os.replace(tmp, DST)
print("saved", DST)
