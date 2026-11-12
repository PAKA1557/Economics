"""5차 수정: '수정4' 복사본의 4·5·6·7쪽만 고친다. 사용: python3 revise5.py <수정4.pptx> <새 파일.pptx>"""
import copy, sys, os, zipfile
from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Emu, Pt

SRC, DST = sys.argv[1], sys.argv[2]
E = lambda px: Emu(int(round(px * 6350)))
PX = lambda emu: emu / 6350
MUTED, SOFT = "5E6672", "D9DDE3"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
PN = "http://schemas.openxmlformats.org/presentationml/2006/main"
P = "{%s}" % PN
ALIGN = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}
prs = Presentation(SRC)
S = {n + 1: s for n, s in enumerate(prs.slides)}


def remove(sh): sh._element.getparent().remove(sh._element)


def named(slide): return {x.name: x for x in slide.shapes}


def place(sh, x=None, y=None, w=None, h=None, align=None):
    if x is not None: sh.left = E(x)
    if y is not None: sh.top = E(y)
    if w is not None: sh.width = E(w)
    if h is not None: sh.height = E(h)
    if align:
        for para in sh.text_frame.paragraphs: para.alignment = ALIGN[align]


def set_text(sh, s, size=None, color=None, bold=None):
    r = sh.text_frame.paragraphs[0].runs[0]; r.text = s
    for extra in sh.text_frame.paragraphs[0].runs[1:]: extra._r.getparent().remove(extra._r)
    if size: r.font.size = Pt(size / 2)
    if color: r.font.color.rgb = RGBColor.from_string(color)
    if bold is not None: r.font.bold = bold


def note(slide, s):
    tf = slide.notes_slide.notes_text_frame; tf.text = (tf.text.rstrip() + "\n\n" + s).strip()


def set_arc(sp, a, b, thick):
    """12시를 0도로 하고 시계방향으로 a도에서 b도까지의 고리 조각"""
    av = sp._element.spPr.find("{%s}prstGeom" % A).find("{%s}avLst" % A)
    for gd in list(av): av.remove(gd)
    for nm_, v in (("adj1", ((270 + a) % 360) * 60000), ("adj2", ((270 + b) % 360) * 60000), ("adj3", thick * 100000)):
        g = etree.SubElement(av, "{%s}gd" % A); g.set("name", nm_); g.set("fmla", "val %d" % round(v))


class Seq:
    def __init__(self): self.i = 2; self.pars = []; self.bld = []; self.grp = {}
    def nid(self): self.i += 1; return self.i
    def _grp(self, sh):
        g = self.grp.get(sh.shape_id, 0); self.grp[sh.shape_id] = g + 1
        if sh._element.tag == P + "sp": self.bld.append(f'<p:bldP spid="{sh.shape_id}" grpId="{g}" animBg="1"/>')
        return g
    def entr(self, sh, node, fade=True):
        spid = sh.shape_id; g = self._grp(sh)
        eff = (f'<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="{self.nid()}" dur="450"/><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl></p:cBhvr></p:animEffect>' if fade else "")
        return (f'<p:par><p:cTn id="{self.nid()}" presetID="{10 if fade else 1}" presetClass="entr" presetSubtype="0" fill="hold" grpId="{g}" nodeType="{node}"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
                f'<p:set><p:cBhvr><p:cTn id="{self.nid()}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>'
                f'{eff}</p:childTnLst></p:cTn></p:par>')
    def spin(self, sh, deg, dur, acc=0, dec=0):
        spid = sh.shape_id; g = self._grp(sh)
        at = (f' accel="{acc * 1000}"' if acc else "") + (f' decel="{dec * 1000}"' if dec else "")
        return (f'<p:par><p:cTn id="{self.nid()}" presetID="8" presetClass="emph" presetSubtype="0" fill="hold" grpId="{g}" nodeType="afterEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
                f'<p:animRot by="{int(round(deg * 60000))}"><p:cBhvr><p:cTn id="{self.nid()}" dur="{dur}"{at} fill="hold"/><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl><p:attrNameLst><p:attrName>r</p:attrName></p:attrNameLst></p:cBhvr></p:animRot></p:childTnLst></p:cTn></p:par>')
    def click(self, shapes):
        effs = "".join(self.entr(sh, "clickEffect" if k == 0 else "withEffect") for k, sh in enumerate(shapes))
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


# ============ 4쪽: 도넛 — 12시에서 실제로 차오르고, 90%에서 급감속, 끝점에 부딪혀 튕김 ============
# 방식: 주황 호 하나를 회색 덮개 밑(12시 바로 앞쪽)에 통째로 숨겨 두고 시계방향으로 돌린다.
# 호의 머리가 12시에서 빠져나오고 꼬리는 끝까지 덮개 밑에 있으므로, 조각이 생기는 것이 아니라 연속으로 차오른다.
s = S[4]; N = named(s)
END = 360 * 0.442  # 159.12도
for k in (2, 3, 4): remove(N[f"도넛 조각 {k}"])
still = N["도넛 조각 1"]; still.name = "도넛 주황(완성)"  # 편집 화면·PDF에 보이는 완성 상태. 쇼에서는 맨 끝에 나타남
place(still, 160, 280, 560, 560); still.rotation = 0; set_arc(still, 0, END, 78 / 560)
next_id = max(x.shape_id for x in s.shapes) + 1
def clone(src, name):
    global next_id
    el = copy.deepcopy(src._element); src._element.addnext(el)
    el.find(P + "nvSpPr").find(P + "cNvPr").set("id", str(next_id)); next_id += 1
    sp = [x for x in s.shapes if x._element is el][0]; sp.name = name; return sp
mover = clone(still, "도넛 주황(움직임)"); mover.rotation = 360 - END  # 시작 자세: 덮개 밑에 전부 숨음
cover = clone(mover, "덮개(회색 고리 조각)"); cover.rotation = 0
# 움직이는 호는 안팎으로 1px 좁게: 덮개 가장자리로 주황빛이 새어 보이지 않게
place(mover, 161, 281, 558, 558); set_arc(mover, 0, END, 76 / 558)
set_arc(cover, 360 - END - 3, 360, 78 / 560); cover.fill.solid(); cover.fill.fore_color.rgb = RGBColor.from_string(SOFT)
q = Seq()
# (각도, 시간 ms, 가속 %, 감속 %). 90%까지 초당 약 220도로 달리고, 남은 10%에서 136 → 71도/초로 급감속해 끝점에 닿은 뒤 두 번 튕긴다.
steps = [(END * 0.90, 650, 0, 0), (END * 0.06, 70, 0, 0), (END * 0.04, 90, 0, 0),
         (-5, 150, 0, 100), (5, 190, 100, 0), (-1.8, 90, 0, 100), (1.8, 120, 100, 0)]
t = 300; timed = []
for deg, dur, acc, dec in steps:
    timed.append((t, q.spin(mover, deg, dur, acc, dec))); t += dur
timed.append((t + 50, q.entr(still, "afterEffect", fade=False)))
q.on_entry(timed)
# 오른쪽 수치: 구분선을 화면 가장자리 쪽에서 10% 줄이고, 글자는 그 끝에 맞춰 오른쪽 정렬
LX, LW = 900, 892; LW2 = round(LW * 0.9); RX = LX + LW2
for nm in ("line 49", "line 52"): place(N[nm], LX, w=LW2)
for nm in ("text 47", "text 48", "text 50", "text 51", "text 53", "text 54"): place(N[nm], LX, w=LW2, align="r")
for r in (["text 47", "text 48"], ["line 49", "text 50", "text 51"], ["line 52", "text 53", "text 54"]):
    q.click([N[n] for n in r])
q.apply(s)
note(s, "[변경 5차] 도넛을 다시 만듦. 주황 호 하나가 12시에서 시계방향으로 끊김 없이 차오름. 90% 지점까지 최고 속도, 남은 10%에서 급감속, 끝점(44.2%)에 닿으면 뒤로 5도 튕겼다 돌아오고 한 번 더 작게 튕김. "
        "원리: 주황 호를 12시 앞쪽의 회색 덮개 밑에 숨겨 두고 돌려서 꺼냄. 편집 화면에서 주황 호가 두 개로 보이는 것은 정상(하나는 움직이는 것, 하나는 완성 상태). "
        "오른쪽 수치는 오른쪽 정렬로 바꾸고 구분선은 오른쪽 끝을 10% 줄여 글자 끝과 맞춤.")

# ============ 5쪽: 왼쪽 아래 해지 기한을 '입증책임 전환' 줄 높이에 맞추고, 50%는 오른쪽 정렬 ============
s = S[5]; N = named(s)
ref_t, ref_d = N["s text 69"], N["s text 70"]  # 입증책임 전환 / 그 설명
place(N["해지 기한"], y=PX(ref_t.top), h=PX(ref_t.height))
place(N["해지 기한 설명"], y=PX(ref_d.top), h=PX(ref_d.height))
set_text(N["해지 기한 설명"], "해지 기한 (안 날부터 · 계약일부터)"); place(N["해지 기한 설명"], w=520)  # 같은 줄의 가운데 열과 겹치지 않게 줄임
place(N["s text 75"], align="r"); place(N["s text 76"], align="r")
note(s, "[변경 5차] 왼쪽 아래 ‘1년 · 5년’과 그 설명을 가운데 열의 ‘입증책임 전환’·설명과 같은 높이로 올림. 옆 열과 겹치지 않도록 설명을 ‘해지 기한 (안 날부터 · 계약일부터)’로 줄임. 오른쪽 아래 50%와 그 설명은 오른쪽 정렬.")

# ============ 6쪽: 단위는 주석으로, 격자는 7% 아래로, 이름표는 한 줄로 ============
s = S[6]; N = named(s)
DY = round(1080 * 0.07)  # 화면 높이의 7% = 76px
cells = [x for x in s.shapes if x.name.startswith("!!손실칸") or x.name.startswith("!!원금칸")]
for c in cells: c.top = c.top + E(DY)
loss = [c for c in cells if c.name.startswith("!!손실칸")]; rest = [c for c in cells if c.name.startswith("!!원금칸")]
right = lambda cs: max(PX(c.left + c.width) for c in cs)
left = lambda cs: min(PX(c.left) for c in cs)
ROW = PX(N["!!손실 라벨"].top)  # 손실이 적혀 있던 줄
place(N["!!원금 제목"], 128, ROW)
place(N["!!손실 라벨"], right(loss) - 400, ROW, 400, align="r")
place(N["!!원금 라벨"], right(rest) - 600, ROW, 600, align="r")
def footnote(slide):
    u = named(slide)["!!단위"]
    set_text(u, "* 한 칸 = 약 0.1조원", size=22, color=MUTED, bold=False)
    place(u, 128, 996, 700, 34, align="l")
footnote(s)
note(s, "[변경 5차] ‘한 칸 = 약 0.1조원’을 제목 옆에서 빼 왼쪽 아래 주석으로 내림. 격자를 화면 높이의 7%만큼 아래로 내림. "
        "원금 10.4조원을 손실이 적혀 있던 줄로 내리고, 손실 4.6조원은 주황 칸의 오른쪽 끝, 남은 원금 5.8조원은 회색 칸의 오른쪽 끝에 맞춰 오른쪽 정렬(세 이름표가 한 줄).")

# ============ 7쪽: 소비자 부담은 주황 칸 기준 가운데, '그 밖'은 '기타' ============
s = S[7]; N = named(s)
grp = N["손실의 구성"]
orange = [c for c in grp.shapes if c.name == "부담 칸"]
ox1, ox2 = left(orange), right(orange)
for nm in ("부담 이름", "부담 값"): place(N[nm], ox1, w=ox2 - ox1, align="c")
set_text(N["그 밖 이름"], "기타"); N["그 밖 이름"].name = "기타 이름"; N["그 밖 값"].name = "기타 값"
footnote(s)  # 앞 장과 같은 자리의 주석 (모핑 전환에서 그대로 유지됨)
note(s, "[변경 5차] ‘소비자 부담 2.8조원’을 주황 칸 전체 폭의 가운데에 맞춤. ‘그 밖’을 ‘기타’로 바꿈(자리는 그대로). 배상 쪽은 그대로. ‘한 칸 = 약 0.1조원’은 앞 장과 같이 왼쪽 아래 주석으로 옮김.")

prs.save(DST)
tmp = DST + ".tmp"; zin = zipfile.ZipFile(DST); zout = zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED)
for it in zin.infolist():
    d = zin.read(it.filename)
    if it.filename == "[Content_Types].xml" and b'Extension="jpg"' not in d:
        d = d.replace(b"<Default ", b'<Default Extension="jpg" ContentType="image/jpeg"/><Default ', 1)
    zout.writestr(it, d)
zout.close(); zin.close(); os.replace(tmp, DST)
print("saved", DST)
