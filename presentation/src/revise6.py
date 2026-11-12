"""6차 수정: '수정5' 복사본의 4·5·7쪽을 고치고 전체 글꼴을 Pretendard로 바꾼다. 사용: python3 revise6.py <수정5.pptx> <새 파일.pptx>"""
import copy, math, sys, os, zipfile
from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Emu

SRC, DST = sys.argv[1], sys.argv[2]
E = lambda px: Emu(int(round(px * 6350)))
PX = lambda emu: emu / 6350
INK = "1B1E24"
OLD_FONT, NEW_FONT = "Noto Sans KR", "Pretendard"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
PN = "http://schemas.openxmlformats.org/presentationml/2006/main"
P = "{%s}" % PN
prs = Presentation(SRC)
S = {n + 1: s for n, s in enumerate(prs.slides)}


def remove(sh): sh._element.getparent().remove(sh._element)
def named(slide): return {x.name: x for x in slide.shapes}
def note(slide, s):
    tf = slide.notes_slide.notes_text_frame; tf.text = (tf.text.rstrip() + "\n\n" + s).strip()


def set_arc(sp, a, b, thick):
    """12시를 0도로 하고 시계방향으로 a도에서 b도까지의 고리 조각"""
    av = sp._element.spPr.find("{%s}prstGeom" % A).find("{%s}avLst" % A)
    for gd in list(av): av.remove(gd)
    for nm_, v in (("adj1", ((270 + a) % 360) * 60000), ("adj2", ((270 + b) % 360) * 60000), ("adj3", thick * 100000)):
        g = etree.SubElement(av, "{%s}gd" % A); g.set("name", nm_); g.set("fmla", "val %d" % round(v))


class Seq:
    """나타나기/사라지기만 쓰는 타임라인"""
    def __init__(self): self.i = 2; self.pars = []; self.bld = []; self.grp = {}
    def nid(self): self.i += 1; return self.i
    def _grp(self, sh):
        g = self.grp.get(sh.shape_id, 0); self.grp[sh.shape_id] = g + 1
        if sh._element.tag == P + "sp": self.bld.append(f'<p:bldP spid="{sh.shape_id}" grpId="{g}" animBg="1"/>')
        return g
    def show(self, sh, node, delay=0, on=True, fade=False):
        spid = sh.shape_id; g = self._grp(sh)
        eff = (f'<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="{self.nid()}" dur="450"/><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl></p:cBhvr></p:animEffect>' if fade else "")
        return (f'<p:par><p:cTn id="{self.nid()}" presetID="{10 if fade else 1}" presetClass="{"entr" if on else "exit"}" presetSubtype="0" fill="hold" grpId="{g}" nodeType="{node}"><p:stCondLst><p:cond delay="{delay}"/></p:stCondLst><p:childTnLst>'
                f'<p:set><p:cBhvr><p:cTn id="{self.nid()}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="{"visible" if on else "hidden"}"/></p:to></p:set>'
                f'{eff}</p:childTnLst></p:cTn></p:par>')
    def _step(self, effs, cond):
        self.pars.append(f'<p:par><p:cTn id="{self.nid()}" fill="hold"><p:stCondLst>{cond}</p:stCondLst><p:childTnLst><p:par><p:cTn id="{self.nid()}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>{effs}</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>')
    def click(self, shapes):
        self._step("".join(self.show(sh, "clickEffect" if k == 0 else "withEffect", fade=True) for k, sh in enumerate(shapes)), '<p:cond delay="indefinite"/>')
    def on_entry(self, events):
        """events: [(ms, 도형, 켜기 여부)] — 슬라이드에 들어오면 자동 재생 ('이전 효과와 함께' + 지연)"""
        events = sorted(events, key=lambda e: e[0])
        effs = "".join(self.show(sh, "afterEffect" if k == 0 else "withEffect", delay=t, on=on) for k, (t, sh, on) in enumerate(events))
        self._step(effs, '<p:cond delay="indefinite"/><p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond>')
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


# ============ 4쪽: 도넛 — 아주 얇은 조각이 차례로 켜지며 차오름 (가리기·회전 없음) ============
s = S[4]; N = named(s)
remove(N["도넛 주황(움직임)"]); remove(N["덮개(회색 고리 조각)"])
tmpl = N["도넛 주황(완성)"]
END = 360 * 0.442; TH1 = END * 0.90   # 끝점 159.12도, 감속 시작 143.2도
N1, N2 = 216, 48                      # 90%까지 216조각(약 0.66도), 남은 10%는 48조각(약 0.33도)
bounds = [TH1 * k / N1 for k in range(N1)] + [TH1 + (END - TH1) * k / N2 for k in range(N2 + 1)]
pieces = []; prev = tmpl; next_id = max(x.shape_id for x in s.shapes) + 1
for k in range(N1 + N2):
    el = copy.deepcopy(tmpl._element); prev._element.addnext(el)
    el.find(P + "nvSpPr").find(P + "cNvPr").set("id", str(next_id)); next_id += 1
    sp = [x for x in s.shapes if x._element is el][0]; sp.name = f"도넛 {k + 1:03d}"
    a, b = bounds[k], bounds[k + 1]
    set_arc(sp, max(0, a - 0.3), b, 78 / 560)  # 앞 조각과 0.3도 겹쳐 이음매가 보이지 않게
    pieces.append((sp, (a + b) / 2)); prev = sp
remove(tmpl)

V1, V2 = 220.0, 40.0   # 도/초: 최고 속도, 끝점에 닿을 때의 속도
T0 = 300               # 슬라이드가 뜬 뒤 출발까지 (ms)
t1 = TH1 / V1 * 1000
def t_fill(th):
    """머리가 th도에 닿는 시각(ms). 90%까지 등속, 그 뒤는 각도에 비례해 속도가 줄어듦(급감속)"""
    if th <= TH1: return T0 + th / V1 * 1000
    L = END - TH1; v = V1 - (V1 - V2) * (th - TH1) / L
    return T0 + t1 + L / (V1 - V2) * math.log(V1 / v) * 1000
events = [(round(t_fill(mid)), sp, True) for sp, mid in pieces]
t = t_fill(END)
for back, t_out, t_in in ((4.0, 180, 220), (1.5, 110, 140)):  # 튕김: 뒤로 4도 → 복귀, 다시 1.5도 → 복귀
    for sp, mid in pieces:
        d = END - mid
        if d >= back: continue
        events.append((round(t + t_out * (1 - math.sqrt(1 - d / back))), sp, False))       # 뒤로 밀리며 감속
        events.append((round(t + t_out + t_in * math.sqrt((back - d) / back)), sp, True))  # 돌아오며 가속
    t += t_out + t_in
q = Seq(); q.on_entry(events)
for r in (["text 47", "text 48"], ["line 49", "text 50", "text 51"], ["line 52", "text 53", "text 54"]):
    q.click([N[n] for n in r])
q.apply(s)
note(s, f"[변경 6차] 도넛을 다시 만듦. 가리개나 회전을 쓰지 않고, 주황 부분을 얇은 조각 {N1 + N2}개(0.3~0.7도)로 나눠 12시부터 차례로 켬. "
        "90% 지점까지 최고 속도, 남은 10%에서 급감속, 끝점(44.2%)에 닿으면 뒤로 4도 밀렸다 돌아오고 한 번 더 작게(1.5도) 튕김. 전체 약 1.8초. "
        "편집 화면에서 주황 부분이 작은 도형 여러 개로 잡히는 것은 정상.")

# ============ 5쪽: '1년 · 5년' 아래 설명을 다른 설명글과 같은 서식으로 ============
s = S[5]; N = named(s)
ref, tgt = N["s text 62"], N["해지 기한 설명"]   # 기준: '위반 계약은 위약금 없이 해지'
ref_rpr = ref.text_frame.paragraphs[0].runs[0]._r.rPr
for r in tgt.text_frame.paragraphs[0].runs:
    old = r._r.rPr; new = copy.deepcopy(ref_rpr)
    if old is not None: r._r.replace(old, new)
    else: r._r.insert(0, new)
tgt.width = E(560)
note(s, "[변경 6차] ‘1년 · 5년’ 아래 설명글을 같은 열의 다른 설명글(위반 계약은 위약금 없이 해지)과 같은 크기·색·굵기로 통일.")

# ============ 7쪽: 손실 합계는 검정, 5.8조원은 회색 칸 오른쪽 끝에 맞춤 ============
s = S[7]; N = named(s)
for r in N["!!손실 라벨"].text_frame.paragraphs[0].runs: r.font.color.rgb = RGBColor.from_string(INK)
grey_right = max(PX(c.left + c.width) for c in s.shapes if c.name.startswith("!!원금칸"))
lab = N["!!원금 라벨"]; lab.left = E(grey_right - PX(lab.width))
for para in lab.text_frame.paragraphs: para.alignment = PP_ALIGN.RIGHT
note(s, "[변경 6차] 왼쪽 위 ‘손실 4.6조원’을 검정으로 바꿔 주황은 소비자 부담에만 남김. 오른쪽 위 ‘5.8조원’은 회색 칸의 오른쪽 끝에 맞춰 오른쪽 정렬(높이는 그대로).")
note(S[14], "[확인 6차] ‘고난도’ 표기는 그대로 둠. ‘고난도금융투자상품’은 자본시장법 시행령 제2조의 법령 용어(최대 원금손실 가능금액이 원금의 20%를 넘는 상품)이며, 2021.5.10 시행.")

prs.save(DST)
# 글꼴 일괄 교체 + jpg 형식 선언 보정
tmp = DST + ".tmp"; zin = zipfile.ZipFile(DST); zout = zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED); n_font = 0
for it in zin.infolist():
    d = zin.read(it.filename)
    if it.filename == "[Content_Types].xml" and b'Extension="jpg"' not in d:
        d = d.replace(b"<Default ", b'<Default Extension="jpg" ContentType="image/jpeg"/><Default ', 1)
    if it.filename.endswith(".xml") and OLD_FONT.encode() in d:
        n_font += d.count(OLD_FONT.encode()); d = d.replace(OLD_FONT.encode(), NEW_FONT.encode())
    zout.writestr(it, d)
zout.close(); zin.close(); os.replace(tmp, DST)
print("saved", DST, "| font refs replaced:", n_font, "| donut pieces:", N1 + N2, "| donut events:", len(events), "| end ms:", round(t))
