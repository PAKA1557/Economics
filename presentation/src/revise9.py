"""9차 수정: '수정8' 복사본의 4쪽 도넛에 숫자 카운트업과 충돌 순간의 '콩' 튀기를 더한다.
사용: python3 revise9.py <수정8.pptx> <새 파일.pptx>"""
import copy, math, sys, os, zipfile
from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Emu

SRC, DST = sys.argv[1], sys.argv[2]
E = lambda px: Emu(int(round(px * 6350)))
BG = "EFF0F2"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
PN = "http://schemas.openxmlformats.org/presentationml/2006/main"
P = "{%s}" % PN
prs = Presentation(SRC)
s = prs.slides[3]
N = {x.name: x for x in s.shapes}


class Seq:
    """나타나기/사라지기만 쓰는 타임라인 (6차와 같음)"""
    def __init__(self): self.i = 2; self.pars = []; self.bld = []; self.grp = {}
    def nid(self): self.i += 1; return self.i
    def show(self, sh, node, delay=0, on=True, fade=False):
        spid = sh.shape_id; g = self.grp.get(spid, 0); self.grp[spid] = g + 1
        self.bld.append(f'<p:bldP spid="{spid}" grpId="{g}" animBg="1"/>')
        eff = (f'<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="{self.nid()}" dur="450"/><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl></p:cBhvr></p:animEffect>' if fade else "")
        return (f'<p:par><p:cTn id="{self.nid()}" presetID="{10 if fade else 1}" presetClass="{"entr" if on else "exit"}" presetSubtype="0" fill="hold" grpId="{g}" nodeType="{node}"><p:stCondLst><p:cond delay="{delay}"/></p:stCondLst><p:childTnLst>'
                f'<p:set><p:cBhvr><p:cTn id="{self.nid()}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="{"visible" if on else "hidden"}"/></p:to></p:set>'
                f'{eff}</p:childTnLst></p:cTn></p:par>')
    def _step(self, effs, cond):
        self.pars.append(f'<p:par><p:cTn id="{self.nid()}" fill="hold"><p:stCondLst>{cond}</p:stCondLst><p:childTnLst><p:par><p:cTn id="{self.nid()}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>{effs}</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>')
    def click(self, shapes):
        self._step("".join(self.show(sh, "clickEffect" if k == 0 else "withEffect", fade=True) for k, sh in enumerate(shapes)), '<p:cond delay="indefinite"/>')
    def on_entry(self, events):
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


# ---- 도넛의 움직임 (6차와 같은 값) ----
PCT = 44.2
END = 360 * PCT / 100; TH1 = END * 0.90
N1, N2 = 216, 48
bounds = [TH1 * k / N1 for k in range(N1)] + [TH1 + (END - TH1) * k / N2 for k in range(N2 + 1)]
pieces = [(N[f"도넛 {k + 1:03d}"], (bounds[k] + bounds[k + 1]) / 2) for k in range(N1 + N2)]
V1, V2, T0 = 220.0, 40.0, 300
L = END - TH1; t1 = TH1 / V1 * 1000; t2 = L / (V1 - V2) * math.log(V1 / V2) * 1000
BOUNCES = ((4.0, 180, 220), (1.5, 110, 140))
T_HIT = T0 + t1 + t2
T_END = T_HIT + sum(a + b for _, a, b in BOUNCES)
def t_fill(th):
    if th <= TH1: return T0 + th / V1 * 1000
    v = V1 - (V1 - V2) * (th - TH1) / L
    return T0 + t1 + L / (V1 - V2) * math.log(V1 / v) * 1000
def theta(t):
    """시각 t(ms)에 주황 머리가 있는 각도"""
    t -= T0
    if t <= 0: return 0.0
    if t <= t1: return V1 * t / 1000
    if t <= t1 + t2:
        v = V1 * math.exp(-(t - t1) / 1000 * (V1 - V2) / L); return TH1 + (V1 - v) * L / (V1 - V2)
    t -= t1 + t2
    for back, t_out, t_in in BOUNCES:
        if t <= t_out: return END - back * (1 - (1 - t / t_out) ** 2)
        t -= t_out
        if t <= t_in: return END - back + back * (t / t_in) ** 2
        t -= t_in
    return END
events = [(round(t_fill(mid)), sp, True) for sp, mid in pieces]
t = T_HIT
for back, t_out, t_in in BOUNCES:
    for sp, mid in pieces:
        d = END - mid
        if d >= back: continue
        events.append((round(t + t_out * (1 - math.sqrt(1 - d / back))), sp, False))
        events.append((round(t + t_out + t_in * math.sqrt((back - d) / back)), sp, True))
    t += t_out + t_in

# ---- 가운데 숫자: 프레임마다 글상자 하나. 뒤 프레임이 앞 프레임을 덮으며 켜진다 ----
# 각 글상자는 배경색으로 칠해져 있어 아래 프레임을 가린다. 맨 위(마지막)가 44.2%라서 편집 화면·PDF에는 44.2%만 보인다.
num = N["s text 45"]
BX, BY, BW, BH = 270, 469, 340, 140          # 도넛 구멍 안에 들어가는 크기
BASE_SZ = int(num.text_frame.paragraphs[0].runs[0]._r.rPr.get("sz"))
POP, POP_MS = 0.10, 320                        # 부딪히는 순간 10% 커졌다가 0.32초에 걸쳐 돌아옴
STEP = 33                                      # 프레임 간격(ms), 초당 약 30장
frames = [(0, 0.0, BASE_SZ)]                   # (시각, 값, 글자 크기) — 첫 프레임 0.0%는 처음부터 보임
tt = T0 + STEP
while tt < T_END + POP_MS + STEP:
    val = round(PCT * theta(min(tt, T_END)) / END, 1)
    k = 1 + POP * math.exp(-(tt - T_HIT) / (POP_MS / 3.5)) if tt >= T_HIT else 1.0
    if tt >= T_HIT + POP_MS: k = 1.0
    frames.append((round(tt), val, int(round(BASE_SZ * k / 50.0)) * 50)); tt += STEP
frames.append((round(tt), PCT, BASE_SZ))
frames.append((round(T_HIT), PCT, int(round(BASE_SZ * (1 + POP) / 50.0)) * 50))  # 부딪히는 바로 그 순간: 44.2%가 가장 크게
frames.sort(key=lambda f: f[0])
dedup = [frames[0]]
for f in frames[1:]:
    if f[1:] != dedup[-1][1:]: dedup.append(f)
frames = dedup
assert frames[-1][1:] == (PCT, BASE_SZ)
next_id = max(x.shape_id for x in s.shapes) + 1; prev = num; made = []
for k, (ft, val, sz) in enumerate(frames):
    el = copy.deepcopy(num._element); prev._element.addnext(el)
    el.find(P + "nvSpPr").find(P + "cNvPr").set("id", str(next_id)); next_id += 1
    sp = [x for x in s.shapes if x._element is el][0]
    sp.name = "숫자 %02d (%.1f%%)" % (k, val) if k < len(frames) - 1 else "숫자 완성 (44.2%)"
    sp.left, sp.top, sp.width, sp.height = E(BX), E(BY), E(BW), E(BH)
    sp.fill.solid(); sp.fill.fore_color.rgb = RGBColor.from_string(BG)
    sp.text_frame.word_wrap = False
    r = sp.text_frame.paragraphs[0].runs[0]; r.text = "%.1f%%" % val; r._r.rPr.set("sz", str(sz))
    if k: events.append((ft, sp, True))
    made.append(sp); prev = sp
num._element.getparent().remove(num._element)

q = Seq(); q.on_entry(events)
for row in (["text 47", "text 48"], ["line 49", "text 50", "text 51"], ["line 52", "text 53", "text 54"]):
    q.click([N[n] for n in row])
q.apply(s)
tf = s.notes_slide.notes_text_frame
tf.text = tf.text.rstrip() + (f"\n\n[변경 9차] 가운데 숫자가 도넛과 같은 속도로 0.0%에서 44.2%까지 올라감(90%에서 함께 느려짐). 주황이 끝점에 부딪히는 순간 숫자가 10% 커졌다 돌아오고, "
                              f"튕기는 동안 43.1% 안팎까지 내려갔다 44.2%로 돌아옴. 숫자 글상자 {len(frames)}개가 차례로 겹쳐 켜지는 방식이며, 편집 화면에는 맨 위의 44.2%만 보임.")

prs.save(DST)
tmp = DST + ".tmp"; zin = zipfile.ZipFile(DST); zout = zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED)
for it in zin.infolist():
    d = zin.read(it.filename)
    if it.filename == "[Content_Types].xml" and b'Extension="jpg"' not in d:
        d = d.replace(b"<Default ", b'<Default Extension="jpg" ContentType="image/jpeg"/><Default ', 1)
    zout.writestr(it, d)
zout.close(); zin.close(); os.replace(tmp, DST)
print("saved", DST, "| frames:", len(frames), "| hit ms:", round(T_HIT), "| end ms:", round(T_END))
print([(f[0], f[1], f[2]) for f in frames])
