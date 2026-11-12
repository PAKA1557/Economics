"""7차 수정: '수정6' 복사본에 5쪽 집중(모핑) 3장 추가, 대표사례 막대 채우기, 개선방안 강조 애니메이션.
사용: python3 revise7.py <수정6.pptx> <새 파일.pptx>"""
import copy, re, sys, os, zipfile
from lxml import etree
from pptx import Presentation
from pptx.util import Emu

SRC, DST = sys.argv[1], sys.argv[2]
E = lambda px: Emu(int(round(px * 6350)))
PX = lambda emu: emu / 6350
INK, MID, SOFT, ACC = "1B1E24", "B4BAC4", "D9DDE3", "EE4E22"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
PN = "http://schemas.openxmlformats.org/presentationml/2006/main"
P = "{%s}" % PN
prs = Presentation(SRC)
S = {n + 1: s for n, s in enumerate(prs.slides)}


def named(slide): return {x.name: x for x in slide.shapes}
def note(slide, s, replace=False):
    tf = slide.notes_slide.notes_text_frame; tf.text = s if replace else (tf.text.rstrip() + "\n\n" + s).strip()


class Seq:
    def __init__(self): self.i = 2; self.pars = []; self.bld = []; self.grp = {}
    def nid(self): self.i += 1; return self.i
    def _head(self, sh, pid, cls, sub, node, delay):
        g = self.grp.get(sh.shape_id, 0); self.grp[sh.shape_id] = g + 1
        if sh._element.tag == P + "sp": self.bld.append(f'<p:bldP spid="{sh.shape_id}" grpId="{g}" animBg="1"/>')
        return (f'<p:par><p:cTn id="{self.nid()}" presetID="{pid}" presetClass="{cls}" presetSubtype="{sub}" fill="hold" grpId="{g}" nodeType="{node}">'
                f'<p:stCondLst><p:cond delay="{delay}"/></p:stCondLst><p:childTnLst>')
    def wipe(self, sh, node, delay, dur):
        """왼쪽에서 오른쪽으로 닦아내기(차오르기)"""
        spid = sh.shape_id
        return (self._head(sh, 22, "entr", 8, node, delay) +
                f'<p:set><p:cBhvr><p:cTn id="{self.nid()}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>'
                f'<p:animEffect transition="in" filter="wipe(left)"><p:cBhvr><p:cTn id="{self.nid()}" dur="{dur}"/><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl></p:cBhvr></p:animEffect></p:childTnLst></p:cTn></p:par>')
    def scale(self, sh, node, pct, dur=250):
        """크게/작게 (튀어나오기)"""
        return (self._head(sh, 6, "emph", 0, node, 0) +
                f'<p:animScale><p:cBhvr><p:cTn id="{self.nid()}" dur="{dur}" fill="hold"/><p:tgtEl><p:spTgt spid="{sh.shape_id}"/></p:tgtEl></p:cBhvr><p:by x="{round(pct * 1000)}" y="{round(pct * 1000)}"/></p:animScale></p:childTnLst></p:cTn></p:par>')
    def color(self, sh, node, rgb, dur=250):
        """글자 색 바꾸기"""
        return (self._head(sh, 3, "emph", 2, node, 0) +
                f'<p:animClr clrSpc="rgb" dir="cw"><p:cBhvr override="childStyle"><p:cTn id="{self.nid()}" dur="{dur}" fill="hold"/><p:tgtEl><p:spTgt spid="{sh.shape_id}"/></p:tgtEl><p:attrNameLst><p:attrName>style.color</p:attrName></p:attrNameLst></p:cBhvr><p:to><a:srgbClr val="{rgb}"/></p:to></p:animClr></p:childTnLst></p:cTn></p:par>')
    def step(self, effs, auto=False):
        cond = '<p:cond delay="indefinite"/>' + ('<p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond>' if auto else "")
        self.pars.append(f'<p:par><p:cTn id="{self.nid()}" fill="hold"><p:stCondLst>{cond}</p:stCondLst><p:childTnLst><p:par><p:cTn id="{self.nid()}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>{"".join(effs)}</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>')
    def apply(self, slide):
        xml = ('<p:timing xmlns:p="%s" xmlns:a="%s"><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>'
               '<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>' % (PN, A) + "".join(self.pars) +
               '</p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
               '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>'
               '</p:childTnLst></p:cTn></p:par></p:tnLst><p:bldLst>' + "".join(self.bld) + '</p:bldLst></p:timing>')
        sld = slide._element
        for old in sld.findall(P + "timing"): sld.remove(old)
        new = etree.fromstring(xml); ext = sld.find(P + "extLst")
        if ext is not None: ext.addprevious(new)
        else: sld.append(new)


# ============ 대표사례(기존 8쪽): 주황 막대가 왼쪽에서 차오름 ============
s = S[8]; N = named(s)
banks = [("하나", 70), ("신한", 45), ("SC제일", 45), ("국민", 40), ("농협", 35)]
q = Seq(); effs = []
for k, (bank, v) in enumerate(banks):  # 같은 속도로 차오르도록 길이에 비례한 시간, 위에서부터 0.12초 간격
    effs.append(q.wipe(N[f"{bank} 부담"], "afterEffect" if k == 0 else "withEffect", 300 + k * 120, round(1000 * v / 70)))
q.step(effs, auto=True); q.apply(s)
note(s, "[변경 7차] 슬라이드에 들어오면 주황 막대(소비자 부담 비율)가 왼쪽에서 차오름. 다섯 막대가 같은 속도로, 위에서부터 0.12초 간격으로 출발.")

# ============ 개선방안(기존 14쪽): 말하는 주체만 주황으로 커졌다가, 다음으로 넘어가면 돌아옴 ============
s = S[14]
labels = sorted([sh for sh in s.shapes if sh.has_text_frame and sh.text_frame.text.strip() and round(PX(sh.left)) == 128 and 200 < PX(sh.top) < 900], key=lambda sh: sh.top)
assert [l.text_frame.text for l in labels] == ["정부·금융당국", "금융기관", "상품 설계 기업", "소비자"], [l.text_frame.text for l in labels]
q = Seq(); UP = 112; DOWN = 100 * 100 / UP
for k, lab in enumerate(labels):
    effs = [q.color(lab, "clickEffect", ACC), q.scale(lab, "withEffect", UP)]
    if k: effs += [q.color(labels[k - 1], "withEffect", INK), q.scale(labels[k - 1], "withEffect", DOWN)]
    q.step(effs)
q.apply(s)
note(s, "[변경 7차] 클릭할 때마다 지금 말하는 주체(정부·금융당국 → 금융기관 → 상품 설계 기업 → 소비자)의 이름이 주황으로 바뀌며 살짝 커지고, 앞의 주체는 원래 색·크기로 돌아옴. 클릭 4번.")

# ============ 5쪽: 열 하나씩 집중하는 장 3개를 뒤에 추가 (모핑 전환) ============
base = S[5]
COLX = [128, 700, 1272]; COLW = 520; TOP = 250
col_of = lambda sh: 0 if PX(sh.left) < 600 else (1 if PX(sh.left) < 1200 else 2)
body = [sh for sh in base.shapes if 200 < PX(sh.top) < 980]          # 제목·참고문헌 제외
for sh in body: sh.name = "!!열%d %s" % (col_of(sh) + 1, sh.name)     # 모핑이 같은 도형으로 짝짓도록 이름을 고정
morph_src = [c for c in S[7]._element if c.tag.endswith("AlternateContent")][0]
full_note = base.notes_slide.notes_text_frame.text
m = re.search(r"\[정확한 표현\](.*)", full_note)
items = re.findall(r"(\d\) .*?)(?=\s\d\) |$)", m.group(1).strip()) if m else []
HEAD = ["권리 강화", "분쟁 시 권익 확보", "처벌 강화"]; ITEMS = [(0, 2), (2, 5), (5, 6)]
SF, SU, Y0, GAP_ALL = 1.25, 0.72, 220, 1664


def recolor(sh, text_rgb, line_rgb):
    for el in sh._element.iter("{%s}srgbClr" % A):
        par = el.getparent().getparent().tag
        if par.endswith("}ln"): el.set("val", line_rgb)
        elif par.endswith("}rPr") or par.endswith("}endParaRPr") or par.endswith("}defRPr"): el.set("val", text_rgb)


def focus_slide(f, after):
    sl = prs.slides.add_slide(base.slide_layout)
    tree = sl.shapes._spTree; src_tree = base.shapes._spTree
    for el in list(tree)[2:]: tree.remove(el)                      # 레이아웃에서 온 자리표시자 제거
    for el in list(src_tree)[2:]: tree.append(copy.deepcopy(el))   # 기준 장의 도형을 그대로 복사
    bg = base._element.find(P + "cSld").find(P + "bg")
    if bg is not None: tree.addprevious(copy.deepcopy(bg))
    tr = copy.deepcopy(morph_src)
    for t_ in tr.iter(P + "transition"):
        t_.set("spd", "med")
        for k in list(t_.attrib):
            if k.endswith("}dur"): t_.set(k, "700")
    sl._element.find(P + "clrMapOvr").addnext(tr)
    scales = [SF if c == f else SU for c in range(3)]
    gap = (GAP_ALL - sum(COLW * x for x in scales)) / 2
    xs = [128]
    for c in (0, 1): xs.append(xs[-1] + COLW * scales[c] + gap)
    for sh in sl.shapes:
        if not sh.name.startswith("!!열"): continue
        c = col_of_name(sh.name); k = scales[c]
        x, y, w, h = PX(sh.left), PX(sh.top), PX(sh.width), PX(sh.height)
        sh.left, sh.top, sh.width, sh.height = E(xs[c] + (x - COLX[c]) * k), E(Y0 + (y - TOP) * k), E(w * k), E(h * k)
        for el in sh._element.iter("{%s}rPr" % A, "{%s}endParaRPr" % A):
            if el.get("sz"): el.set("sz", str(int(round(int(el.get("sz")) * k / 50.0)) * 50))
        if c != f: recolor(sh, MID, SOFT)
    lst = prs.slides._sldIdLst; ids = list(lst)
    me = ids[-1]; lst.remove(me); lst.insert(list(lst).index(after) + 1, me)
    ns = sl.notes_slide._element  # 새 노트에는 글상자가 없으므로 기준 장의 노트 틀을 복사
    ns.replace(ns.find(P + "cSld"), copy.deepcopy(base.notes_slide._element.find(P + "cSld")))
    a, b = ITEMS[f]
    note(sl, f"[화면] 앞 장에서 클릭하면 ‘{HEAD[f]}’ 열이 커지고 나머지 두 열은 작아지며 옅은 회색이 됨(모핑 전환). 이 열만 설명함.\n\n[정확한 표현] " + " ".join(items[a:b]) +
             "\n\n모핑은 PowerPoint 2019 이상·Microsoft 365에서 동작하며, 구버전에서는 서서히 바뀌는 전환으로 보임.", replace=True)
    return me


col_of_name = lambda nm: int(nm[3]) - 1
after = [i for i in prs.slides._sldIdLst if prs.slides.get(int(i.get("id"))) is base][0]
for f in range(3): after = focus_slide(f, after)
note(base, "[변경 7차] 이 장은 세 묶음을 한눈에 보여 주는 전체 보기. 클릭할 때마다 다음 장으로 넘어가며 권리 강화 → 분쟁 시 권익 확보 → 처벌 강화 순으로 한 열씩 커지고 나머지는 옅어짐(3장 추가, 이후 쪽 번호가 3씩 밀림).")

prs.save(DST)
tmp = DST + ".tmp"; zin = zipfile.ZipFile(DST); zout = zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED)
for it in zin.infolist():
    d = zin.read(it.filename)
    if it.filename == "[Content_Types].xml" and b'Extension="jpg"' not in d:
        d = d.replace(b"<Default ", b'<Default Extension="jpg" ContentType="image/jpeg"/><Default ', 1)
    zout.writestr(it, d)
zout.close(); zin.close(); os.replace(tmp, DST)
print("saved", DST, "| slides:", len(prs.slides), "| note items:", len(items))
