"""11차 수정: 13쪽(2년 10개월)의 주황 구간선을 없애고, 연표선이 클릭마다 주황으로 차오르게 한다.
사용: python3 revise11.py <수정10.pptx> <새 파일.pptx>"""
import copy, sys, os, zipfile
from lxml import etree
from pptx import Presentation
from pptx.util import Emu

SRC, DST = sys.argv[1], sys.argv[2]
E = lambda px: Emu(int(round(px * 6350)))
PX = lambda emu: emu / 6350
ACC = "EE4E22"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
PN = "http://schemas.openxmlformats.org/presentationml/2006/main"
P = "{%s}" % PN
prs = Presentation(SRC)
s = prs.slides[12]
N = {x.name: x for x in s.shapes}
assert N["s text 139"].text_frame.text == "2년 10개월"


class Seq:
    def __init__(self): self.i = 2; self.pars = []; self.bld = []; self.grp = {}
    def nid(self): self.i += 1; return self.i
    def entr(self, sh, node, delay=0, wipe=False, dur=450):
        spid = sh.shape_id; g = self.grp.get(spid, 0); self.grp[spid] = g + 1
        self.bld.append(f'<p:bldP spid="{spid}" grpId="{g}" animBg="1"/>')
        pid, sub, flt = (22, 8, "wipe(left)") if wipe else (10, 0, "fade")
        return (f'<p:par><p:cTn id="{self.nid()}" presetID="{pid}" presetClass="entr" presetSubtype="{sub}" fill="hold" grpId="{g}" nodeType="{node}"><p:stCondLst><p:cond delay="{delay}"/></p:stCondLst><p:childTnLst>'
                f'<p:set><p:cBhvr><p:cTn id="{self.nid()}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>'
                f'<p:animEffect transition="in" filter="{flt}"><p:cBhvr><p:cTn id="{self.nid()}" dur="{dur}"/><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl></p:cBhvr></p:animEffect></p:childTnLst></p:cTn></p:par>')
    def click(self, items):
        """items: [(도형, 지연 ms, 닦아내기 여부)]"""
        effs = "".join(self.entr(sh, "clickEffect" if k == 0 else "withEffect", d, w) for k, (sh, d, w) in enumerate(items))
        self.pars.append(f'<p:par><p:cTn id="{self.nid()}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="{self.nid()}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>{effs}</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>')
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


# 1) 주황 구간선(가로선 + 양 끝 눈금) 삭제
for nm in ("s line 140", "s line 141", "s line 142"): N[nm]._element.getparent().remove(N[nm]._element)

# 2) 구간선이 빠진 만큼 연표 전체를 50px 올려 큰 숫자와의 간격을 맞춤
UP = 50
timeline = [sh for sh in s.shapes if 590 < PX(sh.top) < 830]
for sh in timeline: sh.top = sh.top - E(UP)

# 3) 회색 연표선 위에 주황 선 네 토막을 얹음 (마디와 마디 사이). 원은 선보다 위에 있어 선을 가림
base = N["s line 143"]
nodes = [N[n] for n in ("circle 144", "circle 147", "circle 150", "circle 153", "circle 156")]
cx = [PX(c.left + c.width / 2) for c in nodes]
next_id = max(x.shape_id for x in s.shapes) + 1; segs = []; prev = base
for k in range(4):
    el = copy.deepcopy(base._element); prev._element.addnext(el)
    el.find(P + "nvSpPr").find(P + "cNvPr").set("id", str(next_id)); next_id += 1
    sp = [x for x in s.shapes if x._element is el][0]; sp.name = f"주황 구간 {k + 1}"
    sp.left, sp.width = E(cx[k]), E(cx[k + 1] - cx[k])
    for c in el.iter("{%s}srgbClr" % A): c.set("val", ACC)
    segs.append(sp); prev = sp

# 4) 클릭 순서는 그대로, 마디가 나타나기 전에 그 앞 구간이 왼쪽에서 차오름
grp = lambda *names: [N[n] for n in names]
g1 = grp("circle 144", "text 145", "text 146"); g2 = grp("circle 147", "text 148", "text 149"); g3 = grp("circle 150", "text 151", "text 152")
g4 = grp("circle 153", "text 154", "text 155"); g5 = grp("circle 156", "text 157", "text 158")
W = 450  # 한 구간이 차오르는 시간(ms)
q = Seq()
q.click([(sh, 0, False) for sh in g1])
q.click([(segs[0], 0, True)] + [(sh, W - 100, False) for sh in g2] + [(segs[1], W, True)] + [(sh, 2 * W - 100, False) for sh in g3])
q.click([(segs[2], 0, True)] + [(sh, W - 100, False) for sh in g4])
q.click([(segs[3], 0, True)] + [(sh, W - 100, False) for sh in g5])
q.apply(s)
tf = s.notes_slide.notes_text_frame
tf.text = tf.text.rstrip() + ("\n\n[변경 11차] ‘2년 10개월’ 아래의 주황 구간선을 없애고 연표선 하나로 합침. 클릭할 때마다 지나온 구간이 왼쪽에서 주황으로 차오른 뒤 그 마디가 나타남. "
                              "클릭 1: 2024.1 → 클릭 2: 2024.3, 2024.5까지 → 클릭 3: 2025.2까지 → 클릭 4: 2026.11.4까지 차며 선 전체가 주황이 됨. 연표는 50px 위로 올림.")

prs.save(DST)
tmp = DST + ".tmp"; zin = zipfile.ZipFile(DST); zout = zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED)
for it in zin.infolist():
    d = zin.read(it.filename)
    if it.filename == "[Content_Types].xml" and b'Extension="jpg"' not in d:
        d = d.replace(b"<Default ", b'<Default Extension="jpg" ContentType="image/jpeg"/><Default ', 1)
    zout.writestr(it, d)
zout.close(); zin.close(); os.replace(tmp, DST)
print("saved", DST, "| node x:", cx, "| moved:", len(timeline))
