"""12차 수정: 4쪽 수치를 '강조 위, 설명 아래'로 통일하고, 12쪽(유사 사례)에 시행선·홍콩 ELS 등장 효과를 넣는다.
사용: python3 revise12.py <수정11.pptx> <새 파일.pptx>"""
import sys, os, zipfile
from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Emu

SRC, DST = sys.argv[1], sys.argv[2]
E = lambda px: Emu(int(round(px * 6350)))
PX = lambda emu: emu / 6350
BG = "EFF0F2"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
PN = "http://schemas.openxmlformats.org/presentationml/2006/main"
P = "{%s}" % PN
prs = Presentation(SRC)


def note(slide, text):
    tf = slide.notes_slide.notes_text_frame; tf.text = tf.text.rstrip() + "\n\n" + text


# ============ 4쪽: 큰 숫자를 위로, 옅은 설명을 아래로 (구분선·클릭 순서는 그대로) ============
s = prs.slides[3]; N = {x.name: x for x in s.shapes}
assert N["text 48"].text_frame.text.startswith("17.0")
for k, (lab, num) in enumerate((("text 47", "text 48"), ("text 50", "text 51"), ("text 53", "text 54"))):
    base = 250 + 210 * k
    N[num].top = E(base - 6)      # 숫자(높이 124)
    N[lab].top = E(base + 122)    # 설명(높이 46) — 그 아래 20px에 기존 구분선
note(s, "[변경 12차] 오른쪽 수치를 다른 장과 같은 순서로 통일: 큰 숫자가 위, 옅은 설명이 아래(예: 17.0만 계좌 / 손실이 확정된 계좌). 구분선 위치와 클릭 순서는 그대로.")

# ============ 12쪽: 유사 사례 — 시행선이 그어지고, 그 뒤에 홍콩 ELS가 나타남 ============
s = prs.slides[11]; N = {x.name: x for x in s.shapes}
assert N["s text 106"].text_frame.text == "유사 사례"
vline, after, when = N["금소법 시행선"], N["시행 후"], N["시행 시점"]
circle = N["ELS 원"]; els_text = [N[n] for n in ("ELS 연도", "ELS 이름", "ELS 수치", "ELS 설명")]
# 시행 전 사례를 옅게 만드는 가림막: 배경색 반투명 사각형. 시행선과 '2021.3'은 덮지 않음
veil = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, E(60), E(300), E(PX(vline.left) - 12 - 60), E(500)); veil.name = "시행 전 옅게(가림막)"
veil.shadow.inherit = False; veil.line.fill.background(); veil.fill.solid(); veil.fill.fore_color.rgb = RGBColor.from_string(BG)
clr = veil._element.spPr.find("{%s}solidFill" % A).find("{%s}srgbClr" % A)
etree.SubElement(clr, "{%s}alpha" % A).set("val", "62000")
circle._element.addprevious(veil._element)   # 시행 전 사례보다 위, 홍콩 ELS보다 아래

i = [2]
def nid(): i[0] += 1; return i[0]
bld = []
def entr(sh, node, kind, delay=0, dur=450):
    """kind: fade(서서히) / down(위에서 아래로 긋기) / zoom(커지며 나타나기)"""
    spid = sh.shape_id; bld.append(f'<p:bldP spid="{spid}" grpId="0" animBg="1"/>')
    tgt = f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>'
    vis = (f'<p:set><p:cBhvr><p:cTn id="{nid()}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>{tgt}<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>')
    if kind == "zoom":
        pid, sub = 53, 16
        grow = "".join(f'<p:anim calcmode="lin" valueType="num"><p:cBhvr><p:cTn id="{nid()}" dur="{dur}" fill="hold"/>{tgt}<p:attrNameLst><p:attrName>{a}</p:attrName></p:attrNameLst></p:cBhvr>'
                       f'<p:tavLst><p:tav tm="0"><p:val><p:fltVal val="0"/></p:val></p:tav><p:tav tm="100000"><p:val><p:strVal val="#{a}"/></p:val></p:tav></p:tavLst></p:anim>' for a in ("ppt_w", "ppt_h"))
        body = grow + f'<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="{nid()}" dur="{dur}"/>{tgt}</p:cBhvr></p:animEffect>'
    else:
        pid, sub, flt = (22, 1, "wipe(up)") if kind == "down" else (10, 0, "fade")
        body = f'<p:animEffect transition="in" filter="{flt}"><p:cBhvr><p:cTn id="{nid()}" dur="{dur}"/>{tgt}</p:cBhvr></p:animEffect>'
    return (f'<p:par><p:cTn id="{nid()}" presetID="{pid}" presetClass="entr" presetSubtype="{sub}" fill="hold" grpId="0" nodeType="{node}"><p:stCondLst><p:cond delay="{delay}"/></p:stCondLst><p:childTnLst>{vis}{body}</p:childTnLst></p:cTn></p:par>')
def click(effs):
    return (f'<p:par><p:cTn id="{nid()}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="{nid()}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>{"".join(effs)}</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>')
pars = [
    # 클릭 1: 주황 시행선이 위에서 아래로 그어지고 '2021.3 | 시행 후'가 나타남
    click([entr(vline, "clickEffect", "down", dur=600), entr(when, "withEffect", "fade", 150), entr(after, "withEffect", "fade", 350)]),
    # 클릭 2: 시행 전 사례가 옅어지고, 홍콩 ELS 원이 커지며 나타남
    click([entr(veil, "clickEffect", "fade", dur=500), entr(circle, "withEffect", "zoom", 150, 500)] + [entr(t, "withEffect", "fade", 450) for t in els_text]),
]
xml = ('<p:timing xmlns:p="%s"><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>'
       '<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>' % PN + "".join(pars) +
       '</p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
       '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>'
       '</p:childTnLst></p:cTn></p:par></p:tnLst><p:bldLst>' + "".join(bld) + '</p:bldLst></p:timing>')
for old in s._element.findall(P + "timing"): s._element.remove(old)
ext = s._element.find(P + "extLst"); new = etree.fromstring(xml)
if ext is not None: ext.addprevious(new)
else: s._element.append(new)
note(s, "[변경 12차] 효과 추가. 들어오면 시행 전 사례 여섯 개와 가로선만 보임 → 클릭 1: 주황 시행선이 위에서 아래로 그어지고 ‘2021.3 | 시행 후’가 나타남 → "
        "클릭 2: 시행 전 사례가 옅어지고 ‘2024 홍콩 ELS 4.6조원’ 원이 커지며 나타남. 편집 화면에서 시행 전 사례가 옅게 보이는 것은 그 위에 놓인 반투명 가림막(‘시행 전 옅게’) 때문이며 정상.")

prs.save(DST)
tmp = DST + ".tmp"; zin = zipfile.ZipFile(DST); zout = zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED)
for it in zin.infolist():
    d = zin.read(it.filename)
    if it.filename == "[Content_Types].xml" and b'Extension="jpg"' not in d:
        d = d.replace(b"<Default ", b'<Default Extension="jpg" ContentType="image/jpeg"/><Default ', 1)
    zout.writestr(it, d)
zout.close(); zin.close(); os.replace(tmp, DST)
print("saved", DST)
