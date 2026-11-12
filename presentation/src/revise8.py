"""8차 수정: '수정7' 복사본의 개선방안(17쪽) 뒤에 줄 하나씩 집중하는 장 4개를 추가(모핑 전환).
사용: python3 revise8.py <수정7.pptx> <새 파일.pptx>"""
import copy, sys, os, zipfile
from pptx import Presentation
from pptx.util import Emu

SRC, DST = sys.argv[1], sys.argv[2]
E = lambda px: Emu(int(round(px * 6350)))
PX = lambda emu: emu / 6350
MID, SOFT = "B4BAC4", "D9DDE3"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
PN = "http://schemas.openxmlformats.org/presentationml/2006/main"
P = "{%s}" % PN
prs = Presentation(SRC)
S = {n + 1: s for n, s in enumerate(prs.slides)}
base = S[17]
assert base.shapes[0].text_frame.text == "개선방안"

TOP, RH, NROW = 230, 172, 4             # 첫 줄 위치, 줄 높이, 줄 수
HF = 1.6; HU = (NROW - HF) / (NROW - 1)  # 집중한 줄은 높이 1.6배, 나머지는 0.8배 (전체 높이는 그대로)
KF, KU = 1.2, 0.75                       # 글자 크기 배율
ACTORS = ["정부·금융당국", "금융기관", "상품 설계 기업", "소비자"]

# 줄별로 도형을 나누고, 모핑이 같은 도형끼리 짝짓도록 이름을 고정
for el in base._element.findall(P + "timing"): base._element.remove(el)  # 7차의 색·크기 강조는 이 방식으로 대체
kind = {}
for sh in base.shapes:
    y, h = PX(sh.top), PX(sh.height)
    if not (TOP - 1 <= y <= TOP + NROW * RH + 1): continue
    if h < 1:                                  # 가로 구분선: 몇 번째 경계인지
        sh.name = "!!선%d" % round((y - TOP) / RH)
    else:                                      # 글상자·점: 몇 번째 줄인지
        sh.name = "!!줄%d %s" % (int((y + h / 2 - TOP) // RH) + 1, sh.name)
morph_src = [c for c in S[10]._element if c.tag.endswith("AlternateContent")][0]
base_note = "\n\n".join(p for p in base.notes_slide.notes_text_frame.text.split("\n\n") if not p.startswith("[변경") and not p.startswith("[확인"))


def recolor(sh, rgb):
    for el in sh._element.iter("{%s}srgbClr" % A):
        if not el.getparent().getparent().tag.endswith("}ln"): el.set("val", rgb)


def focus_slide(f, after):
    sl = prs.slides.add_slide(base.slide_layout)
    tree = sl.shapes._spTree; src_tree = base.shapes._spTree
    for el in list(tree)[2:]: tree.remove(el)
    for el in list(src_tree)[2:]: tree.append(copy.deepcopy(el))
    bg = base._element.find(P + "cSld").find(P + "bg")
    if bg is not None: tree.addprevious(copy.deepcopy(bg))
    tr = copy.deepcopy(morph_src)
    for t_ in tr.iter(P + "transition"):
        t_.set("spd", "med")
        for k in list(t_.attrib):
            if k.endswith("}dur"): t_.set(k, "700")
    sl._element.find(P + "clrMapOvr").addnext(tr)
    hs = [RH * (HF if r == f else HU) for r in range(NROW)]
    tops = [TOP + sum(hs[:r]) for r in range(NROW + 1)]
    for sh in sl.shapes:
        nm = sh.name
        if nm.startswith("!!선"): sh.top = E(tops[int(nm[3])])
        elif nm.startswith("!!줄"):
            r = int(nm[3]) - 1; k = KF if r == f else KU
            if sh.has_text_frame and sh.text_frame.text.strip():
                sh.top, sh.height = E(tops[r]), E(hs[r])
                for el in sh._element.iter("{%s}rPr" % A, "{%s}endParaRPr" % A):
                    if el.get("sz"): el.set("sz", str(int(round(int(el.get("sz")) * k / 50.0)) * 50))
                if r != f: recolor(sh, MID)
            else:                              # 글머리 점: 가운데를 지키며 크기만 바꿈
                cx = PX(sh.left + sh.width / 2); d = PX(sh.width) * k
                sh.left, sh.top, sh.width, sh.height = E(cx - d / 2), E(tops[r] + hs[r] / 2 - d / 2), E(d), E(d)
                if r != f: recolor(sh, SOFT)
    lst = prs.slides._sldIdLst; me = list(lst)[-1]; lst.remove(me); lst.insert(list(lst).index(after) + 1, me)
    ns = sl.notes_slide._element
    ns.replace(ns.find(P + "cSld"), copy.deepcopy(base.notes_slide._element.find(P + "cSld")))
    sl.notes_slide.notes_text_frame.text = (f"[화면] 앞 장에서 클릭하면 ‘{ACTORS[f]}’ 줄이 커지고 나머지 줄은 작아지며 옅은 회색이 됨(모핑 전환). 이 줄만 설명함.\n\n" + base_note)
    return me


after = [i for i in prs.slides._sldIdLst if prs.slides.get(int(i.get("id"))) is base][0]
for f in range(NROW): after = focus_slide(f, after)
tf = base.notes_slide.notes_text_frame
tf.text = tf.text.rstrip() + ("\n\n[변경 8차] 이 장은 네 주체를 한눈에 보여 주는 전체 보기. 클릭할 때마다 다음 장으로 넘어가며 정부·금융당국 → 금융기관 → 상품 설계 기업 → 소비자 순으로 "
                              "한 줄씩 커지고 나머지는 작아지며 옅어짐(4장 추가, 이후 쪽 번호가 4씩 밀림). 7차의 이름 색·크기 강조는 이 방식으로 대체함.")

prs.save(DST)
tmp = DST + ".tmp"; zin = zipfile.ZipFile(DST); zout = zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED)
for it in zin.infolist():
    d = zin.read(it.filename)
    if it.filename == "[Content_Types].xml" and b'Extension="jpg"' not in d:
        d = d.replace(b"<Default ", b'<Default Extension="jpg" ContentType="image/jpeg"/><Default ', 1)
    zout.writestr(it, d)
zout.close(); zin.close(); os.replace(tmp, DST)
print("saved", DST, "| slides:", len(prs.slides))
