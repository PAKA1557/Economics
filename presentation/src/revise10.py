"""10차 수정: 오른쪽 위에 '소비자의 권리 · 1조' 표기를 넣는다(표지·13쪽·14쪽·참고문헌 제외).
사용: python3 revise10.py <수정9.pptx> <새 파일.pptx>"""
import copy, sys, os, zipfile
from pptx import Presentation
from pptx.util import Emu

SRC, DST = sys.argv[1], sys.argv[2]
E = lambda px: Emu(int(round(px * 6350)))
P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
LABEL = "소비자의 권리 · 1조"
SKIP = {1, 13, 14}
prs = Presentation(SRC)
done = []; tmpl = None
for n, s in enumerate(prs.slides, 1):
    texts = [sh for sh in s.shapes if sh.has_text_frame]
    title = texts[0].text_frame.text if texts else ""
    refs = [sh for sh in texts if sh.text_frame.text.startswith("참고문헌 [")]
    if n in SKIP or title.startswith("참고문헌"): continue
    if refs: tmpl = refs[0]._element                # 참고문헌 표기와 같은 서식(작은 회색, 오른쪽 정렬). 없는 장은 앞 장 것을 씀
    el = copy.deepcopy(tmpl); s.shapes._spTree.append(el)
    el.find(P + "nvSpPr").find(P + "cNvPr").set("id", str(max(x.shape_id for x in s.shapes) + 1))
    sp = [x for x in s.shapes if x._element is el][0]; sp.name = "!!머리 표기"
    sp.left, sp.top, sp.width, sp.height = E(1192), E(107), E(600), E(34)   # 제목과 같은 높이, 오른쪽 여백에 맞춤
    runs = sp.text_frame.paragraphs[0].runs; runs[0].text = LABEL
    for r in runs[1:]: r._r.getparent().remove(r._r)
    done.append(n)
prs.save(DST)
tmp = DST + ".tmp"; zin = zipfile.ZipFile(DST); zout = zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED)
for it in zin.infolist():
    d = zin.read(it.filename)
    if it.filename == "[Content_Types].xml" and b'Extension="jpg"' not in d:
        d = d.replace(b"<Default ", b'<Default Extension="jpg" ContentType="image/jpeg"/><Default ', 1)
    zout.writestr(it, d)
zout.close(); zin.close(); os.replace(tmp, DST)
print("saved", DST, "| labelled slides:", done)
