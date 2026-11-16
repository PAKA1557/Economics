"""5차 수정: '수정4' 복사본의 5·13·17쪽 배치만 고친다. 사용: python3 revise5.py <수정4.pptx> <새 파일.pptx>

 5쪽  '50%'를 왼쪽 아래 '1년 · 5년'과 같은 기준선에 놓고, 두 설명글을 같은 줄·같은 크기로 맞춘다.
13쪽  카드 안 글 묶음을 세로 가운데로 내리고, 본문 줄바꿈을 고정한다.
17쪽  문헌 6개를 한 글상자로 합쳐 줄 수와 상관없이 간격이 같게 한다.
"""
import copy, sys
from lxml import etree
from pptx import Presentation
from pptx.util import Emu, Pt

SRC, DST = sys.argv[1], sys.argv[2]
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
prs = Presentation(SRC)
S = {n + 1: s for n, s in enumerate(prs.slides)}
by_id = lambda slide: {sh.shape_id: sh for sh in slide.shapes}

# ── 5쪽 ──────────────────────────────────────────────────────────────
s = by_id(S[5])
term, term_cap, pct, pct_cap = s[26], s[27], s[21], s[22]   # 1년·5년 / 그 설명 / 50% / 그 설명
BASE = 0.436                                                 # 글상자 중심에서 기준선까지 거리(em), Noto Sans KR
em = lambda sh: sh.text_frame.paragraphs[0].runs[0].font.size
mid = term.top + term.height // 2
pct.top = Emu(int(mid + BASE * (em(term) - em(pct)) - pct.height // 2))
pct_cap.top, pct_cap.height = term_cap.top, term_cap.height
for r in pct_cap.text_frame.paragraphs[0].runs:
    r.font.size = em(term_cap)

# ── 13쪽 ─────────────────────────────────────────────────────────────
s = by_id(S[13])
SHIFT = 700000
for i in (4, 5, 6, 8, 9, 10, 12, 13, 14, 16, 17, 18):
    s[i].top = Emu(s[i].top + SHIFT)
for i in (6, 10, 14, 18):
    s[i].text_frame.word_wrap = False

# ── 17쪽 ─────────────────────────────────────────────────────────────
s = by_id(S[17])
first, rest = s[3], [s[i] for i in (4, 5, 6, 7, 8)]
body = first.text_frame._txBody
for sh in rest:
    for p in sh.text_frame._txBody.findall(A + "p"):
        body.append(copy.deepcopy(p))
    sh._element.getparent().remove(sh._element)
for p in first.text_frame.paragraphs:
    p.space_after = Pt(25)
first.height = Emu(rest[-1].top + rest[-1].height - first.top)

prs.save(DST)
