"""업로드본(24쪽, Pretendard) 1차 수정. 사용: python3 revise_upload1.py <업로드본.pptx> <새 파일.pptx>

 1쪽  밑줄과 과목 표기를 제목 쪽으로 올린다.
10쪽  손실 44칸의 구성을 13/27/4칸에서 12/27/5칸으로 바로잡는다(한 칸 약 0.1조원).
12쪽  수치 설명을 13pt에서 14pt로 키운다.
13쪽  오른쪽 위 머리 표기를 넣고, 내용을 조금 내린다.
16쪽  카드 안 글 묶음을 세로 가운데로 내리고 본문 줄바꿈을 고정한다.
18~21쪽  네 장에 똑같이 들어 있던 노트를 해당 행 내용만 남긴다.
23쪽  [3]·[5]·[8]·[10]에도 「」를 넣어 표기를 맞춘다.
"""
import copy, sys
from pptx import Presentation
from pptx.util import Emu, Pt

SRC, DST = sys.argv[1], sys.argv[2]
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
prs = Presentation(SRC)
S = {n + 1: s for n, s in enumerate(prs.slides)}
by_id = lambda shapes: {sh.shape_id: sh for sh in shapes}
down = lambda sh, d: setattr(sh, "top", Emu(sh.top + d))

# ── 1쪽 ──────────────────────────────────────────────────────────────
s = by_id(S[1].shapes)
down(s[4], -620000); down(s[5], -620000)

# ── 10쪽 ─────────────────────────────────────────────────────────────
grp = by_id(S[10].shapes)[3]
cell = by_id(grp.shapes)
def restyle(dst, src, name, left, top):
    """dst 칸을 src 칸과 같은 모양으로 바꿔 (left, top)에 놓는다."""
    old = dst._element.find(P + "spPr")
    old.getparent().replace(old, copy.deepcopy(src._element.find(P + "spPr")))
    dst.name, dst.left, dst.top = name, left, top
restyle(cell[150], cell[151], "부담 칸", cell[151].left, cell[154].top)   # 4열 1행: 배상 → 소비자 부담
restyle(cell[177], cell[178], "그 밖 칸", cell[177].left, cell[177].top)  # 10열 4행: 소비자 부담 → 기타

# ── 12쪽 ─────────────────────────────────────────────────────────────
for sh in S[12].shapes:
    if sh.name == "수치 설명":
        mid = sh.left + sh.width // 2
        sh.width = Emu(1500000); sh.left = Emu(mid - 750000)
        for p in sh.text_frame.paragraphs:
            for r in p.runs: r.font.size = Pt(14)

# ── 13쪽 ─────────────────────────────────────────────────────────────
head = next(sh for sh in S[16].shapes if sh.name == "!!머리 표기")
for sh in S[13].shapes:
    if sh.shape_id != 23: down(sh, 380000)                 # 23 = 참고문헌 꼬리표
el = copy.deepcopy(head._element)
el.find(P + "nvSpPr").find(P + "cNvPr").set("id", str(max(sh.shape_id for sh in S[13].shapes) + 1))
S[13].shapes._spTree.append(el)

# ── 16쪽 ─────────────────────────────────────────────────────────────
s = by_id(S[16].shapes)
for i in (4, 5, 6, 8, 9, 10, 12, 13, 14, 16, 17, 18): down(s[i], 700000)
for i in (6, 10, 14, 18): s[i].text_frame.word_wrap = False

# ── 18~21쪽 노트 ─────────────────────────────────────────────────────
NOTES = {
    18: "[항목별 상태] 최선이익 원칙을 담은 금소법 개정안은 국회 계류 중(2026.11.4 금융위 발표).\n\n"
        "[슬라이드에서 뺀 항목] 투자성 상품 설명 가이드라인(2027년 1분기 시행 예정).\n\n"
        "[강의 연결] 금융감독의 목적은 금융소비자 보호(6장 4절).",
    19: "[항목별 상태] 거점점포·전담직원 판매, 고객 손익을 반영한 KPI는 2025.2 대책의 과제이며 모든 기관의 이행 완료 여부까지 확인한 것은 아님. "
        "전담직원 경력 요건(3년 이상)은 후속 규정 확인이 필요함.\n\n"
        "[슬라이드에서 뺀 항목] 책무구조도로 경영진 책임 명확화.",
    20: "[항목별 상태] 상품명 표시는 2025.2 대책의 과제이며 모든 기관의 이행 완료 여부까지 확인한 것은 아님.\n\n"
        "[슬라이드에서 뺀 항목] 부적합 소비자 유형 기재, 전문용어 대신 쉬운 표현.",
    21: "[슬라이드에서 뺀 항목] 가족 지정인 확인 서비스(65세 이상 중 희망자 대상).\n\n"
        "[강의 연결] 소비자의 책무(8장 1절 5).",
}
for n, t in NOTES.items():
    S[n].notes_slide.notes_text_frame.text = t

# ── 23쪽 ─────────────────────────────────────────────────────────────
FIX = {
    "홍콩 H지수 ELS 대표사례 분쟁조정 결과(보도자료)": "「홍콩 H지수 ELS 대표사례 분쟁조정 결과」(보도자료)",
    "위법계약해지권 안내.": "「위법계약해지권 안내」.",
    "복잡한 구조채 판매에 관한 자율규제 개정,": "「복잡한 구조채 판매에 관한 자율규제 개정」,",
    "홍콩 H지수 ELS 현장검사 결과,": "「홍콩 H지수 ELS 현장검사 결과」,",
}
done = 0
for sh in S[23].shapes:
    if not sh.has_text_frame: continue
    for p in sh.text_frame.paragraphs:
        for r in p.runs:
            for a, b in FIX.items():
                if a in r.text: r.text = r.text.replace(a, b); done += 1
assert done == len(FIX), done

prs.save(DST)
