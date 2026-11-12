"""13차 수정: 발표자 노트에서 변경 이력·화면 순서·예상 질문을 걷어 내고 내용 주석만 남긴다. 예상 질문은 텍스트 파일로 따로 낸다.
사용: python3 revise13.py <수정12.pptx> <새 파일.pptx> <예상질문.txt>"""
import re, sys, os, zipfile
from pptx import Presentation

SRC, DST, QA = sys.argv[1], sys.argv[2], sys.argv[3]
prs = Presentation(SRC)
DROP = ("[변경", "[예상 질문]", "[화면 순서]", "[화면]", "[화면의 큰 숫자]", "모핑은 PowerPoint")
# 꼬리표나 낡은 표현을 지금 화면에 맞게 고침: {쪽: [(찾을 말, 바꿀 말)]}
FIX = {
    9: [("그 밖 0.5조원", "기타 0.5조원")],
    12: [("[새 장 3차]", "[핵심]")],
    14: [("[그 외 130억]", "[기타 130억]"), ("띠의 굵기는 금액에 비례함", "칸의 넓이는 금액에 비례함")],
    17: [("[확인 6차] ‘고난도’ 표기는 그대로 둠. ", "[용어] ")],
}
# 변경 이력에만 적혀 있던 내용 중 주석으로 남겨야 하는 것
ADD = {
    10: ["[핵심] 손실 4.6조원의 구성: 배상 1.3조원(평균 배상비율 31.4%), 소비자 부담 2.8조원, 기타 0.5조원. 한 칸은 약 0.1조원이며 칸 수는 반올림한 값(윤곽선 13칸 = 배상, 주황 27칸 = 소비자 부담, 진회색 4칸 = 기타).",
         "[팀 계산] 소비자 부담 2.8조원 = 배상 동의 계좌의 손실 4.1 - 배상 1.3. 기타 0.5조원 = 전체 손실 4.6 - 4.1.",
         "[주의] 31.4%는 배상 동의 15.9만 계좌 기준이며 전체 피해자에게 똑같이 적용된 확정 부담률이 아님."],
    11: ["[수치] 화면의 ‘소비자 부담 비율’은 100 - 배상비율로 계산한 값(팀 계산): 하나 70% / 신한 45% / SC제일 45% / 국민 40% / 농협 35%."],
}
TITLES = {}
qa = {}   # 제목 -> [(Q, A)]
for n, s in enumerate(prs.slides, 1):
    texts = [sh.text_frame.text for sh in s.shapes if sh.has_text_frame and sh.text_frame.text.strip()]
    TITLES[n] = texts[0].replace("\n", " ") if texts else ""
    if n == 14: TITLES[n] = "제재 금액"   # 이 장은 제목 글상자가 없음
    if not s.has_notes_slide: continue
    tf = s.notes_slide.notes_text_frame
    if tf is None: continue
    paras = [p.strip() for p in tf.text.split("\n\n") if p.strip()]
    keep = []
    for p in paras:
        if p.startswith("[예상 질문]"):
            for q, a in re.findall(r"Q\.\s*(.*?)\s*A\.\s*(.*?)(?=\s*Q\.|$)", p[len("[예상 질문]"):].strip(), flags=re.S):
                lst = qa.setdefault(n, [])
                if (q, a) not in [x for v in qa.values() for x in v]: lst.append((q, a))
        if p.startswith(DROP): continue
        for old, new in FIX.get(n, []): p = p.replace(old, new)
        keep.append(p)
    keep += ADD.get(n, [])
    if n in (18, 19, 20, 21): keep = [p for p in keep if not p.startswith("[핵심]")]   # 17쪽과 같은 문단의 반복은 뺌
    tf.text = "\n\n".join(keep)
    left = [t for t in ("[변경", "예상 질문", "화면 순서", "차]") if t in tf.text]
    assert not left, (n, left)

prs.save(DST)
tmp = DST + ".tmp"; zin = zipfile.ZipFile(DST); zout = zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED)
for it in zin.infolist():
    d = zin.read(it.filename)
    if it.filename == "[Content_Types].xml" and b'Extension="jpg"' not in d:
        d = d.replace(b"<Default ", b'<Default Extension="jpg" ContentType="image/jpeg"/><Default ', 1)
    zout.writestr(it, d)
zout.close(); zin.close(); os.replace(tmp, DST)

# ---------------- 예상 질문 파일 ----------------
# 노트의 다른 문단에 '질문이 나오면 이렇게 답한다'로 적혀 있던 것들
EXTRA = {
    2: [("DLF와 ELS의 손실 규모를 그대로 비교해도 되나?", "집계 범위가 다름. DLF는 전체 투자자 기준이고, ELS는 5대 판매은행의 손실확정 계좌 기준임. 원의 크기는 손실액 4.6조원과 4,468억원의 비(약 10.3배, 팀 계산)를 보여 주는 것임.")],
    9: [("‘기타 0.5조원’은 무엇인가?", "전체 손실 4.6조원에서 배상 동의 계좌의 손실 4.1조원을 뺀 값(팀 계산)임. 어떤 계좌인지는 원자료로 확인하지 못했으므로 ‘배상 동의가 끝나지 않은 계좌의 손실’ 정도로만 답함."),
        ("평균 배상비율 31.4%면 소비자가 68.6%를 부담했다는 뜻인가?", "그렇게 단정하지 않음. 31.4%는 배상 동의 15.9만 계좌(배상 진행 16.9만 건의 93.8%) 기준이며, 전체 피해자에게 똑같이 적용된 확정 부담률이 아님."),
        ("위법계약해지권을 쓰면 손실을 돌려받을 수 있지 않나?", "해지 효과는 장래를 향해 발생함. 만기에 이미 확정된 손실을 돌려받는 장치가 아니며 손해배상청구와 구별됨(법 제47조, 시행령 제38조). ‘ELS에는 적용되지 않았다’고 단정하지는 않고 구조적 한계까지만 말함.")],
    11: [("대표사례는 금소법 시행 후 판매된 것인가?", "아님. 5건 모두 2021.3.24 이전 판매분임. 또 배상비율은 판매사 책임에 투자자의 경험·가입 목적 등 개별 사정을 반영한 것이어서, 낮은 배상비율만으로 보호 실패가 입증되는 것은 아님.")],
    17: [("‘고난도’ 표시는 왜 ‘고위험’이 아닌가?", "‘고난도금융투자상품’이 자본시장법 시행령 제2조의 법령 용어이기 때문임. 최대 원금손실 가능금액이 원금의 20%를 넘는 상품을 뜻하며 2021.5.10 시행.")],
    22: [("결국 하고 싶은 말이 무엇인가?", "금소법의 목적은 ‘금융소비자의 권익 증진’(제1조)임. 홍콩 ELS 사태에서 그 목적이 실제로 달성됐는지 확인했고, 권리는 있었지만 충분히 작동하지 않았으므로 실질적 증진을 위한 개선방안을 제시함.")],
}
for n, items in EXTRA.items(): qa.setdefault(n, []).extend(items)
CLICKS = [
    (2, "처음에는 2019년 DLF만 보임 → 클릭 1: 2021.3 금소법 → 클릭 2: 2024 ELS"),
    (3, "은행 창구 → 클릭 1: 정기예금 갈래 → 클릭 2: ELS 갈래"),
    (4, "들어오면 도넛이 차오르며 숫자가 44.2%까지 올라감(자동) → 클릭 1: 17.0만 계좌 → 클릭 2: 10.4조원 → 클릭 3: 4.6조원"),
    (5, "전체 보기 → 클릭 1(6쪽): 권리 강화 확대 → 클릭 2(7쪽): 분쟁 시 권익 확보 확대 → 클릭 3(8쪽): 처벌 강화 확대"),
    (9, "원금 격자(손실 44칸, 남은 원금 56칸) → 클릭 1(10쪽): 손실 부분이 가로로 늘어남 → 클릭 2: 배상·소비자 부담·기타로 갈림"),
    (11, "들어오면 주황 막대가 왼쪽에서 차오름(자동). 클릭 없음"),
    (12, "시행 전 사례 여섯 개 → 클릭 1: 주황 시행선과 ‘2021.3 | 시행 후’ → 클릭 2: 시행 전 사례가 옅어지고 2024 홍콩 ELS 원이 나타남"),
    (13, "큰 숫자와 빈 선 → 클릭 1: 2024.1 → 클릭 2: 2024.3, 2024.5 → 클릭 3: 2025.2 → 클릭 4: 2026.11.4 (구간이 주황으로 차오름)"),
    (16, "1단계만 보임 → 클릭 1: 2단계 → 클릭 2: 3단계 → 클릭 3: 결과(손실)"),
    (17, "전체 보기 → 클릭 1(18쪽): 정부·금융당국 → 클릭 2(19쪽): 금융기관 → 클릭 3(20쪽): 상품 설계 기업 → 클릭 4(21쪽): 소비자"),
    (22, "제도와 현실이 점선으로 떨어져 있음 → 클릭 1: 실선으로 이어지며 과제 두 가지가 나타남"),
]
out = ["홍콩 ELS 사태로 알아보는 소비자 권리 — 예상 질문과 답변", "소비자의 권리 · 1조", "",
       "쪽 번호는 발표 자료(24장) 기준입니다. 답변은 발표자 노트에 있던 내용을 옮긴 것이며, 근거는 각 장 아래의 참고문헌 번호를 따릅니다.", ""]
total = 0
for n in sorted(qa):
    if not qa[n]: continue
    out += ["=" * 40, f"{n}쪽  {TITLES[n]}", "=" * 40, ""]
    for q, a in qa[n]:
        total += 1; out += [f"Q{total}. {q}", f"A. {a}", ""]
out += ["=" * 40, "부록: 장별 클릭 순서", "=" * 40, ""] + [f"{n}쪽  {TITLES[n]}: {c}" for n, c in CLICKS] + [""]
open(QA, "w", encoding="utf-8").write("\n".join(out))
print("saved", DST, "|", QA, "| questions:", total)
