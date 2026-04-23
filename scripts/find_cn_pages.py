"""원본 HWPX에서 CN 산정결과 표가 있는 물리 페이지 범위를 찾는다.

특정 키워드("개발단계별 유출곡선지수(CN) 산정결과" 등)로 본문 이동 후
현재 커서의 페이지 번호를 읽어 CN 섹션 시작/끝 페이지를 보고한다.
"""
from __future__ import annotations

import time
from pathlib import Path

SRC = Path(r"D:/DATA/카카오톡 받은 파일/제4장 재해영향 예측 및 평가_최종.hwpx")

ANCHORS_START = [
    "우리나라 토지이용 형태에 따른 유출곡선지수",
    "[표 4-19]",
    "유출곡선지수(CN) 산정결과",
]
ANCHORS_END = [
    "홍수량 산정",
    "홍수량 산정방법 선정",
    "도달시간 산정결과",
]


def current_page(hwp) -> int:
    """현재 커서의 물리 페이지 번호."""
    try:
        return int(hwp.hwp.KeyIndicator()[3])  # (sec,page,col,line,...)
    except Exception:
        pass
    try:
        return int(hwp.hwp.GetPos()[1])
    except Exception:
        return -1


def find_first(hwp, keyword: str) -> int:
    """문서 처음부터 keyword 찾기. 찾은 페이지 번호 반환, 없으면 -1."""
    hwp.HAction.Run("MoveDocBegin")
    act = hwp.hwp.CreateAction("ForwardFind")
    pset = act.CreateSet()
    act.GetDefault(pset)
    pset.SetItem("FindString", keyword)
    pset.SetItem("IgnoreMessage", 1)
    pset.SetItem("Direction", 0)
    pset.SetItem("WholeWordOnly", 0)
    pset.SetItem("UseWildCards", 0)
    pset.SetItem("IgnoreCase", 1)
    ok = act.Execute(pset)
    if not ok:
        return -1
    return current_page(hwp)


def run() -> None:
    from pyhwpx import Hwp

    t0 = time.time()
    print(f"원본 로드...")
    hwp = Hwp(visible=False)
    hwp.open(str(SRC))
    print(f"  loaded ({time.time()-t0:.1f}s)")

    # 현재 커서 페이지 (문서 시작)
    hwp.HAction.Run("MoveDocBegin")
    print(f"문서 시작 page = {current_page(hwp)}")

    # 문서 끝 페이지
    hwp.HAction.Run("MoveDocEnd")
    print(f"문서 끝   page = {current_page(hwp)}")

    # 시작 앵커 검색
    for kw in ANCHORS_START:
        p = find_first(hwp, kw)
        print(f"  START  {kw!r} -> page {p}")

    # 끝 앵커 검색
    for kw in ANCHORS_END:
        p = find_first(hwp, kw)
        print(f"  END    {kw!r} -> page {p}")

    try:
        hwp.quit()
    except Exception:
        pass


if __name__ == "__main__":
    run()
