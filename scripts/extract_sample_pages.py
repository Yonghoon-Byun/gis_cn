"""원본 HWPX에서 CN 섹션만 추출 — 키워드 검색 기반 (페이지 번호 불안정 회피).

시작 앵커("우리나라 토지이용 형태에 따른 유출곡선지수")를 찾아 그 문단 시작점에 커서를
두고, 끝 앵커("홍수량 산정방법")까지 블록 선택 후 복사·붙여넣기.
"""
from __future__ import annotations

import time
from pathlib import Path

SRC = Path(r"D:/DATA/카카오톡 받은 파일/제4장 재해영향 예측 및 평가_최종.hwpx")
OUT = Path(__file__).resolve().parent.parent / "gis_cn" / "templates" / "v1.0" / "cn_report_sample.hwpx"
TMP_OUT = Path(r"C:/temp/cn_report_sample.hwpx")

START_ANCHOR  = "우리나라 토지이용 형태에 따른 유출곡선지수"  # CN 분류표 시작
PAGES_TO_COPY = 10   # START에서부터 이 수만큼 선택 (CN 섹션 전체 커버)


def _find(hwp, keyword: str, select: bool = False) -> bool:
    act = hwp.hwp.CreateAction("ForwardFind")
    pset = act.CreateSet()
    act.GetDefault(pset)
    pset.SetItem("FindString", keyword)
    pset.SetItem("IgnoreMessage", 1)
    pset.SetItem("Direction", 0)
    pset.SetItem("WholeWordOnly", 0)
    pset.SetItem("UseWildCards", 0)
    pset.SetItem("IgnoreCase", 1)
    if select:
        pset.SetItem("FindType", 1)  # 찾은 문자열을 선택
    return bool(act.Execute(pset))


def _cur_pos(hwp):
    """(list_id, para_id, char_pos) - SetPos/GetPos 용."""
    return hwp.hwp.GetPos()


def _set_pos(hwp, pos):
    hwp.hwp.SetPos(*pos)


def _page(hwp) -> int:
    try:
        return int(hwp.hwp.KeyIndicator()[3])
    except Exception:
        return -1


def run() -> None:
    from pyhwpx import Hwp

    if not SRC.exists():
        raise FileNotFoundError(f"원본 파일이 없습니다: {SRC}")

    t0 = time.time()
    print(f"원본 열기...")
    src_hwp = Hwp(visible=False)
    src_hwp.open(str(SRC))
    print(f"  loaded ({time.time()-t0:.1f}s)")

    # 1) START 앵커로 이동 후 문단 시작점에 커서
    src_hwp.HAction.Run("MoveDocBegin")
    if not _find(src_hwp, START_ANCHOR):
        raise RuntimeError(f"START 앵커 못 찾음: {START_ANCHOR!r}")
    src_hwp.HAction.Run("MoveParaBegin")
    print(f"START: page={_page(src_hwp)}  pos={_cur_pos(src_hwp)}")

    # 2) 고정 페이지 수 만큼 Shift-PageDown 으로 선택
    for i in range(PAGES_TO_COPY):
        src_hwp.HAction.Run("MoveSelPageDown")
    print(f"END:   page={_page(src_hwp)}  pos={_cur_pos(src_hwp)}  ({PAGES_TO_COPY}p)")

    # 3) 복사
    src_hwp.HAction.Run("Copy")
    print("  copied")

    # 5) 새 문서 + 붙여넣기
    print("→ new document + paste")
    try:
        src_hwp.hwp.XHwpDocuments.Add(0)
    except Exception:
        src_hwp.HAction.Run("FileNew")
    src_hwp.HAction.Run("Paste")

    # 6) 저장
    TMP_OUT.parent.mkdir(parents=True, exist_ok=True)
    if TMP_OUT.exists():
        TMP_OUT.unlink()
    print(f"→ save: {TMP_OUT}")
    src_hwp.save_as(str(TMP_OUT), "HWPX")
    print(f"  saved {TMP_OUT.stat().st_size:,} bytes")

    try:
        src_hwp.quit()
    except Exception as e:
        print(f"  quit 무시: {e}")

    # 7) 최종 위치로 복사
    import shutil
    OUT.parent.mkdir(parents=True, exist_ok=True)
    if OUT.exists():
        OUT.unlink()
    shutil.copy2(TMP_OUT, OUT)
    print(f"→ 최종: {OUT} ({OUT.stat().st_size:,} bytes, {time.time()-t0:.1f}s)")


if __name__ == "__main__":
    run()
