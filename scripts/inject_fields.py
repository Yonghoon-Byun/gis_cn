"""샘플 HWPX(cn_report_sample.hwpx)의 CN 결과 표 데이터 셀에 누름틀 삽입.

단계:
 1) 맨 앞 2페이지 삭제 (사용자 요청)
 2) CN 기준표(8열)는 유지 — 표준 분류
 3) CN 산정결과 표(9표면 14실열) 중 첫 표 하나만 유지
 4) 유지한 표의 데이터 행 1개만 남기고 나머지 삭제
 5) 남긴 데이터 행의 셀에 우리 스펙 누름틀 삽입
 6) 결과를 templates/v1.0/cn_report.hwpx로 저장 (기존 자동생성본 덮어씀)
"""
from __future__ import annotations

import time
from pathlib import Path

# 사용자가 한글에서 샘플 데이터를 제거한 clean 템플릿을 입력으로 사용
SRC = Path(r"C:/temp/cn_report_cleaned.hwpx")
OUT = Path(__file__).resolve().parent.parent / "gis_cn" / "templates" / "v1.0" / "cn_report.hwpx"
TMP_OUT = Path(r"C:/temp/cn_report_injected.hwpx")

# CN 결과 표 데이터 행 — 왼쪽부터 14개 셀 (C0~C13)
# 원본 레이아웃:
#   C0: 단계, C1: 유역구분, C2: 토지이용
#   C3: A면적, C4: A_CN, C5: B면적, C6: B_CN, C7: C면적, C8: C_CN, C9: D면적, C10: D_CN
#   C11: 총면적, C12: AMC-Ⅱ CN, C13: AMC-Ⅲ CN
DATA_CELL_FIELDS = [
    None,          # C0 단계 — 사용자가 원본 유지("개" 등)
    "res.ws",      # C1 유역
    "res.lu",      # C2 토지이용
    "res.a_area", "res.a_cn",
    "res.b_area", "res.b_cn",
    "res.c_area", "res.c_cn",
    "res.d_area", "res.d_cn",
    "res.total_area", "res.amc2_cn", "res.amc3_cn",
]


def _clear_cell(hwp):
    """현재 커서가 위치한 셀의 모든 내용 삭제."""
    # 셀 시작으로
    hwp.HAction.Run("MoveCellBegin")
    # 셀 끝까지 선택
    hwp.HAction.Run("MoveSelCellEnd")
    hwp.HAction.Run("Delete")


def _insert_field(hwp, name: str):
    """현재 커서 위치에 누름틀 삽입 (direction은 빈 문자열 — 다중치환 버그 회피)."""
    hwp.create_field(name=name, direction="", memo="")
    hwp.HAction.Run("MoveRight")


def _delete_row(hwp):
    """현재 커서가 있는 표 행 삭제."""
    hwp.HAction.Run("TableDeleteRow")


def _delete_table_below(hwp):
    """현재 커서 아래의 다음 표를 찾아 통째로 삭제."""
    # 표 밖으로 나가서 다음 표 시작까지 이동
    hwp.HAction.Run("TableRightCell")  # 표 안이면 무해
    hwp.HAction.Run("CloseEx")


def run():
    from pyhwpx import Hwp

    t0 = time.time()
    print(f"SRC: {SRC}")
    if not SRC.exists():
        raise FileNotFoundError(SRC)

    hwp = Hwp(visible=False)
    hwp.open(str(SRC))
    print(f"opened ({time.time()-t0:.1f}s)")

    # === 1) 맨 앞 2페이지 삭제 ===
    print("→ 앞 2페이지 삭제")
    hwp.HAction.Run("MoveDocBegin")
    for _ in range(2):
        hwp.HAction.Run("MoveSelPageDown")
    hwp.HAction.Run("Delete")

    # === 1.5) CN 기준표(table[0], 43행 × 8열) 데이터 셀을 비우고 ref.* 누름틀 주입 ===
    # 구조: 대분류 | 중분류 | 코드 | 세분류 | A | B | C | D
    # 대/중/코드는 표준 분류체계(정적) — 유지. 세분류 + A/B/C/D = 누름틀로.
    #   C3: ref.lu, C4: ref.cn_a, C5: ref.cn_b, C6: ref.cn_c, C7: ref.cn_d
    print("→ CN 기준표(43행) 이동 및 누름틀 주입")
    hwp.HAction.Run("MoveDocBegin")
    act_find = hwp.hwp.CreateAction("ForwardFind")
    pset_find = act_find.CreateSet()
    act_find.GetDefault(pset_find)
    pset_find.SetItem("FindString", "토지이용 형태(수치토지피복도 분류기준)")
    pset_find.SetItem("IgnoreMessage", 1)
    pset_find.SetItem("IgnoreCase", 1)
    act_find.Execute(pset_find)
    # 헤더 행(R0, R1) 지나 데이터 첫 행(R2) 이동
    hwp.HAction.Run("TableLowerCell")
    hwp.HAction.Run("TableLowerCell")
    hwp.HAction.Run("TableColBegin")

    REF_CELL_FIELDS = [None, None, None, "ref.lu", "ref.cn_a", "ref.cn_b", "ref.cn_c", "ref.cn_d"]
    ref_rows_done = 0
    for ri in range(60):  # 최대 60행 시도 (실제 ~41 데이터 행)
        hwp.HAction.Run("TableColBegin")
        for ci, fname in enumerate(REF_CELL_FIELDS):
            if fname:
                # 셀은 이미 비어있음 (사용자가 사전 clean) → 바로 누름틀만 삽입
                _insert_field(hwp, fname)
            if ci < len(REF_CELL_FIELDS) - 1:
                hwp.HAction.Run("TableRightCell")
        ref_rows_done += 1
        ok_down = hwp.HAction.Run("TableLowerCell")
        if not ok_down:
            break
    print(f"  ref 기준표 주입: {ref_rows_done}행")

    # === 2) 첫 번째 CN 결과 표로 이동 ===
    # CN 기준표(table[0])는 건너뛰고 첫 산정결과 표로. "Type A" 문자열이 헤더에 있음.
    print("→ 첫 CN 산정결과 표 찾기")
    act = hwp.hwp.CreateAction("ForwardFind")
    pset = act.CreateSet()
    act.GetDefault(pset)
    pset.SetItem("FindString", "Type A")
    pset.SetItem("IgnoreMessage", 1)
    pset.SetItem("Direction", 0)
    pset.SetItem("IgnoreCase", 1)
    ok = act.Execute(pset)
    print(f"  find Type A: {ok}")

    # 헤더 행 이후 데이터 첫 행으로 이동 — 2번 TableLowerCell
    print("→ 데이터 첫 행(R2)으로 이동")
    hwp.HAction.Run("TableLowerCell")
    hwp.HAction.Run("TableLowerCell")
    hwp.HAction.Run("TableColBegin")
    time.sleep(0.2)

    # === 3) 데이터 행 누름틀 주입 + 부족 시 TableAppendRow 로 자동 확장 ===
    # 목표: TARGET_ROWS 개의 누름틀 데이터 행. 샘플 정리 후 남은 빈 행 수가 적으면
    # TableAppendRow 로 확장하고 새 행에도 누름틀 주입.
    TARGET_ROWS = 60
    rows_done = 0
    table_ended = False
    for row_idx in range(TARGET_ROWS):
        hwp.HAction.Run("TableColBegin")
        for ci, fname in enumerate(DATA_CELL_FIELDS):
            if fname:
                _insert_field(hwp, fname)
            if ci < len(DATA_CELL_FIELDS) - 1:
                hwp.HAction.Run("TableRightCell")
        rows_done += 1
        if (rows_done % 10) == 0:
            print(f"  injected {rows_done} rows")

        if table_ended:
            # 이미 표 끝 도달 → 새 행 append
            hwp.HAction.Run("TableAppendRow")
            hwp.HAction.Run("TableColBegin")
            continue

        # 다음 기존 행으로 이동 시도
        ok_down = hwp.HAction.Run("TableLowerCell")
        if not ok_down:
            # 표 끝 — 이후부터는 TableAppendRow
            table_ended = True
            hwp.HAction.Run("TableAppendRow")
            hwp.HAction.Run("TableColBegin")
    print(f"→ 총 주입 행: {rows_done}")

    # === 5) 저장 (HWPX) ===
    TMP_OUT.parent.mkdir(parents=True, exist_ok=True)
    if TMP_OUT.exists():
        TMP_OUT.unlink()
    print(f"→ save: {TMP_OUT}")
    hwp.save_as(str(TMP_OUT), "HWPX")
    print(f"  saved {TMP_OUT.stat().st_size:,} bytes")

    try:
        hwp.quit()
    except Exception as e:
        print(f"  quit 무시: {e}")

    # === 6) 최종 위치로 복사 ===
    import shutil
    OUT.parent.mkdir(parents=True, exist_ok=True)
    if OUT.exists():
        OUT.unlink()
    shutil.copy2(TMP_OUT, OUT)
    print(f"→ 최종: {OUT} ({OUT.stat().st_size:,} bytes, {time.time()-t0:.1f}s)")


if __name__ == "__main__":
    run()
