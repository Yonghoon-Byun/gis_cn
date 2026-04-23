"""templates/v1.0/cn_report.hwp 빈 템플릿 자동 생성.

한글(HWP) + pyhwpx로 누름틀/표/책갈피를 스펙(`templates/v1.0/README.md`)대로
심어 저장한다. 한 번만 실행하면 된다. 스펙이 바뀌면 재실행.

전제: 한글 설치 + pyhwpx import 동작. (사전 검증: `python -c "from pyhwpx import Hwp"`)
"""
from __future__ import annotations

from pathlib import Path

OUT = Path(__file__).resolve().parent.parent / "gis_cn" / "templates" / "v1.0" / "cn_report.hwpx"


# ── 구조 정의 (README.md 섹션 4와 일치) ─────────────────────────────────────
META_FIELDS = [
    ("과업명",   "meta.project_name"),
    ("사업지구", "meta.site_name"),
    ("작성자",   "meta.author"),
    ("기관",     "meta.organization"),
    ("분석일자", "meta.analysis_date"),
    ("개발단계", "meta.development_stage"),
]

# 필드명은 namespace 접두어 필수 (여러 표 간 이름 충돌 방지).
# hwp_renderer.py의 {REF,RES,SUM,CRES,CSUM}_PREFIX와 1:1 동기화.

def _ns(prefix: str, local: str) -> str:
    return f"{prefix}.{local}"

CN_REF_HEADERS = ["토지이용분류", "A", "B", "C", "D"]
CN_REF_LOCAL   = ["lu", "cn_a", "cn_b", "cn_c", "cn_d"]

# CN result 표 — 2행 복합 헤더 (원본 보고서 스타일)
#   Row0: 유역구분 | 토지이용 | TYPE A (merge→) | TYPE B (merge→) | TYPE C (merge→) | TYPE D (merge→) | 총면적(㎡) | AMC조건 (merge→)
#   Row1: [merge↓] | [merge↓] | 면적(㎡) | CN | 면적(㎡) | CN | 면적(㎡) | CN | 면적(㎡) | CN | [merge↓] | AMC-Ⅱ | AMC-Ⅲ
CN_RESULT_HEADER_ROW0 = [
    "유역구분", "토지이용분류",
    "TYPE A", "",
    "TYPE B", "",
    "TYPE C", "",
    "TYPE D", "",
    "총면적(㎡)",
    "AMC조건", "",
]
CN_RESULT_HEADER_ROW1 = [
    "", "",
    "면적(㎡)", "CN",
    "면적(㎡)", "CN",
    "면적(㎡)", "CN",
    "면적(㎡)", "CN",
    "",
    "AMC-Ⅱ", "AMC-Ⅲ",
]
# 호환용 (단일 행 헤더 표현 — 참고용)
CN_RESULT_HEADERS = CN_RESULT_HEADER_ROW0
CN_RESULT_LOCAL = [
    "ws", "lu",
    "a_area", "a_cn", "b_area", "b_cn",
    "c_area", "c_cn", "d_area", "d_cn",
    "total_area", "amc2_cn", "amc3_cn",
]

CN_SUMMARY_HEADERS = ["유역", "총면적", "AMC2 CN", "AMC3 CN"]
CN_SUMMARY_LOCAL   = ["ws", "area", "amc2", "amc3"]

NUM_MAP_BOOKMARKS = 3   # 기본 bm_map_0..2 책갈피 자리 예약


# ── 헬퍼 ─────────────────────────────────────────────────────────────────────

def _insert_field(hwp, name: str, hint: str = "") -> None:
    """누름틀 삽입 후 커서를 필드 밖으로 뺀다.

    중요 제약 두 가지:
      1) create_field 직후 커서가 누름틀 안에 남으면 이후 삽입이 그 안에서 일어남
         → 반드시 MoveRight 로 필드 밖으로 빠져나온다.
      2) `direction`(안내문/가이드 텍스트)이 비어있지 않으면 이 HWP 버전에서
         `PutFieldText`가 첫 호출 이후 모든 필드를 '잃어버리는' 버그가 있다.
         → 안내문은 무조건 빈 문자열로 고정.

    `hint` 인자는 로깅/주석 용도로만 남겨두고 실제 direction엔 전달하지 않는다.
    """
    _ = hint  # unused (bug workaround)
    try:
        hwp.create_field(name=name, direction="", memo="")
    except TypeError:
        hwp.create_field("", "", name)
    hwp.HAction.Run("MoveRight")


def _newline(hwp, n: int = 1) -> None:
    for _ in range(n):
        hwp.HAction.Run("BreakPara")


def _insert_bookmark(hwp, name: str) -> None:
    """책갈피 삽입 — 커서 위치에 이름으로."""
    act = hwp.CreateAction("Bookmark")
    pset = act.CreateSet()
    act.GetDefault(pset)
    pset.SetItem("Name", name)
    act.Execute(pset)


def _create_table(hwp, rows: int, cols: int) -> None:
    """표 삽입. pyhwpx `create_table`이 있으면 그걸, 없으면 액션으로."""
    if hasattr(hwp, "create_table"):
        try:
            hwp.create_table(rows=rows, cols=cols)
            return
        except Exception:
            pass
    act = hwp.CreateAction("TableCreate")
    pset = act.CreateSet()
    act.GetDefault(pset)
    pset.SetItem("Rows", rows)
    pset.SetItem("Cols", cols)
    pset.SetItem("WidthType", 2)
    pset.SetItem("HeightType", 1)
    act.Execute(pset)


def _set_table_layout_props(hwp) -> None:
    """첫 행을 '제목 셀'로 지정 → 쪽마다 헤더 반복. 실패 시 무시.

    주의: 이 함수는 반드시 _create_table 직후, 표가 비어있을 때 호출되어야 함.
    커서는 표 A1에 있어야 하며, 함수 종료 후 A1에 남도록 한다.
    """
    import logging
    log = logging.getLogger(__name__)
    try:
        hwp.HAction.Run("TableColBegin")
        hwp.HAction.Run("TableSelRow")
        hwp.HAction.Run("TableHeaderCell")
        hwp.HAction.Run("Cancel")
        hwp.HAction.Run("TableColBegin")
    except Exception as e:
        log.debug(f"TableHeaderCell 설정 실패: {e}")


def _move_right(hwp) -> None:
    hwp.HAction.Run("TableRightCell")


def _move_next_row_first(hwp) -> None:
    """다음 행의 첫 셀로 이동 (표 안)."""
    hwp.HAction.Run("TableLowerCell")
    hwp.HAction.Run("TableColBegin")


def _escape_table(hwp) -> None:
    """표 밖으로 커서 이동."""
    hwp.HAction.Run("CloseEx")
    _newline(hwp)


def _fill_header_row(hwp, headers: list[str]) -> None:
    for i, h in enumerate(headers):
        hwp.insert_text(h)
        if i < len(headers) - 1:
            _move_right(hwp)


def _fill_field_row(hwp, fields: list[str]) -> None:
    for i, f in enumerate(fields):
        _insert_field(hwp, f, f)
        if i < len(fields) - 1:
            _move_right(hwp)


# ── 섹션 작성 ────────────────────────────────────────────────────────────────

def _write_title(hwp) -> None:
    hwp.insert_text("재해영향평가 CN값 산정 보고서")
    _newline(hwp, 2)


def _write_meta_block(hwp) -> None:
    for label, name in META_FIELDS:
        hwp.insert_text(f"{label}: ")
        _insert_field(hwp, name, label)
        _newline(hwp)
    _newline(hwp)


def _write_table_with_fields(hwp, title: str, headers, fields: list[str],
                              source_note: str = "", data_rows: int = 20) -> None:
    """헤더(1행 또는 여러 행) + 데이터 data_rows행.

    Args:
        headers: list[str] 단일 헤더 | list[list[str]] 복합 헤더(여러 행)
        fields:  데이터 행의 셀별 누름틀 이름 (전체 열 수 결정)
        data_rows: 데이터 행 수 (미리 할당; 렌더러가 필요한 만큼만 채움)
    """
    # headers 정규화
    if headers and isinstance(headers[0], list):
        header_rows = headers
    else:
        header_rows = [headers]

    n_cols = len(fields)
    for hr in header_rows:
        if len(hr) != n_cols:
            raise ValueError(f"헤더 행 열 수({len(hr)})와 필드 개수({n_cols}) 불일치")

    hwp.insert_text(title)
    _newline(hwp)
    total_rows = len(header_rows) + max(1, data_rows)
    _create_table(hwp, rows=total_rows, cols=n_cols)

    # 헤더 행들 채우기 (커서: A1 → 헤더 전체 순회 → 마지막 데이터 행)
    for hi, hr in enumerate(header_rows):
        if hi > 0:
            _move_next_row_first(hwp)
        _fill_header_row(hwp, hr)

    # 데이터 행들
    for r in range(data_rows):
        _move_next_row_first(hwp)
        _fill_field_row(hwp, fields)

    _escape_table(hwp)
    if source_note:
        hwp.insert_text(source_note)
        _newline(hwp, 2)


def _write_image_bookmarks(hwp) -> None:
    hwp.insert_text("[그림] 삽도 영역")
    _newline(hwp)
    for i in range(NUM_MAP_BOOKMARKS):
        hwp.insert_text(f"  · ")
        _insert_bookmark(hwp, f"bm_map_{i}")
        _insert_field(hwp, f"caption_{i}", f"지도 설명 {i}")
        _newline(hwp)
    _newline(hwp)


def _write_notes(hwp) -> None:
    hwp.insert_text("[비고]")
    _newline(hwp)
    _insert_bookmark(hwp, "bm_null_rows")
    _newline(hwp)
    _insert_field(hwp, "notes", "비고")
    _newline(hwp)


# ── 엔트리 ───────────────────────────────────────────────────────────────────

def main() -> None:
    import shutil
    from pyhwpx import Hwp

    OUT.parent.mkdir(parents=True, exist_ok=True)
    if OUT.exists():
        OUT.unlink()

    # 긴 한글 경로(SaveAs RPC 실패 회피)를 피하기 위해 임시 ASCII 경로에 저장 후 복사
    tmp = Path(r"C:\temp\cn_report_tmp.hwpx")
    tmp.parent.mkdir(parents=True, exist_ok=True)
    if tmp.exists():
        tmp.unlink()

    hwp = Hwp(visible=False)
    try:
        print("→ title/meta")
        _write_title(hwp)
        _write_meta_block(hwp)

        print("→ table: cn reference")
        _write_table_with_fields(
            hwp,
            title="[표 4-1] 우리나라 토지이용 형태에 따른 유출곡선지수 (AMC-II, Ia=0.2S)",
            headers=CN_REF_HEADERS,
            fields=[_ns("ref", x) for x in CN_REF_LOCAL],
            source_note="[자료] 수치토지피복도 분류기준",
            data_rows=20,
        )

        print("→ table: cn result (compound header)")
        _write_table_with_fields(
            hwp,
            title="[표 4-19] 소유역별 유출곡선지수(CN) 산정결과",
            headers=[CN_RESULT_HEADER_ROW0, CN_RESULT_HEADER_ROW1],
            fields=[_ns("res", x) for x in CN_RESULT_LOCAL],
            source_note="[주] 합계 행은 소유역 블록 마지막에 렌더러가 자동 삽입",
            data_rows=50,
        )

        print("→ table: cn summary")
        _write_table_with_fields(
            hwp,
            title="[표 4-20] 유역별 CN정리",
            headers=CN_SUMMARY_HEADERS,
            fields=[_ns("sum", x) for x in CN_SUMMARY_LOCAL],
            data_rows=20,
        )

        print("→ table: composite result (compound header)")
        _write_table_with_fields(
            hwp,
            title="[표 4-21] 유역합성 CN값 산정결과",
            headers=[CN_RESULT_HEADER_ROW0, CN_RESULT_HEADER_ROW1],
            fields=[_ns("cres", x) for x in CN_RESULT_LOCAL],
            data_rows=30,
        )
        print("→ table: composite summary")
        _write_table_with_fields(
            hwp,
            title="[표 4-22] 유역합성 CN정리",
            headers=CN_SUMMARY_HEADERS,
            fields=[_ns("csum", x) for x in CN_SUMMARY_LOCAL],
            data_rows=10,
        )

        print("→ image bookmarks")
        _write_image_bookmarks(hwp)
        print("→ notes")
        _write_notes(hwp)

        print(f"→ save_as (HWPX): {tmp}")
        hwp.save_as(str(tmp), "HWPX")
        print("   save OK")
    finally:
        try:
            hwp.quit()
        except Exception:
            pass

    print(f"→ copy: {tmp} → {OUT}")
    shutil.copy2(tmp, OUT)
    tmp.unlink(missing_ok=True)
    print(f"wrote: {OUT}  ({OUT.stat().st_size:,} bytes)")


if __name__ == "__main__":
    main()
