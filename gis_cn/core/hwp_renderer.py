"""HWP(.hwp) 렌더러 — pyhwpx(OLE) 기반.

`AnalysisResult`를 입력으로 받아 `templates/v1.0/cn_report.hwp` 템플릿의
누름틀/책갈피/셀필드를 치환하여 `.hwp`로 저장한다.

템플릿 스펙: `gis_cn/templates/v1.0/README.md` 참조.

필수 런타임:
  - Windows
  - Hancom Office 한글 설치 및 COM 등록(HncReg.exe)
  - `pip install pyhwpx`

OLE는 단일 워커만 안정적으로 동작하므로 UI에서 동시 실행을 막아야 한다.
"""
from __future__ import annotations

import logging
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

from .analysis_result import (
    AnalysisResult, WatershedBlock, LandUseRow, WatershedSummary, CnReferenceRow,
)

logger = logging.getLogger(__name__)


# ─────────────────────────────────────────────────────────────────────────────
# 페이지 분할 설정
# ─────────────────────────────────────────────────────────────────────────────

# A4 기준 한 쪽에 수용 가능한 데이터 행 수(2행 복합 헤더 + 본문). 경험값.
# 값이 너무 크면 넘침, 너무 작으면 페이지 과분할. 템플릿 폰트/여백 조정 시 재튜닝.
# 렌더링 결과에서 블록이 여전히 넘칠 경우 이 값을 낮추세요.
ROWS_PER_PAGE = 22

# 템플릿(cn_report.hwpx)의 표별 미리 할당된 데이터 행 수.
# gen_template.py의 data_rows 인자와 1:1 동기화. 템플릿 재생성 시 이 값도 맞춰야 함.
# 렌더 후 사용하지 않은 여분 행은 `_trim_table_rows`가 삭제하여 빈 페이지 방지.
TEMPLATE_DATA_ROWS = {
    "cn_reference":          20,  # ref.*
    "cn_result":             50,  # res.*
    "cn_summary":            20,  # sum.*
    "cn_composite":          30,  # cres.*
    "cn_composite_summary":  10,  # csum.*
}

# 헤더 반복은 gen_template.py의 `_set_table_layout_props`가 표 생성 시 TableHeaderCell 속성을
# 2행 복합 헤더에 지정하여 한글의 네이티브 "제목 줄 반복" 기능에 위임한다. 렌더러는 블록
# 경계에서 BreakPage만 삽입하면 충분하며, 분할 표에 헤더를 수동으로 복제할 필요 없음.


class HwpRendererError(RuntimeError):
    """HWP 렌더링 실패 (한글 미설치/COM 미등록/템플릿 오류 포함)."""


def _import_pyhwpx():
    """지연 import — QGIS 플러그인 로드 시점에 pyhwpx 없어도 동작하도록."""
    try:
        from pyhwpx import Hwp   # type: ignore
    except ImportError as e:
        raise HwpRendererError(
            "pyhwpx가 설치되어 있지 않습니다. "
            "`pip install pyhwpx` 후 한글(Hancom Office)을 설치·등록하세요."
        ) from e
    except Exception as e:
        raise HwpRendererError(
            f"pyhwpx 로드 실패 ({e}). 한글 설치 후 `HncReg.exe`로 COM을 재등록하세요."
        ) from e
    return Hwp


# ─────────────────────────────────────────────────────────────────────────────
# 필드 ID 네이밍 규칙 (templates/v1.0/README.md와 일치시켜야 함)
# ─────────────────────────────────────────────────────────────────────────────

META_FIELDS = {
    "project_name":   "meta.project_name",
    "site_name":      "meta.site_name",
    "author":         "meta.author",
    "organization":   "meta.organization",
    "analysis_date":  "meta.analysis_date",
    "development_stage": "meta.development_stage",
}

# 필드명은 표별로 namespace 접두어 사용 (다수 표 간 이름 충돌 방지).
#   ref.*   : CN 기준 분류표
#   res.*   : 소유역별 산정결과표
#   sum.*   : 유역별 CN정리 요약
#   cres.*  : 유역합성 산정결과표
#   csum.*  : 유역합성 요약
REF_PREFIX    = "ref"
RES_PREFIX    = "res"
SUM_PREFIX    = "sum"
CRES_PREFIX   = "cres"
CSUM_PREFIX   = "csum"

TBL_CN_REF               = "tbl_cn_reference"
TBL_CN_RESULT            = "tbl_cn_result"
TBL_SUMMARY              = "tbl_cn_summary"
TBL_COMPOSITE            = "tbl_cn_composite"
TBL_COMPOSITE_SUMMARY    = "tbl_cn_composite_summary"

# 행 단위 스칼라 필드 (prefix 제외한 로컬명)
CN_REF_LOCAL     = ("lu", "cn_a", "cn_b", "cn_c", "cn_d")
CN_RESULT_LOCAL  = (
    "ws", "lu",
    "a_area", "a_cn", "b_area", "b_cn",
    "c_area", "c_cn", "d_area", "d_cn",
    "total_area", "amc2_cn", "amc3_cn",
)
SUMMARY_LOCAL    = ("ws", "area", "amc2", "amc3")


def _ns(prefix: str, local: str) -> str:
    """prefix.local 결합 (`res.ws`, `ref.lu`)."""
    return f"{prefix}.{local}"

# 책갈피
BM_MAP_PREFIX    = "bm_map_"           # 지도 이미지 삽입 위치
BM_NULL_ROWS     = "bm_null_rows"      # CN 매칭 실패 목록


# ─────────────────────────────────────────────────────────────────────────────
# 메인 엔트리
# ─────────────────────────────────────────────────────────────────────────────

def render_hwp(result: AnalysisResult, template: Path, out: Path,
               *, visible: bool = False) -> Path:
    """AnalysisResult → HWP 렌더.

    Args:
        result:   렌더할 분석 결과 (Single Source of Truth).
        template: 빈 양식 HWP (`templates/v1.0/cn_report.hwp`).
        out:      저장 경로 (`.hwp`).
        visible:  한글 창 가시성 (디버깅 시 True).

    Raises:
        HwpRendererError: pyhwpx 미설치, 한글 미등록, 템플릿 오류 등.
    """
    Hwp = _import_pyhwpx()

    template = Path(template)
    out = Path(out)
    if not template.exists():
        raise HwpRendererError(f"템플릿을 찾을 수 없습니다: {template}")

    hwp: Optional[object] = None
    try:
        hwp = Hwp(visible=visible)
        hwp.open(str(template))

        _render_meta(hwp, result)

        _render_cn_reference(hwp, result.cn_reference)
        _trim_table_rows(hwp, _ns(REF_PREFIX, "lu"),
                         len(result.cn_reference), TEMPLATE_DATA_ROWS["cn_reference"])

        _render_detail(hwp, RES_PREFIX, result.detail_blocks)
        _trim_table_rows(hwp, _ns(RES_PREFIX, "ws"),
                         _count_detail_rows(result.detail_blocks),
                         TEMPLATE_DATA_ROWS["cn_result"])

        _render_summary(hwp, SUM_PREFIX, result.summary_rows)
        _trim_table_rows(hwp, _ns(SUM_PREFIX, "ws"),
                         len(result.summary_rows), TEMPLATE_DATA_ROWS["cn_summary"])

        _render_detail(hwp, CRES_PREFIX, result.composite_detail)
        _trim_table_rows(hwp, _ns(CRES_PREFIX, "ws"),
                         _count_detail_rows(result.composite_detail),
                         TEMPLATE_DATA_ROWS["cn_composite"])

        _render_summary(hwp, CSUM_PREFIX, result.composite_summary)
        _trim_table_rows(hwp, _ns(CSUM_PREFIX, "ws"),
                         len(result.composite_summary),
                         TEMPLATE_DATA_ROWS["cn_composite_summary"])

        _render_map_images(hwp, result.map_images)
        _render_notes(hwp, result)

        # 확장자로 포맷 결정. `.hwpx`는 OWPML(zip+xml), `.hwp`는 전통 OLE 포맷.
        fmt = "HWPX" if out.suffix.lower() == ".hwpx" else "HWP"
        hwp.save_as(str(out), fmt)
        logger.info(f"{fmt} 저장 완료: {out}")
        return out
    except HwpRendererError:
        raise
    except Exception as e:
        logger.exception("HWP 렌더링 오류")
        raise HwpRendererError(f"HWP 렌더링 실패: {e}") from e
    finally:
        if hwp is not None:
            try:
                hwp.quit()
            except Exception:
                pass


# ─────────────────────────────────────────────────────────────────────────────
# 개별 섹션 렌더러 — 누름틀/책갈피/셀필드 치환
# ─────────────────────────────────────────────────────────────────────────────

def _put_field_safe(hwp, name: str, value: str, idx: int = 0) -> None:
    """pyhwpx: put_field_text(field, text, idx). 해당 필드/인덱스 없어도 조용히 패스."""
    try:
        hwp.put_field_text(name, value, idx)
    except Exception as e:
        logger.debug(f"필드 {name}[{idx}] 치환 건너뜀: {e}")


def _count_detail_rows(blocks: list[WatershedBlock]) -> int:
    """detail 블록들의 총 데이터 행 수(토지이용 행 + 블록별 합계 행)."""
    return sum(len(b.rows) + 1 for b in blocks)


def _trim_table_rows(hwp, field_name: str, first_unused_idx: int,
                     template_total: int) -> None:
    """템플릿의 데이터 행 중 채우지 않은 여분을 삭제.

    `field_name` 필드의 idx=first_unused_idx 부터 template_total-1 까지 각 행을 제거.
    삭제는 idx 높은 순으로(뒤→앞) 수행하여 인덱스 재정렬 문제를 회피.
    `text=False`로 누름틀 코드에 위치시켜 셀 삭제가 누름틀 텍스트 스코프가 아닌
    셀 자체에 적용되도록 함. 실패 시 조용히 패스.
    """
    if template_total <= 0 or first_unused_idx >= template_total:
        return
    ok, fail = 0, 0
    for idx in range(template_total - 1, first_unused_idx - 1, -1):
        try:
            try:
                hwp.move_to_field(field_name, idx=idx, text=False, start=True)
            except TypeError:
                hwp.move_to_field(field_name)
            hwp.HAction.Run("TableDeleteRow")
            ok += 1
        except Exception as e:
            fail += 1
            logger.debug(f"행 삭제 실패 {field_name}[{idx}]: {e}")
    logger.info(f"trim {field_name}: {ok} deleted, {fail} failed (range {first_unused_idx}..{template_total-1})")


def _fmt_area(v: Optional[float]) -> str:
    return "" if v is None else f"{v:,.3f}"


def _fmt_cn(v) -> str:
    if v is None:
        return ""
    if isinstance(v, float):
        return f"{v:.2f}"
    return str(v)


def _render_meta(hwp, result: AnalysisResult) -> None:
    m = result.meta
    _put_field_safe(hwp, META_FIELDS["project_name"],     m.project_name)
    _put_field_safe(hwp, META_FIELDS["site_name"],        m.site_name)
    _put_field_safe(hwp, META_FIELDS["author"],           m.author)
    _put_field_safe(hwp, META_FIELDS["organization"],     m.organization)
    _put_field_safe(hwp, META_FIELDS["analysis_date"],    m.analysis_date.isoformat())
    _put_field_safe(hwp, META_FIELDS["development_stage"], m.development_stage)


def _render_cn_reference(hwp, rows: list[CnReferenceRow]) -> None:
    """p.25 기준 분류표 — 템플릿의 미리 할당된 행에 idx로 채움."""
    if not rows:
        return
    for i, row in enumerate(rows):
        _put_field_safe(hwp, _ns(REF_PREFIX, "lu"),   row.land_use,  i)
        _put_field_safe(hwp, _ns(REF_PREFIX, "cn_a"), _fmt_cn(row.a), i)
        _put_field_safe(hwp, _ns(REF_PREFIX, "cn_b"), _fmt_cn(row.b), i)
        _put_field_safe(hwp, _ns(REF_PREFIX, "cn_c"), _fmt_cn(row.c), i)
        _put_field_safe(hwp, _ns(REF_PREFIX, "cn_d"), _fmt_cn(row.d), i)


def _plan_page_breaks(blocks: list[WatershedBlock],
                      rows_per_page: int = ROWS_PER_PAGE) -> list[int]:
    """블록 누적 행수가 rows_per_page 초과 시 해당 블록 시작 직전에 넣을 페이지 브레이크 인덱스 리스트.

    블록 하나는 두 페이지로 나누지 않음(내부 분할은 한글에 위임).
    블록 행수 = `len(block.rows) + 1` (+1은 합계 행).
    """
    breaks: list[int] = []
    if not blocks or rows_per_page <= 0:
        return breaks
    acc = 0
    for i, block in enumerate(blocks):
        block_rows = len(block.rows) + 1
        if i > 0 and acc + block_rows > rows_per_page:
            breaks.append(i)
            acc = block_rows
        else:
            acc += block_rows
    return breaks


def _render_detail(hwp, prefix: str, blocks: list[WatershedBlock]) -> None:
    """p.26~ 소유역별 상세 — 블록 반복 + 각 블록 내 행 복제.

    ROWS_PER_PAGE를 기준으로 블록 경계에서 페이지 강제 분할. 블록 내부는 한글의
    자동 페이지 분할 + 표 속성(헤더 반복/셀 나눔 금지)에 위임.

    prefix는 네임스페이스('res' 또는 'cres')로 템플릿 표와 1:1 매칭됨.
    """
    if not blocks:
        return

    break_indices = set(_plan_page_breaks(blocks, ROWS_PER_PAGE))
    global_idx = 0
    for i, block in enumerate(blocks):
        if i in break_indices:
            _insert_page_break_before_row(hwp, prefix, global_idx)
        for row in block.rows:
            _write_detail_row(hwp, prefix, global_idx, block.name, row)
            global_idx += 1
        _write_detail_summary_row(hwp, prefix, global_idx, block)
        global_idx += 1


def _insert_page_break_before_row(hwp, prefix: str, idx: int) -> None:
    """idx번째 데이터 행 시작 직전에 페이지 강제 분할.

    전략: 해당 행의 ws 누름틀로 `text=False`(누름틀 코드 위치)로 이동하여 셀 시작점에
    커서를 둔 뒤 `BreakPage` 실행. 헤더 반복은 템플릿의 TableHeaderCell 속성에 위임
    (gen_template.py가 표 생성 시 복합 헤더 2행을 제목 셀로 지정).

    `text=True`로 누름틀 텍스트 내부에 들어간 상태에서 BreakPage를 실행하면 누름틀
    안에 페이지 문자가 들어가 페이지가 실제 분할되지 않는 문제가 있었음 — text=False로 해결.
    """
    field_name = _ns(prefix, "ws")
    try:
        try:
            hwp.move_to_field(field_name, idx=idx, text=False, start=True)
        except TypeError:
            hwp.move_to_field(field_name)
    except Exception as e:
        logger.debug(f"move_to_field 실패 ({field_name}[{idx}]): {e}")
        return

    try:
        hwp.HAction.Run("TableColBegin")
    except Exception:
        pass

    try:
        hwp.HAction.Run("BreakPage")
        logger.debug(f"BreakPage OK ({field_name}[{idx}])")
    except Exception as e:
        logger.warning(f"BreakPage 실패 ({field_name}[{idx}]): {e}")


def _write_detail_row(hwp, prefix: str, idx: int, ws_name: str, row: LandUseRow) -> None:
    put = lambda local, val: _put_field_safe(hwp, _ns(prefix, local), val, idx)
    put("ws",         ws_name)
    put("lu",         row.land_use)
    put("a_area",     _fmt_area(row.a_area))
    put("a_cn",       _fmt_cn(row.a_cn))
    put("b_area",     _fmt_area(row.b_area))
    put("b_cn",       _fmt_cn(row.b_cn))
    put("c_area",     _fmt_area(row.c_area))
    put("c_cn",       _fmt_cn(row.c_cn))
    put("d_area",     _fmt_area(row.d_area))
    put("d_cn",       _fmt_cn(row.d_cn))
    put("total_area", _fmt_area(row.total_area))
    put("amc2_cn",    _fmt_cn(row.amc2_cn))
    put("amc3_cn",    _fmt_cn(row.amc3_cn))


def _write_detail_summary_row(hwp, prefix: str, idx: int, block: WatershedBlock) -> None:
    put = lambda local, val: _put_field_safe(hwp, _ns(prefix, local), val, idx)
    put("ws",         block.name)
    put("lu",         "합계")
    put("a_area",     _fmt_area(block.total_a or None))
    put("a_cn",       "")
    put("b_area",     _fmt_area(block.total_b or None))
    put("b_cn",       "")
    put("c_area",     _fmt_area(block.total_c or None))
    put("c_cn",       "")
    put("d_area",     _fmt_area(block.total_d or None))
    put("d_cn",       "")
    put("total_area", _fmt_area(block.total_area))
    put("amc2_cn",    _fmt_cn(block.amc2_cn))
    put("amc3_cn",    _fmt_cn(block.amc3_cn))


def _render_summary(hwp, prefix: str, rows: list[WatershedSummary]) -> None:
    if not rows:
        return
    for i, row in enumerate(rows):
        _put_field_safe(hwp, _ns(prefix, "ws"),   row.name,                 i)
        _put_field_safe(hwp, _ns(prefix, "area"), _fmt_area(row.total_area), i)
        _put_field_safe(hwp, _ns(prefix, "amc2"), _fmt_cn(row.amc2_cn),      i)
        _put_field_safe(hwp, _ns(prefix, "amc3"), _fmt_cn(row.amc3_cn),      i)


def _render_map_images(hwp, images) -> None:
    for i, img in enumerate(images):
        bm = img.bookmark_id or f"{BM_MAP_PREFIX}{i}"
        try:
            if hasattr(hwp, "move_to_field"):
                hwp.move_to_field(bm)
            if hasattr(hwp, "insert_picture"):
                hwp.insert_picture(str(img.path))
            if img.caption:
                _put_field_safe(hwp, f"caption_{i}", img.caption)
        except Exception as e:
            logger.warning(f"지도 이미지 삽입 실패 ({bm}): {e}")


def _render_notes(hwp, result: AnalysisResult) -> None:
    _put_field_safe(hwp, "notes", result.notes)
    if result.null_cn_rows:
        try:
            if hasattr(hwp, "move_to_field"):
                hwp.move_to_field(BM_NULL_ROWS)
            text = "\n".join(
                f"- {r.watershed} / {r.land_use} / {r.hydro_type}"
                for r in result.null_cn_rows
            )
            if hasattr(hwp, "insert_text"):
                hwp.insert_text(text)
        except Exception as e:
            logger.warning(f"NULL CN 목록 삽입 실패: {e}")


# ─────────────────────────────────────────────────────────────────────────────
# 헬퍼 — 템플릿 표 행 수 맞추기
# ─────────────────────────────────────────────────────────────────────────────

def _ensure_table_rows(hwp, table_name: str, anchor_field: str, n: int) -> None:
    """템플릿 표의 행 수를 n개로 맞춘다.

    템플릿에는 데이터 행 1줄만 두고, 렌더링 시 필요한 만큼 복제한다.
    anchor_field(첫 열의 누름틀 이름)로 커서 이동 후 TableInsertRowBelow 반복.

    복제된 행의 누름틀들은 pyhwpx의 `put_field_text(name, text, idx)`에서
    idx=1, 2, ... 로 접근된다. idx=0는 템플릿 원본 행.
    """
    if n <= 0:
        return
    try:
        hwp.move_to_field(anchor_field)
    except Exception as e:
        logger.debug(f"anchor '{anchor_field}' 이동 실패 (표 '{table_name}'): {e}")
        return
    for _ in range(max(0, n - 1)):
        try:
            hwp.HAction.Run("TableInsertRowBelow")
        except Exception as e:
            logger.warning(f"행 복제 실패 (표 '{table_name}'): {e}")
            break
