"""페이지 분할 전용 테스트 — 많은 블록으로 ROWS_PER_PAGE 초과 상황 재현.

디버그 로그를 enabled 상태로 렌더링하여 어느 페이지 분할 전략이
성공/실패했는지 콘솔에 출력한다.

사용법:
  python scripts/test_hwp_pagebreak.py          # visible=False
  python scripts/test_hwp_pagebreak.py --show   # 한글 창 가시화 (디버깅)
"""
from __future__ import annotations

import logging
import sys
from pathlib import Path
from datetime import date

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from gis_cn.core.analysis_result import (
    AnalysisResult, ProjectMeta, WatershedBlock, WatershedSummary,
    LandUseRow, CnReferenceRow,
)
from gis_cn.core.hwp_renderer import render_hwp, ROWS_PER_PAGE

TPL = ROOT / "gis_cn" / "templates" / "v1.0" / "cn_report.hwpx"
OUT = Path(r"C:\temp\test_pagebreak.hwpx")


def _mk_row(lu: str, area: float) -> LandUseRow:
    """토지이용별 데이터 행."""
    return LandUseRow(
        land_use=lu,
        a_area=area * 0.25, a_cn=77,
        b_area=area * 0.25, b_cn=85,
        c_area=area * 0.25, c_cn=90,
        d_area=area * 0.25, d_cn=92,
        total_area=area,
        amc2_cn=86.0,
        amc3_cn=93,
    )


def _mk_block(name: str) -> WatershedBlock:
    """7 토지이용 × 소유역 1개 — 실제 시나리오 유사."""
    land_uses = ["공급처리시설", "공업지", "논", "임야", "초지", "하천", "주거지"]
    rows = [_mk_row(lu, 1000.0) for lu in land_uses]
    return WatershedBlock(
        name=name,
        rows=rows,
        total_a=1750, total_b=1750, total_c=1750, total_d=1750,
        total_area=7000,
        amc2_cn=86.0, amc3_cn=93.0,
    )


def main() -> int:
    visible = "--show" in sys.argv

    logging.basicConfig(
        level=logging.DEBUG,
        format="%(levelname)s %(name)s: %(message)s",
        stream=sys.stdout,
    )

    # 5개 블록 × 8행(7 토지이용 + 합계) = 40행 → ROWS_PER_PAGE=22 초과 → 분할 예상
    blocks = [_mk_block(nm) for nm in ("GB.1", "MN2b", "MN3", "MN1", "MN2a")]
    total_rows = sum(len(b.rows) + 1 for b in blocks)

    print(f"ROWS_PER_PAGE = {ROWS_PER_PAGE}")
    print(f"blocks: {len(blocks)}, total_rows = {total_rows}")

    summary = [
        WatershedSummary(name=b.name, total_area=7000, amc2_cn=86.0, amc3_cn=93.0)
        for b in blocks
    ]

    result = AnalysisResult(
        meta=ProjectMeta(
            project_name="페이지 분할 테스트",
            site_name="테스트 지구",
            author="테스터",
            analysis_date=date(2026, 4, 23),
            development_stage="현 상태",
        ),
        cn_reference=[
            CnReferenceRow(land_use="논",   a=72, b=81, c=88, d=91),
            CnReferenceRow(land_use="임야", a=36, b=60, c=73, d=79),
        ],
        detail_blocks=blocks,
        summary_rows=summary,
        notes="페이지 분할 디버깅",
    )

    OUT.parent.mkdir(parents=True, exist_ok=True)
    if OUT.exists():
        OUT.unlink()

    print(f"\nrendering → {OUT} (visible={visible})")
    render_hwp(result, TPL, OUT, visible=visible)
    print(f"saved: {OUT.stat().st_size:,} bytes")
    print(f"\n파일을 한글에서 직접 열어 블록이 올바르게 다음 쪽으로 넘어가는지 확인:")
    print(f"  {OUT}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
