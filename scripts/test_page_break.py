"""_plan_page_breaks 순수 함수 단위 테스트.

pyhwpx/한글 설치 없이 Python만으로 페이지 분할 계산 로직을 검증한다.
실행: python scripts/test_page_break.py → 모든 테스트 통과 시 "OK"로 종료.
"""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from gis_cn.core.analysis_result import WatershedBlock, LandUseRow
from gis_cn.core.hwp_renderer import _plan_page_breaks, ROWS_PER_PAGE


def _mk_block(name: str, n_rows: int) -> WatershedBlock:
    """n_rows개의 데이터 행을 가진 블록 생성 (합계 행은 렌더러가 +1 추가)."""
    row = LandUseRow(
        land_use="x",
        a_area=None, a_cn=None, b_area=None, b_cn=None,
        c_area=None, c_cn=None, d_area=None, d_cn=None,
        total_area=None, amc2_cn=None, amc3_cn=None,
    )
    return WatershedBlock(name=name, rows=[row] * n_rows)


def _assert(label: str, actual, expected) -> None:
    if actual != expected:
        raise AssertionError(f"{label}: expected {expected!r}, got {actual!r}")
    print(f"  PASS {label}: {actual}")


def test_no_break_under_threshold() -> None:
    """3개 블록 × 5행(+합계=6행) = 18행 < 25 → 브레이크 없음."""
    blocks = [_mk_block(f"B{i}", 5) for i in range(3)]
    _assert("no_break_under_threshold", _plan_page_breaks(blocks, 25), [])


def test_break_at_overflow() -> None:
    """5개 블록 × 7행(+합계=8행) = 40행, 25행 페이지 → 4번째 블록(idx=3) 시작 직전 브레이크.

    누적: 8, 16, 24, 32 → idx=3에서 24+8=32 > 25 → breaks=[3], acc 리셋=8
         idx=4: 8+8=16 ≤ 25 → 추가 브레이크 없음
    """
    blocks = [_mk_block(f"B{i}", 7) for i in range(5)]
    _assert("break_at_overflow", _plan_page_breaks(blocks, 25), [3])


def test_empty_blocks() -> None:
    _assert("empty_blocks", _plan_page_breaks([], 25), [])


def test_single_large_block() -> None:
    """단일 블록이 rows_per_page 초과해도 브레이크 없음 (블록 내부 분할은 한글에 위임)."""
    blocks = [_mk_block("Big", 29)]  # 30행
    _assert("single_large_block", _plan_page_breaks(blocks, 25), [])


def test_rows_per_page_constant() -> None:
    _assert("ROWS_PER_PAGE default", ROWS_PER_PAGE, 22)


def test_zero_or_negative_rows_per_page() -> None:
    blocks = [_mk_block(f"B{i}", 5) for i in range(3)]
    _assert("rows_per_page=0 guard",  _plan_page_breaks(blocks, 0),  [])
    _assert("rows_per_page=-1 guard", _plan_page_breaks(blocks, -1), [])


def test_multi_page_breaks() -> None:
    """10개 블록 × 6행 = 60행, 25행 페이지 → 여러 브레이크.

    누적: 6,12,18,24 → idx=4에서 24+6=30>25 → brk[4], acc=6
         6,12,18,24 → idx=8에서 24+6=30>25 → brk[8], acc=6
         마지막 블록 (idx=9): 6+6=12 ≤ 25 → 추가 브레이크 없음
    """
    blocks = [_mk_block(f"B{i}", 5) for i in range(10)]
    _assert("multi_page_breaks", _plan_page_breaks(blocks, 25), [4, 8])


def main() -> int:
    tests = [
        test_no_break_under_threshold,
        test_break_at_overflow,
        test_empty_blocks,
        test_single_large_block,
        test_rows_per_page_constant,
        test_zero_or_negative_rows_per_page,
        test_multi_page_breaks,
    ]
    failed = 0
    for t in tests:
        print(f"\n[{t.__name__}]")
        try:
            t()
        except AssertionError as e:
            print(f"  FAIL: {e}")
            failed += 1
        except Exception as e:
            print(f"  ERROR: {type(e).__name__}: {e}")
            failed += 1

    print(f"\n{'=' * 50}")
    if failed:
        print(f"FAILED: {failed}/{len(tests)} tests")
        return 1
    print(f"OK ({len(tests)} tests passed)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
