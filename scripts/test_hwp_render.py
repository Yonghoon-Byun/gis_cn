"""hwp_renderer.py 엔드투엔드 테스트 (QGIS Python 사용).

샘플 AnalysisResult 를 만들어 렌더 → 출력 HWP 재열어 필드 치환 확인.
"""
import sys
from pathlib import Path
from datetime import date

# 플러그인 소스 경로 추가
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from gis_cn.core.analysis_result import (
    AnalysisResult, ProjectMeta, WatershedBlock, WatershedSummary,
    LandUseRow, CnReferenceRow,
)
from gis_cn.core.hwp_renderer import render_hwp

TPL = ROOT / "gis_cn" / "templates" / "v1.0" / "cn_report.hwpx"
OUT = Path(r"C:\temp\test_render.hwpx")


def main():
    # 샘플 데이터
    meta = ProjectMeta(
        project_name="테스트 재해영향평가",
        site_name="양주시 테스트지구",
        author="홍길동",
        analysis_date=date(2026, 4, 22),
        development_stage="현 상태",
    )

    cn_ref = [
        CnReferenceRow(land_use="논",    a=72, b=81, c=88, d=91),
        CnReferenceRow(land_use="밭",    a=67, b=78, c=85, d=89),
        CnReferenceRow(land_use="임야",  a=36, b=60, c=73, d=79),
        CnReferenceRow(land_use="주거지", a=77, b=85, c=90, d=92),
    ]

    def mkrow(lu, aa, ac, ba, bc, ca, cc, da, dc, tot, a2, a3):
        return LandUseRow(
            land_use=lu,
            a_area=aa, a_cn=ac, b_area=ba, b_cn=bc,
            c_area=ca, c_cn=cc, d_area=da, d_cn=dc,
            total_area=tot, amc2_cn=a2, amc3_cn=a3,
        )

    block_a = WatershedBlock(
        name="WS1",
        rows=[
            mkrow("논",   1000, 81, 200, 88, None, None, None, None, 1200, 82.17, 91),
            mkrow("임야", 500, 60, 300, 73, 100, 79, None, None, 900, 66.89, 82),
        ],
        total_a=1500, total_b=500, total_c=100, total_d=0,
        total_area=2100, amc2_cn=75.62, amc3_cn=87.5,
    )
    block_b = WatershedBlock(
        name="WS2",
        rows=[
            mkrow("주거지", 800, 85, None, None, None, None, None, None, 800, 85.0, 92),
        ],
        total_a=800, total_b=0, total_c=0, total_d=0,
        total_area=800, amc2_cn=85.0, amc3_cn=92.0,
    )

    summary = [
        WatershedSummary(name="WS1", total_area=2100, amc2_cn=75.62, amc3_cn=87.5),
        WatershedSummary(name="WS2", total_area=800,  amc2_cn=85.0,  amc3_cn=92.0),
    ]

    result = AnalysisResult(
        meta=meta,
        cn_reference=cn_ref,
        detail_blocks=[block_a, block_b],
        summary_rows=summary,
        notes="테스트 렌더 비고",
    )

    OUT.parent.mkdir(parents=True, exist_ok=True)
    if OUT.exists():
        OUT.unlink()

    print(f"rendering → {OUT}")
    render_hwp(result, TPL, OUT)
    print(f"saved: {OUT.stat().st_size:,} bytes")

    # 재열어 필드 값 검증
    from pyhwpx import Hwp
    h = Hwp(visible=False)
    h.open(str(OUT))
    checks = [
        ("meta.project_name", None, "테스트 재해영향평가"),
        ("meta.site_name",    None, "양주시 테스트지구"),
        ("ref.lu",  0, "논"),
        ("ref.lu",  1, "밭"),
        ("ref.cn_a", 0, "72"),
        ("res.ws",  0, "WS1"),
        ("res.lu",  0, "논"),
        ("res.amc2_cn", 0, "82.17"),
        ("res.ws",  1, "WS1"),
        ("res.lu",  1, "임야"),
        ("res.ws",  3, "WS2"),
        ("res.lu",  3, "주거지"),
        ("sum.ws",   0, "WS1"),
        ("sum.amc2", 0, "75.62"),
        ("sum.ws",   1, "WS2"),
    ]
    ok = 0
    fail = []
    for name, idx, expected in checks:
        try:
            if idx is None:
                v = h.get_field_text(name)
            else:
                v = h.get_field_text(name, idx)
        except Exception as e:
            fail.append((name, idx, f"ERR {e}"))
            continue
        v = str(v).strip()
        if v == expected:
            ok += 1
        else:
            fail.append((name, idx, f"expected={expected!r} got={v!r}"))
    h.quit()
    print(f"\n{ok}/{len(checks)} checks passed")
    for nm, idx, msg in fail:
        print(f"  FAIL {nm}[{idx}]: {msg}")

if __name__ == "__main__":
    main()
