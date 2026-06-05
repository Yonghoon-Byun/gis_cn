# -*- coding: utf-8 -*-
"""한글 육안 검증용 샘플 결과 렌더 — 3단계 비교 보고서를 Desktop 에 출력.

build_template/test 와 동일한 모듈 로더로 qgis/pandas 없이 렌더한다.
사용: python scripts/render_sample.py [출력경로]
"""
from __future__ import annotations

import importlib.util
import sys
import types
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

ROOT = Path(__file__).resolve().parent.parent
TEMPLATE = ROOT / "gis_cn" / "templates" / "v1.0" / "cn_report.hwpx"


def _load():
    g = types.ModuleType("gis_cn"); g.__path__ = [str(ROOT / "gis_cn")]
    gc = types.ModuleType("gis_cn.core"); gc.__path__ = [str(ROOT / "gis_cn" / "core")]
    sys.modules["gis_cn"] = g
    sys.modules["gis_cn.core"] = gc

    def load(name, rel):
        spec = importlib.util.spec_from_file_location(name, ROOT / rel)
        m = importlib.util.module_from_spec(spec)
        m.__package__ = "gis_cn.core"
        sys.modules[name] = m
        spec.loader.exec_module(m)
        return m

    ar = load("gis_cn.core.analysis_result", "gis_cn/core/analysis_result.py")
    hw = load("gis_cn.core.hwpx_writer", "gis_cn/core/hwpx_writer.py")
    return ar, hw


def _lur(ar, lu, aa, acn, ta, a2, a3):
    return ar.LandUseRow(land_use=lu, a_area=aa, a_cn=acn, b_area=None, b_cn=None,
                         c_area=None, c_cn=None, d_area=None, d_cn=None,
                         total_area=ta, amc2_cn=a2, amc3_cn=a3)


def main():
    ar, hw = _load()
    out = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(r"C:/Users/admin/Desktop/results_secfix.hwpx")

    delta = [
        ar.CnDeltaRow(watershed="GR1", cn_pre=86.52, cn_during=87.14, delta_during=0.62,
                      reason_during="불투수면 증가", cn_post=88.30, delta_post=1.78,
                      reason_post="포장면적 확대"),
        ar.CnDeltaRow(watershed="GR2", cn_pre=85.96, cn_during=87.90, delta_during=1.94,
                      reason_during="공사 중 나지", cn_post=None, delta_post=None,
                      reason_post="개발후 GR1 편입"),
    ]
    pre = ar.AnalysisResult(detail_blocks=[ar.WatershedBlock(
        name="GR1",
        rows=[_lur(ar, "논", 10000.0, 79, 12345.678, 76.83, 88.41),
              _lur(ar, "밭", 1000.0, 63, 2345.678, 75.0, 87.0)],
        total_a=11000.0, total_area=14691.356, amc2_cn=76.0, amc3_cn=86.52)])
    report = ar.StagedReport(
        meta=ar.ProjectMeta(project_name="양주장흥 재해영향평가(샘플)"),
        stages={ar.STAGE_PRE: pre, ar.STAGE_DURING: ar.AnalysisResult(),
                ar.STAGE_POST: ar.AnalysisResult()},
        stage_order=ar.DEFAULT_STAGE_ORDER,
        cn_delta_rows=delta,
    )
    hw.render_staged_report(report, TEMPLATE, out)
    print(f"렌더 완료: {out}")


if __name__ == "__main__":
    main()
