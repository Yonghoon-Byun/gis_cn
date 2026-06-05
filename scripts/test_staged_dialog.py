# -*- coding: utf-8 -*-
"""골든 테스트 — dialog 3단계 export 데이터 경로 (한컴/QGIS 불필요).

dialog._export_staged 가 구동하는 계약을 Qt 없이 검증한다:
  calculate_results(단계별) → build_analysis_result → assemble_staged_report
  → cn_delta_rows 에 적정성 사유 주입 → render_staged_report.

result_calculator 는 qgis/pandas 를 import 하므로, 호출하지 않는 qgis 심볼만 stub 한다
(assemble_staged_report·build_cn_delta·build_analysis_result 는 런타임에 qgis 미사용).
pandas 는 실제 설치본 사용.

실행: python scripts/test_staged_dialog.py
"""
from __future__ import annotations

import importlib.util
import sys
import tempfile
import types
import zipfile
from pathlib import Path

from lxml import etree

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

ROOT = Path(__file__).resolve().parent.parent
TEMPLATE = ROOT / "gis_cn" / "templates" / "v1.0" / "cn_report.hwpx"


def _stub_qgis():
    """result_calculator import 가 요구하는 qgis 심볼만 최소 stub."""
    qgis = types.ModuleType("qgis"); qgis.__path__ = []
    core = types.ModuleType("qgis.core")
    core.QgsVectorLayer = object
    pyqt = types.ModuleType("qgis.PyQt"); pyqt.__path__ = []
    qtcore = types.ModuleType("qgis.PyQt.QtCore")
    class _QVariant:  # _is_null 의 isinstance 검사용(런타임 미사용)
        def isNull(self):
            return False
    qtcore.QVariant = _QVariant
    sys.modules.update({
        "qgis": qgis, "qgis.core": core,
        "qgis.PyQt": pyqt, "qgis.PyQt.QtCore": qtcore,
    })


def _load_modules():
    _stub_qgis()
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
    rc = load("gis_cn.core.result_calculator", "gis_cn/core/result_calculator.py")
    return ar, hw, rc


ar, hw, rc = _load_modules()


def _hp(root):
    ns = root.nsmap.get("hp")
    return lambda t: f"{{{ns}}}{t}"


def _parse(path):
    root = etree.fromstring(zipfile.ZipFile(path).read("Contents/section0.xml"),
                            etree.XMLParser(huge_tree=True))
    return root, _hp(root)


def _all_texts(root, hp):
    return [t.text for t in root.iter(hp("t")) if t.text]


def _assert(cond, msg):
    if not cond:
        raise AssertionError(msg)


def _stage_data(blocks):
    """blocks: [(ws, amc3_ws, [(lu, amc3_lu), ...]), ...] → (r1, r2) (calculate_results 형식)."""
    r1, r2 = [], []
    for ws, amc3, lus in blocks:
        rows = [{
            'land_use': lu, 'A_area': 100.0, 'A_cn': 79,
            'B_area': None, 'B_cn': None, 'C_area': None, 'C_cn': None,
            'D_area': None, 'D_cn': None, 'total_area': 100.0,
            'amc2_cn': 76.0, 'amc3_cn': a3,
        } for lu, a3 in lus]
        r1.append({
            'watershed': ws, 'rows': rows,
            'total_A': 100.0, 'total_B': 0.0, 'total_C': 0.0, 'total_D': 0.0,
            'total_area': 100.0, 'amc2_cn': 76.0, 'amc3_cn': amc3,
        })
        r2.append({'watershed': ws, 'total_area': 100.0, 'amc2_cn': 76.0, 'amc3_cn': amc3})
    return r1, r2


def _ares(blocks):
    r1, r2 = _stage_data(blocks)
    return rc.build_analysis_result(r1, r2, meta=ar.ProjectMeta())


def test_assemble_delta():
    """단계별 AnalysisResult → assemble_staged_report 의 증감 계산 정확도."""
    pre = _ares([("GR1", 86.52, [("논", 88.41)]), ("GR2", 85.96, [("밭", 87.0)])])
    during = _ares([("GR1", 87.14, [("논", 88.41)]), ("GR2", 87.90, [("밭", 87.0)])])
    post = _ares([("GR1", 88.30, [("논", 88.41)])])          # GR2 개발후 편입(없음)

    report = rc.assemble_staged_report(
        {ar.STAGE_PRE: pre, ar.STAGE_DURING: during, ar.STAGE_POST: post},
        meta=ar.ProjectMeta(project_name="3단계검증"))

    _assert(report.stage_order == [ar.STAGE_PRE, ar.STAGE_DURING, ar.STAGE_POST],
            f"stage_order 오류: {report.stage_order}")
    by = {r.watershed: r for r in report.cn_delta_rows}
    _assert(set(by) == {"GR1", "GR2"}, f"소유역 집합 오류: {set(by)}")
    g1 = by["GR1"]
    _assert(abs(g1.delta_during - 0.62) < 1e-9, f"GR1 증감(중) 오류: {g1.delta_during}")
    _assert(abs(g1.delta_post - 1.78) < 1e-9, f"GR1 증감(후) 오류: {g1.delta_post}")
    g2 = by["GR2"]
    _assert(abs(g2.delta_during - 1.94) < 1e-9, f"GR2 증감(중) 오류: {g2.delta_during}")
    _assert(g2.cn_post is None and g2.delta_post is None, "GR2 개발후 None 이어야(편입)")
    print("  ✓ test_assemble_delta: 증감 ②-①/③-① 정확, 편입 소유역 None")
    return report


def test_reason_injection_render(tmp: Path):
    """사유 주입(dialog 시뮬) → render_staged_report 가 표4-17 에 사유·증감 출력."""
    report = test_assemble_delta()
    # dialog._export_staged: 사용자 입력 사유를 cn_delta_rows 에 주입
    inj = {"GR1": ("불투수 면적 증가로 CN 상승", "개발 완료 후 추가 상승"),
           "GR2": ("개발중 편입 예정", "")}
    for row in report.cn_delta_rows:
        rd = inj.get(row.watershed)
        if rd:
            row.reason_during, row.reason_post = rd

    out = tmp / "staged_dialog.hwpx"
    hw.render_staged_report(report, TEMPLATE, out)

    _assert(zipfile.ZipFile(out).testzip() is None, "zip 무결성 실패")
    root, hp = _parse(out)
    txt = _all_texts(root, hp)
    for v in ("GR1", "86.52", "증) 0.62", "증) 1.78",
              "불투수 면적 증가로 CN 상승", "개발 완료 후 추가 상승", "개발중 편입 예정"):
        _assert(v in txt, f"표4-17 주입 누락: {v}")
    fids = [fb.get("id") for fb in root.iter(hp("fieldBegin"))]
    _assert(len(fids) == len(set(fids)), "fieldBegin id 중복")
    print("  ✓ test_reason_injection_render: 적정성 사유·증감 표기 hwpx 주입, id 유니크")


def main() -> int:
    print(f"3단계 dialog 데이터경로 골든 테스트 (template={TEMPLATE.name})")
    with tempfile.TemporaryDirectory() as d:
        try:
            test_assemble_delta()
            test_reason_injection_render(Path(d))
        except AssertionError as e:
            print(f"  ✗ FAIL: {e}")
            return 1
    print("전부 통과 ✓ (3단계 조립·사유주입·렌더 데이터경로 회귀 없음 — Qt UI 는 한글 육안+QGIS 게이트)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
