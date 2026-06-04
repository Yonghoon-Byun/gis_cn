# -*- coding: utf-8 -*-
"""골든 테스트 — 순수 Python HWPX 렌더러 구조 회귀 방지 (한컴/QGIS 불필요).

ralplan 원칙: 골든 테스트는 '한글 렌더 보증'이 아니라 '구조 회귀 방지'다.
시각 충실도는 별도 한글 육안 게이트로 확인(자동화 불가). 여기서는 XML/필드/행수/
id-유니크/그리드 타일링 등 구조만 검증한다. lxml만 필요.

실행: python scripts/test_hwpx_writer.py
"""
from __future__ import annotations

import importlib.util
import sys
import tempfile
import types
import zipfile
import zlib  # noqa: F401  (zip 무결성 의존)
from pathlib import Path

from lxml import etree

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

ROOT = Path(__file__).resolve().parent.parent
TEMPLATE = ROOT / "gis_cn" / "templates" / "v1.0" / "cn_report.hwpx"


def _load_modules():
    """qgis/pandas 의존 없이 analysis_result + hwpx_writer 직접 로드."""
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


ar, hw = _load_modules()


def _hp(root):
    ns = root.nsmap.get("hp")
    return lambda t: f"{{{ns}}}{t}"


def _parse(path):
    root = etree.fromstring(zipfile.ZipFile(path).read("Contents/section0.xml"),
                            etree.XMLParser(huge_tree=True))
    return root, _hp(root)


def _data_rows(root, hp, prefix):
    tbl = next((t for t in root.iter(hp("tbl"))
                if any((fb.get("name") or "").startswith(prefix + ".") for fb in t.iter(hp("fieldBegin")))), None)
    if tbl is None:
        return 0
    return sum(1 for tr in tbl.findall(hp("tr"))
               if any((fb.get("name") or "").startswith(prefix + ".") for fb in tr.iter(hp("fieldBegin"))))


def _all_texts(root, hp):
    return [t.text for t in root.iter(hp("t")) if t.text]


def _assert(cond, msg):
    if not cond:
        raise AssertionError(msg)


def _lur(lu, aa, acn, ta, a2, a3):
    return ar.LandUseRow(land_use=lu, a_area=aa, a_cn=acn, b_area=None, b_cn=None,
                         c_area=None, c_cn=None, d_area=None, d_cn=None,
                         total_area=ta, amc2_cn=a2, amc3_cn=a3)


def test_single_stage(tmp: Path) -> None:
    blocks = [ar.WatershedBlock(name="GR1",
                                rows=[_lur("논", 10000.0, 79, 12345.678, 76.83, 88.41),
                                      _lur("밭", 1000.0, 63, 2345.678, 75.0, 87.0)],
                                total_a=11000.0, total_area=14691.356, amc2_cn=76.0, amc3_cn=86.52)]
    result = ar.AnalysisResult(meta=ar.ProjectMeta(project_name="단일단계"), detail_blocks=blocks)
    out = tmp / "single.hwpx"
    hw.render_hwpx(result, TEMPLATE, out)            # 내부 _validate 실패 시 예외

    _assert(zipfile.ZipFile(out).testzip() is None, "zip 무결성 실패")
    root, hp = _parse(out)
    _assert(_data_rows(root, hp, "res") == 3, "res 데이터행 != 3 (2 토지이용 + 1 합계)")
    fids = [fb.get("id") for fb in root.iter(hp("fieldBegin"))]
    _assert(len(fids) == len(set(fids)), "fieldBegin id 중복")
    txt = _all_texts(root, hp)
    for v in ("논", "12,345.678", "합계", "88.41"):
        _assert(v in txt, f"주입값 누락: {v}")
    print("  ✓ test_single_stage: res 3행, 값 주입, id 유니크, zip 무결")


def test_staged(tmp: Path) -> None:
    delta = [
        ar.CnDeltaRow(watershed="GR1", cn_pre=86.52, cn_during=87.14, delta_during=0.62,
                      cn_post=88.30, delta_post=1.78),
        ar.CnDeltaRow(watershed="GR2", cn_pre=85.96, cn_during=87.90, delta_during=1.94,
                      cn_post=None, delta_post=None),  # 개발후 편입
    ]
    pre = ar.AnalysisResult(detail_blocks=[ar.WatershedBlock(
        name="GR1", rows=[_lur("논", 10000.0, 79, 12345.678, 76.83, 88.41)],
        total_a=10000.0, total_area=12345.678, amc2_cn=76.0, amc3_cn=86.52)])
    report = ar.StagedReport(
        meta=ar.ProjectMeta(project_name="3단계"),
        stages={ar.STAGE_PRE: pre, ar.STAGE_DURING: ar.AnalysisResult(),
                ar.STAGE_POST: ar.AnalysisResult()},
        stage_order=ar.DEFAULT_STAGE_ORDER,
        cn_delta_rows=delta,
    )
    out = tmp / "staged.hwpx"
    hw.render_staged_report(report, TEMPLATE, out)

    _assert(zipfile.ZipFile(out).testzip() is None, "zip 무결성 실패")
    root, hp = _parse(out)
    ntbl = len(list(root.iter(hp("tbl"))))
    _assert(ntbl == 3, f"표 개수 != 3 (표4-17/18/19): {ntbl}")
    _assert(_data_rows(root, hp, "delta") == 2, "표4-17 행 != 2")
    _assert(_data_rows(root, hp, "res") == 2, "표4-19 행 != 2 (GR1 토지이용 1행 + 합계 1행)")
    txt = _all_texts(root, hp)
    for v in ("GR1", "86.52", "증) 0.62", "88.30", "증) 1.78"):
        _assert(v in txt, f"표4-17 값 누락: {v}")
    fids = [fb.get("id") for fb in root.iter(hp("fieldBegin"))]
    _assert(len(fids) == len(set(fids)), "fieldBegin id 중복")
    print("  ✓ test_staged: 표 3개, 표4-17 2행(증감 표기), id 유니크")


def test_increment_format() -> None:
    _assert(hw._fmt_increment(0.62) == "증) 0.62", "증가 표기 오류")
    _assert(hw._fmt_increment(-1.22) == "감) 1.22", "감소 표기 오류")
    _assert(hw._fmt_increment(0.0) == "-", "변화없음 표기 오류")
    _assert(hw._fmt_increment(None) == "-", "None 표기 오류")
    print("  ✓ test_increment_format: 증)/감)/- 표기")


def main() -> int:
    print(f"골든 테스트 (template={TEMPLATE.name})")
    with tempfile.TemporaryDirectory() as d:
        tmp = Path(d)
        try:
            test_increment_format()
            test_single_stage(tmp)
            test_staged(tmp)
        except AssertionError as e:
            print(f"  ✗ FAIL: {e}")
            return 1
    print("전부 통과 ✓ (구조 회귀 없음 — 시각 충실도는 한글 육안 게이트에서 별도 확인)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
