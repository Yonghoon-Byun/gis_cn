# -*- coding: utf-8 -*-
"""원본 hwpx의 CN 표 스타일과 생성 hwpx(또는 템플릿)를 비교 검증.

원본(authentic 스타일)을 기준으로 표4-17/4-18/4-19의 **쪽 여백·셀 테두리·글꼴**이
생성물에서 충실히 재현됐는지 확인한다. 핵심: borderFill/charPr 을 ID 가 아니라 각 문서의
header.xml 정의로 **해석(resolve)** 한 뒤 비교한다(같은 ID라도 문서마다 정의가 다르므로).

사용: python scripts/compare_to_original.py [생성물.hwpx]
      (생략 시 templates/v1.0/cn_report.hwpx 의 빈 템플릿과 비교)
"""
from __future__ import annotations

import sys
import zipfile
from pathlib import Path

from lxml import etree

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

ROOT = Path(__file__).resolve().parent.parent
ORIGINAL = Path(r"D:\DATA\카카오톡 받은 파일\제4장 재해영향 예측 및 평가_최종.hwpx")
TEMPLATE = ROOT / "gis_cn" / "templates" / "v1.0" / "cn_report.hwpx"

# 비교할 3개 CN 표 (caption 키워드, colCnt)
TARGETS = [
    ("표4-17 CN변화", ("유출곡선지수", "변화"), "8"),
    ("표4-18 기준표", ("AMC-II",), "8"),
    ("표4-19 산정결과", ("산정결과",), "14"),
]
PARSER = etree.XMLParser(huge_tree=True, recover=True)
BORDER_SIDES = ("leftBorder", "rightBorder", "topBorder", "bottomBorder")


def _nsmap(root):
    hp = root.nsmap.get("hp")
    hh = next((v for k, v in root.nsmap.items() if k == "hh"), None)
    return hp, hh


class Doc:
    def __init__(self, path: Path):
        z = zipfile.ZipFile(path)
        self.sections = [etree.fromstring(z.read(n), PARSER)
                         for n in sorted(z.namelist()) if "section" in n and n.endswith(".xml")]
        self.header = etree.fromstring(z.read("Contents/header.xml"), PARSER)
        self.hp = self.sections[0].nsmap.get("hp")
        self.hh = next((v for k, v in self.header.nsmap.items() if k == "hh"), None)
        self.bf = self._index_borderfills()
        self.cp = self._index_charprs()

    def H(self, t):  # header(hh) tag
        return f"{{{self.hh}}}{t}"

    def P(self, t):  # paragraph(hp) tag
        return f"{{{self.hp}}}{t}"

    def _index_borderfills(self) -> dict:
        out = {}
        for bf in self.header.iter(self.H("borderFill")):
            sides = {}
            for side in BORDER_SIDES:
                el = bf.find(self.H(side))
                if el is not None:
                    sides[side] = (el.get("type"), el.get("width"), el.get("color"))
                else:
                    sides[side] = None
            fill = bf.find(".//" + self.H("fillBrush"))
            face = bf.find(".//" + self.H("faceColor"))
            out[bf.get("id")] = {"sides": sides, "fill": face.get("color") if face is not None else None}
        return out

    def _index_charprs(self) -> dict:
        out = {}
        for cp in self.header.iter(self.H("charPr")):
            fr = cp.find(self.H("fontRef"))
            out[cp.get("id")] = {
                "height": cp.get("height"), "bold": cp.find(self.H("bold")) is not None,
                "font": fr.get("hangul") if fr is not None else None,
            }
        return out

    def find_table(self, keywords, colcnt):
        for sec in self.sections:
            for tbl in sec.iter(self.P("tbl")):
                cap_el = tbl.find(self.P("caption"))
                cap = "".join(x.text or "" for x in (cap_el if cap_el is not None else tbl).iter(self.P("t")))
                if tbl.get("colCnt") == colcnt and all(k in cap for k in keywords):
                    return tbl
        return None

    def page_margin_mm(self):
        pp = None
        for sec in self.sections:
            pp = sec.find(".//" + self.P("pagePr"))
            if pp is not None:
                break
        if pp is None:
            return None
        mg = pp.find(self.P("margin"))
        u = lambda v: round(int(v) / 7200 * 25.4, 1) if v else None
        return {k: u(mg.get(k)) for k in ("top", "bottom", "left", "right")} if mg is not None else None

    def secpr_placement(self):
        """secPr 가 어떻게 배치됐는지 — (전용run?, 같은run에 표?, 첫단락?) 반환.

        한글이 쪽 여백을 적용하려면 secPr 가 '표 없는 전용 run'에 있어야 한다.
        같은 run 에 tbl 이 섞이면 섹션정의를 못 읽어 전 페이지가 기본 여백으로 떨어진다.
        """
        for sec in self.sections:
            sp = sec.find(".//" + self.P("secPr"))
            if sp is None:
                continue
            run = sp.getparent()
            if run is None or run.tag != self.P("run"):
                return {"found": True, "own_run": False, "tbl_in_run": False, "first_para": False}
            tbl_in_run = run.find(self.P("tbl")) is not None
            # secPr 와 같은 run 에 colPr(단 설정) 컨트롤이 있는지 — 없으면 한글이
            # 섹션정의를 폐기하고 기본 여백으로 떨어진다(여백 무시의 직접 원인).
            colpr_in_run = any(
                c.tag == self.P("ctrl") and c.find(self.P("colPr")) is not None
                for c in run)
            # secPr 단락이 섹션의 첫 단락인지
            para = next((a for a in sp.iterancestors() if a.tag == self.P("p")), None)
            paras = list(sec.iter(self.P("p")))
            first_para = bool(paras) and para is paras[0]
            return {"found": True, "own_run": not tbl_in_run,
                    "tbl_in_run": tbl_in_run, "colpr_in_run": colpr_in_run,
                    "first_para": first_para}
        return {"found": False}

    def table_border_profile(self, tbl):
        """표의 모든 셀이 쓰는 borderFill 을 해석해 (테두리 4면, 채움) 멀티셋으로."""
        from collections import Counter
        prof = Counter()
        for tc in tbl.iter(self.P("tc")):
            bid = tc.get("borderFillIDRef")
            resolved = self.bf.get(bid)
            if resolved is None:
                prof[("MISSING", bid)] += 1
                continue
            sig = tuple((resolved["sides"][s][0] if resolved["sides"][s] else "NONE") for s in BORDER_SIDES)
            prof[(sig, resolved["fill"])] += 1
        return prof


def _border_kinds(prof):
    """프로파일에서 등장하는 테두리 선종류 집합."""
    kinds = set()
    for (sig, _fill), _n in prof.items():
        if sig == "MISSING":
            kinds.add("MISSING")
        else:
            kinds.update(sig)
    return kinds


def main() -> int:
    out_path = Path(sys.argv[1]) if len(sys.argv) > 1 else TEMPLATE
    if not ORIGINAL.exists():
        print(f"원본 없음: {ORIGINAL}")
        return 2
    print(f"원본:   {ORIGINAL.name}")
    print(f"비교물: {out_path}")
    orig, mine = Doc(ORIGINAL), Doc(out_path)

    print("\n=== 쪽 여백 (mm) ===")
    print(f"  원본:   {orig.page_margin_mm()}")
    print(f"  비교물: {mine.page_margin_mm()}")

    problems = 0
    if orig.page_margin_mm() != mine.page_margin_mm():
        print("   ★불일치★ 여백 값이 다름")
        problems += 1

    print("\n=== secPr 배치 (쪽 여백 적용 가능 여부) ===")
    op, mp = orig.secpr_placement(), mine.secpr_placement()
    print(f"  원본:   {op}")
    print(f"  비교물: {mp}")
    if mp.get("tbl_in_run"):
        print("   ★불일치★ secPr 가 표와 같은 run 에 있음 → 한글이 여백을 무시하고 "
              "전 페이지가 기본 여백으로 떨어짐(원본은 secPr 전용 run)")
        problems += 1
    elif not mp.get("colpr_in_run"):
        print("   ★불일치★ secPr-run 에 colPr(단 설정) 컨트롤 없음 → 한글이 섹션정의를 "
              "폐기하고 기본 여백으로 떨어짐(원본은 [secPr, ctrl(colPr)])")
        problems += 1
    elif not mp.get("own_run") or not mp.get("first_para"):
        print("   ⚠ secPr 가 전용 run/첫 단락이 아님 — 한글 여백 적용이 불안정할 수 있음")
        problems += 1

    print("\n=== 표별 테두리 스타일 (borderFill 해석 비교) ===")
    for name, kw, cc in TARGETS:
        ot = orig.find_table(kw, cc)
        mt = mine.find_table(kw, cc)
        if ot is None or mt is None:
            print(f"\n[{name}] 원본={ot is not None} 비교물={mt is not None} — 표를 못 찾음")
            problems += 1
            continue
        op, mp = orig.table_border_profile(ot), mine.table_border_profile(mt)
        ok_kinds, mk_kinds = _border_kinds(op), _border_kinds(mp)
        miss = mk_kinds - ok_kinds            # 비교물에만 있는 선종류(원본에 없음)
        lost = ok_kinds - mk_kinds            # 원본엔 있으나 비교물에 없음
        has_missing = any(k == "MISSING" for k in mk_kinds)
        status = "✓ 일치" if (not miss and not lost and not has_missing) else "★불일치★"
        print(f"\n[{name}] 셀 {sum(mp.values())}개  {status}")
        print(f"   원본 선종류:   {sorted(ok_kinds)}")
        print(f"   비교물 선종류: {sorted(mk_kinds)}")
        if has_missing:
            print(f"   ⚠ borderFill 정의 없음(테두리 사라짐→빨간 점선): {[k for k in mk_kinds if k=='MISSING']}")
        if miss or lost:
            print(f"   ⚠ 원본대비 추가:{sorted(miss)} 누락:{sorted(lost)}")
        if status.startswith("★"):
            problems += 1

    print("\n" + ("=" * 60))
    print(f"결과: {'스타일 불일치 ' + str(problems) + '건 — 원본 스타일 충실 이식 필요' if problems else '원본과 스타일 일치 ✓'}")
    return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(main())
