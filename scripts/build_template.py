# -*- coding: utf-8 -*-
"""F2: 클린 HWPX 템플릿 빌더 (한컴오피스/COM 불필요, lxml only).

현 cn_report.hwpx의 깨진 부분(852cb76 '알려진 한계')을 정리한다:
  · 빈 "(계속)" res 쓰레기 표 5개 제거 (조부모 <hp:p> 통째)
  · res 산정결과표를 헤더 + 프로토타입 데이터 1행으로 축소(런타임에 동적 복제)
  · 원본 재해영향평가 보고서 쪽 여백(사방 15mm / 머리·꼬리 12.5mm) 적용
  · 전체 행 rowAddr 위치순 재번호 + 검증(한글 크래시 회피)

후속(F2b): ref 8열 국가표준 기준표(표4-18)·meta 표지·표4-17 증감 비교표 추가.

사용: python scripts/build_template.py [출력경로]   (생략 시 템플릿 덮어쓰기)
"""
from __future__ import annotations

import sys
import zipfile
from copy import deepcopy
from pathlib import Path

from lxml import etree

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

ROOT = Path(__file__).resolve().parent.parent
TEMPLATE = ROOT / "gis_cn" / "templates" / "v1.0" / "cn_report.hwpx"
SECTION = "Contents/section0.xml"
ORIG_MARGIN = {"top": "4252", "bottom": "4252", "left": "4252", "right": "4252",
               "header": "3543", "footer": "3543", "gutter": "0"}
RES_PREFIX = "res"


def _ns(root):
    ns = root.nsmap.get("hp")
    return ns, (lambda t: f"{{{ns}}}{t}")


def _res_field_count(tbl, hp) -> int:
    return sum(1 for fb in tbl.iter(hp("fieldBegin"))
              if (fb.get("name") or "").startswith(RES_PREFIX + "."))


def clean_template(src: Path, out: Path) -> dict:
    zin = zipfile.ZipFile(src)
    names = zin.namelist()
    raw = {n: zin.read(n) for n in names}
    zin.close()
    root = etree.fromstring(raw[SECTION], etree.XMLParser(huge_tree=True))
    ns, hp = _ns(root)
    stats = {"junk_removed": 0, "res_rows_before": 0, "res_rows_after": 0, "margins": 0}

    # 1) 쓰레기 res 표(필드 0) 제거 — 조부모 <hp:p> 통째
    tables = list(root.iter(hp("tbl")))
    res_tables = [t for t in tables
                  if "산정결과" in "".join(x.text or "" for x in
                                       (t.find(hp("caption")) if t.find(hp("caption")) is not None else t).iter(hp("t")))]
    real_res = next((t for t in res_tables if _res_field_count(t, hp) > 0), None)
    if real_res is None:
        raise RuntimeError("res.* 필드를 가진 산정결과표를 찾지 못함")
    for t in res_tables:
        if t is real_res:
            continue
        p = t.getparent().getparent()      # tbl -> run -> p
        p.getparent().remove(p)
        stats["junk_removed"] += 1

    # 2) real res 표: 헤더 + 프로토타입 1행으로 축소
    trs = real_res.findall(hp("tr"))
    stats["res_rows_before"] = len(trs)

    def is_data(tr):
        return any((fb.get("name") or "").startswith(RES_PREFIX + ".")
                   for fb in tr.iter(hp("fieldBegin")))

    first_data = next(i for i, tr in enumerate(trs) if is_data(tr))
    prototype = deepcopy(trs[first_data])
    for tr in trs[first_data:]:            # 데이터영역(끝 빈 행 포함) 제거
        real_res.remove(tr)
    if first_data > 0:
        trs[first_data - 1].addnext(prototype)   # 마지막 헤더 행 뒤에 프로토타입 삽입
    else:
        real_res.append(prototype)

    # 3) rowAddr 위치순 재번호 + rowCnt
    all_rows = real_res.findall(hp("tr"))
    for pos, tr in enumerate(all_rows):
        for ca in tr.iter(hp("cellAddr")):
            ca.set("rowAddr", str(pos))
    real_res.set("rowCnt", str(len(all_rows)))
    stats["res_rows_after"] = len(all_rows)

    # 4) 원본 여백
    for mg in root.iter(hp("margin")):
        if mg.getparent().tag == hp("pagePr"):
            for k, v in ORIG_MARGIN.items():
                mg.set(k, v)
            stats["margins"] += 1

    # 5) 검증 — 모든 표 rowAddr 범위·연속, 그리드 타일링
    _validate(root, hp)

    # 6) 저장
    raw[SECTION] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    out = Path(out)
    if out.exists():
        out.unlink()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        zi = zipfile.ZipInfo("mimetype")
        zi.compress_type = zipfile.ZIP_STORED
        zout.writestr(zi, raw.get("mimetype", b"application/hwp+zip"))
        for n in names:
            if n != "mimetype":
                zout.writestr(n, raw[n])
    return stats


def _validate(root, hp) -> None:
    problems = []
    for ti, tbl in enumerate(root.iter(hp("tbl"))):
        rc, cc = int(tbl.get("rowCnt")), int(tbl.get("colCnt"))
        grid = {}
        over = False
        for tr in tbl.findall(hp("tr")):
            for tc in tr.findall(hp("tc")):
                ca, sp = tc.find(hp("cellAddr")), tc.find(hp("cellSpan"))
                r, c = int(ca.get("rowAddr")), int(ca.get("colAddr"))
                rs = int(sp.get("rowSpan")) if sp is not None else 1
                cs = int(sp.get("colSpan")) if sp is not None else 1
                if r >= rc or c >= cc:
                    problems.append(f"표{ti} cellAddr({r},{c}) 범위초과 {rc}x{cc}"); over = True
                for dr in range(rs):
                    for dc in range(cs):
                        k = (r + dr, c + dc)
                        if k in grid:
                            problems.append(f"표{ti} 셀 중복 {k}")
                        grid[k] = 1
        if not over:
            miss = [(r, c) for r in range(rc) for c in range(cc) if (r, c) not in grid]
            if miss:
                problems.append(f"표{ti} 그리드 미충족 {len(miss)}칸 예{miss[:3]}")
    if problems:
        raise RuntimeError("검증 실패(한글 크래시 위험): " + "; ".join(problems[:8]))


def main() -> int:
    out = Path(sys.argv[1]) if len(sys.argv) > 1 else TEMPLATE
    stats = clean_template(TEMPLATE, out)
    print(f"클린 템플릿 생성: {out}")
    print(f"  쓰레기 표 제거: {stats['junk_removed']}개")
    print(f"  res 행: {stats['res_rows_before']} → {stats['res_rows_after']} (헤더 + 프로토타입 1)")
    print(f"  여백 적용: pagePr margin {stats['margins']}개 (사방 15mm)")
    print("  검증 통과 ✓")
    return 0


if __name__ == "__main__":
    sys.exit(main())
