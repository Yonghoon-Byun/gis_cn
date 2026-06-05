# -*- coding: utf-8 -*-
"""두 hwpx 의 단락별 run 시그니처를 나란히 비교 — 구조적 차이(추가/삭제/변형) 지점 탐지.

각 top-level <hp:p> 의 (paraPrIDRef, [run별 charPrIDRef:자식태그…], linesegarray 유무)를 뽑아
A/B 를 정렬 비교한다. run 삭제·삽입·자식변형을 한눈에.
"""
from __future__ import annotations

import sys
import zipfile
from pathlib import Path

from lxml import etree

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

PARSER = etree.XMLParser(huge_tree=True, recover=True)


def sigs(path):
    z = zipfile.ZipFile(path)
    sn = [n for n in sorted(z.namelist()) if "section" in n and n.endswith(".xml")][0]
    root = etree.fromstring(z.read(sn), PARSER)
    hp = root.nsmap.get("hp")
    H = lambda t: f"{{{hp}}}{t}"
    out = []
    for p in root:
        if p.tag != H("p"):
            continue
        runs = []
        has_lsa = False
        for ch in p:
            ln = etree.QName(ch).localname
            if ln == "run":
                kids = "+".join(etree.QName(c).localname for c in ch)
                runs.append(f"{ch.get('charPrIDRef')}:{kids or '∅'}")
            elif ln == "linesegarray":
                has_lsa = True
        out.append((p.get("paraPrIDRef"), runs, has_lsa))
    return out


def main():
    a, b = Path(sys.argv[1]), Path(sys.argv[2])
    sa, sb = sigs(a), sigs(b)
    print(f"A: {a.name}  (단락 {len(sa)})")
    print(f"B: {b.name}  (단락 {len(sb)})\n")
    n = max(len(sa), len(sb))
    for i in range(n):
        ra = sa[i] if i < len(sa) else None
        rb = sb[i] if i < len(sb) else None
        mark = "  " if ra == rb else "≠≠"
        print(f"{mark} #{i}")
        if ra != rb:
            print(f"    A: para={ra[0] if ra else '-'} lsa={ra[2] if ra else '-'} runs={ra[1] if ra else '-'}")
            print(f"    B: para={rb[0] if rb else '-'} lsa={rb[2] if rb else '-'} runs={rb[1] if rb else '-'}")


if __name__ == "__main__":
    main()
