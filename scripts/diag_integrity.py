# -*- coding: utf-8 -*-
"""두 hwpx 의 무결성(필드 짝·run 구조·zip 엔트리)을 비교해 '손상/변조' 원인을 찾는다."""
from __future__ import annotations

import sys
import zipfile
from pathlib import Path

from lxml import etree

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

PARSER = etree.XMLParser(huge_tree=True, recover=True)


def analyze(path):
    z = zipfile.ZipFile(path)
    info = {"names": z.namelist(), "fields": {}, "runs_empty": 0, "ctrl_outside_run": 0}
    sn = [n for n in sorted(z.namelist()) if "section" in n and n.endswith(".xml")][0]
    root = etree.fromstring(z.read(sn), PARSER)
    hp = root.nsmap.get("hp")
    H = lambda t: f"{{{hp}}}{t}"
    begins = [fb.get("id") for fb in root.iter(H("fieldBegin"))]
    ends = [fe.get("beginIDRef") for fe in root.iter(H("fieldEnd"))]
    info["fields"] = {
        "fieldBegin": len(begins), "fieldEnd": len(ends),
        "begin_ids": begins, "end_refs": ends,
        "dup_begin": [x for x in set(begins) if begins.count(x) > 1],
        "dup_end": [x for x in set(ends) if ends.count(x) > 1],
        "orphan_end": [r for r in ends if r not in set(begins)],
        "begin_no_end": [b for b in begins if b not in set(ends)],
    }
    # 빈 run / run 밖 ctrl
    for run in root.iter(H("run")):
        if len(run) == 0 and not (run.text or "").strip():
            info["runs_empty"] += 1
    for ctrl in root.iter(H("ctrl")):
        if ctrl.getparent() is None or ctrl.getparent().tag != H("run"):
            info["ctrl_outside_run"] += 1
    # mimetype 압축 여부
    for zi in z.infolist():
        if zi.filename == "mimetype":
            info["mimetype_compress"] = zi.compress_type  # 0=STORED 정상
    return info


def main():
    a, b = Path(sys.argv[1]), Path(sys.argv[2])
    ia, ib = analyze(a), analyze(b)
    print(f"A(정상): {a.name}")
    print(f"B(손상): {b.name}\n")
    print("=== zip 엔트리 차이 ===")
    print("  A에만:", set(ia["names"]) - set(ib["names"]))
    print("  B에만:", set(ib["names"]) - set(ia["names"]))
    print(f"  mimetype 압축: A={ia.get('mimetype_compress')} B={ib.get('mimetype_compress')} (0=STORED 정상)")
    print("\n=== 필드 무결성 ===")
    for key in ("A", "B"):
        f = (ia if key == "A" else ib)["fields"]
        print(f"  [{key}] begin={f['fieldBegin']} end={f['fieldEnd']} "
              f"중복begin={f['dup_begin']} 중복end={f['dup_end']} "
              f"고아end={f['orphan_end']} 짝없는begin={f['begin_no_end']}")
    print(f"\n  빈 run: A={ia['runs_empty']} B={ib['runs_empty']}")
    print(f"  run 밖 ctrl: A={ia['ctrl_outside_run']} B={ib['ctrl_outside_run']}")
    # begin id 집합 차이
    sa, sb = set(ia["fields"]["begin_ids"]), set(ib["fields"]["begin_ids"])
    print(f"\n  begin id 차이  A-B={sorted(sa-sb)[:10]}  B-A={sorted(sb-sa)[:10]}")


if __name__ == "__main__":
    main()
