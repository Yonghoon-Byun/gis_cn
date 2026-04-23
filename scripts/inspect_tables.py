"""샘플 HWPX의 표 구조를 모두 덤프 — 각 표의 행/열, 첫 행(헤더) 내용 표시."""
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

SAMPLE = Path(__file__).resolve().parent.parent / "gis_cn" / "templates" / "v1.0" / "cn_report_sample.hwpx"

NS = {
    "hp": "http://www.hancom.co.kr/hwpml/2011/paragraph",
    "hs": "http://www.hancom.co.kr/hwpml/2011/section",
    "hh": "http://www.hancom.co.kr/hwpml/2011/head",
    "hc": "http://www.hancom.co.kr/hwpml/2011/core",
}


def local(tag):
    return tag.split("}", 1)[-1] if "}" in tag else tag


def para_text(p):
    parts = []
    for t in p.iter():
        if local(t.tag) == "t" and t.text:
            parts.append(t.text)
    return " ".join(parts).strip()


def main():
    z = zipfile.ZipFile(SAMPLE)
    sections = sorted(n for n in z.namelist()
                      if n.startswith("Contents/section") and n.endswith(".xml"))
    all_tables = []
    for sec in sections:
        root = ET.fromstring(z.open(sec).read())
        for tbl in root.findall(".//hp:tbl", NS):
            all_tables.append(tbl)

    print(f"total tables: {len(all_tables)}\n")
    for i, tbl in enumerate(all_tables):
        trs = tbl.findall("hp:tr", NS)
        first_cells = trs[0].findall("hp:tc", NS) if trs else []
        cols = len(first_cells)
        rows = len(trs)
        print(f"=== table[{i}]  rows={rows}  cols={cols} ===")
        # 첫 2행의 셀 텍스트 출력
        for ri, tr in enumerate(trs[:3]):
            for ci, tc in enumerate(tr.findall("hp:tc", NS)):
                ps = tc.findall(".//hp:p", NS)
                txt = " | ".join(para_text(p) for p in ps if para_text(p))
                txt = txt[:50]
                print(f"  R{ri}C{ci}: {txt}")
            print()
        if rows > 3:
            # 마지막 2행도
            for ri_offset, tr in enumerate(trs[-2:]):
                ri = rows - 2 + ri_offset
                for ci, tc in enumerate(tr.findall("hp:tc", NS)):
                    ps = tc.findall(".//hp:p", NS)
                    txt = " | ".join(para_text(p) for p in ps if para_text(p))
                    txt = txt[:50]
                    print(f"  R{ri}C{ci}: {txt}")
                print()
        print()


if __name__ == "__main__":
    main()
