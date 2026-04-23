"""Extract pages 25-34 from HWPX (OWPML).

Approach:
 - Page boundaries are only *partially* represented in HWPX: `<hp:p pageBreak="1">`
   marks **user-forced** page breaks (chapter/section starts). Natural overflow
   breaks aren't stored (they'd require a layout engine).
 - We therefore dump:
     (a) a TOC of all forced page-break paragraphs with their first-line text,
         so the user can map logical chapters to page numbers
     (b) if the user-forced breaks align with page 25, the text/struct dump
 - Additionally we try numeric page tracking via a best-effort counter:
     count = 1 + (# forced breaks seen so far)
   This is usually LOWER than the real page number, but useful for anchoring.
"""
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

SRC = Path(r"D:/DATA/카카오톡 받은 파일/제4장 재해영향 예측 및 평가_최종.hwpx")
OUT_DIR = Path(__file__).parent
OUT_TOC = OUT_DIR / "hwpx_toc.txt"
OUT_FULL_TXT = OUT_DIR / "hwpx_full_text.txt"
OUT_STRUCT = OUT_DIR / "hwpx_struct_blocks.txt"

NS = {
    "hp": "http://www.hancom.co.kr/hwpml/2011/paragraph",
    "hs": "http://www.hancom.co.kr/hwpml/2011/section",
    "hh": "http://www.hancom.co.kr/hwpml/2011/head",
    "hc": "http://www.hancom.co.kr/hwpml/2011/core",
}

def local(tag: str) -> str:
    return tag.split("}", 1)[-1] if "}" in tag else tag

def para_text(p) -> str:
    parts = []
    for t in p.iter():
        tag = local(t.tag)
        if tag == "t" and t.text:
            parts.append(t.text)
        elif tag == "tab":
            parts.append("\t")
        elif tag == "lineBreak":
            parts.append("\n")
    return "".join(parts)

def structured_summary(p) -> str:
    out = []
    def rec(e, depth):
        tag = local(e.tag)
        a = e.attrib
        pad = "  " * depth
        if tag == "tbl":
            trs = e.findall(".//hp:tr", NS)
            rowcnt = len(trs)
            # column count = # cells in first row
            first_cells = trs[0].findall("hp:tc", NS) if trs else []
            out.append(f"{pad}[TABLE rows={rowcnt} cols={len(first_cells)}]")
            for ri, tr in enumerate(trs):
                for ci, tc in enumerate(tr.findall("hp:tc", NS)):
                    ps = tc.findall(".//hp:p", NS)
                    ct = " ".join(para_text(pp).strip() for pp in ps)
                    ct = " ".join(ct.split())[:100]
                    out.append(f"{pad}  R{ri}C{ci}: {ct}")
            return
        if tag == "pic":
            img = e.find(".//hc:img", NS)
            bid = img.get("binaryItemIDRef") if img is not None else "?"
            out.append(f"{pad}[IMAGE binRef={bid}]")
            return
        if tag == "fieldBegin":
            out.append(f"{pad}[FIELD-BEGIN type={a.get('type','')} name={a.get('name','')}]")
            return
        if tag == "fieldEnd":
            out.append(f"{pad}[FIELD-END]")
            return
        if tag == "bookmark":
            out.append(f"{pad}[BOOKMARK name={a.get('name','')}]")
            return
        if tag == "t" and e.text:
            out.append(f"{pad}TEXT: {e.text[:120]}")
            return
        for c in e:
            rec(c, depth + 1)
    rec(p, 0)
    return "\n".join(out)

def iter_paragraphs():
    with zipfile.ZipFile(SRC) as z:
        secs = sorted(n for n in z.namelist()
                      if n.startswith("Contents/section") and n.endswith(".xml"))
        for sec in secs:
            root = ET.fromstring(z.open(sec).read())
            for p in root.findall("hp:p", NS):
                yield sec, p

def main():
    toc = []  # (page_counter, sec, idx_in_sec, first_line)
    page = 1
    p_in_sec_count = 0
    last_sec = None
    all_paras = []  # (page, sec, p)

    for sec, p in iter_paragraphs():
        if sec != last_sec:
            p_in_sec_count = 0
            last_sec = sec
        p_in_sec_count += 1
        if p.get("pageBreak") == "1":
            page += 1
            first_line = para_text(p).strip().replace("\n", " ")[:120]
            toc.append((page, sec, p_in_sec_count, first_line))
        all_paras.append((page, sec, p))

    # TOC
    toc_lines = [f"forced page breaks: {len(toc)}", ""]
    for pg, sec, idx, fl in toc:
        toc_lines.append(f"p.{pg:3d}  {sec}  para#{idx:4d}  {fl}")
    OUT_TOC.write_text("\n".join(toc_lines), encoding="utf-8")

    # Dump full text per detected page counter range 25..34
    full = []
    struct = []
    for target in range(25, 35):
        full.append(f"\n===== PAGE {target} (forced-break counter) =====")
        struct.append(f"\n=========== PAGE {target} ===========")
        for pg, sec, p in all_paras:
            if pg != target:
                continue
            t = para_text(p).strip()
            if t:
                full.append(t)
            s = structured_summary(p)
            if s.strip():
                struct.append(s)
                struct.append("---")

    OUT_FULL_TXT.write_text("\n".join(full), encoding="utf-8")
    OUT_STRUCT.write_text("\n".join(struct), encoding="utf-8")

    print(f"TOC   : {OUT_TOC}  (entries={len(toc)})")
    print(f"TEXT  : {OUT_FULL_TXT}")
    print(f"STRUCT: {OUT_STRUCT}")

if __name__ == "__main__":
    main()
