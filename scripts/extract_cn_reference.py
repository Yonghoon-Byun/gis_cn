# -*- coding: utf-8 -*-
"""P0: CN 기준표(국가표준) 추출 자산 생성.

소스: gis_cn/templates/v1.0/cn_report_sample.hwpx 의 CN 기준표
      ("우리나라 토지이용 형태에 따른 유출곡선지수", 43행x8열, 직접 실측).
출력: gis_cn/data/cn_reference_std.json — 41 세분류행 [{l1,l2,code,lu,l3_code,l3_name,a,b,c,d}]

교차검증: docs/토지피복도_분류체계.md (세분류 41항목, 고유 3자리 코드)와
- 행수 41 == docs 세분류 41 (분류순서 동일 → 위치 기반 l3_code 부여)
- CN표 '코드'(중분류 110/120/...)가 docs 중분류 코드집합의 부분집합

주의: CN 기준표의 '코드' 열은 **중분류 코드**(110 등)이고, land_cover.l3_code는
**세분류 코드**(111 등)다. 둘은 다르므로 위치매핑으로 l3_code를 별도 부여한다.
한컴오피스/COM 불필요 (lxml only). 결정적(deterministic) 재생성.
"""
from __future__ import annotations
import json, re, sys, zipfile
from pathlib import Path
from lxml import etree

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

ROOT = Path(__file__).resolve().parent.parent
SAMPLE = ROOT / "gis_cn" / "templates" / "v1.0" / "cn_report_sample.hwpx"
DOCS_MD = ROOT / "docs" / "토지피복도_분류체계.md"
OUT = ROOT / "gis_cn" / "data" / "cn_reference_std.json"


def _cell_text(el, ns: str) -> str:
    return re.sub(r"\s+", "", "".join(t.text or "" for t in el.iter(f"{{{ns}}}t")))


def extract_ref_table(sample: Path) -> list[dict]:
    """샘플 hwpx의 8열 CN 기준표를 병합셀 인지 논리그리드로 복원 → 데이터행 추출."""
    root = etree.fromstring(zipfile.ZipFile(sample).read("Contents/section0.xml"),
                            etree.XMLParser(huge_tree=True, recover=True))
    ns = root.nsmap.get("hp")
    hp = lambda t: f"{{{ns}}}{t}"
    tbl = next(t for t in root.iter(hp("tbl"))
               if "AMC-II" in "".join(x.text or "" for x in (t.find(hp("caption")) if t.find(hp("caption")) is not None else t).iter(hp("t")))
               and t.get("colCnt") == "8")
    rows, cols = int(tbl.get("rowCnt")), int(tbl.get("colCnt"))
    grid = [["" for _ in range(cols)] for _ in range(rows)]
    for tr in tbl.findall(hp("tr")):
        for tc in tr.findall(hp("tc")):
            ca, sp = tc.find(hp("cellAddr")), tc.find(hp("cellSpan"))
            r, c = int(ca.get("rowAddr")), int(ca.get("colAddr"))
            rs = int(sp.get("rowSpan")) if sp is not None else 1
            cs = int(sp.get("colSpan")) if sp is not None else 1
            txt = _cell_text(tc, ns)
            for dr in range(rs):
                for dc in range(cs):
                    if r + dr < rows and c + dc < cols:
                        grid[r + dr][c + dc] = txt
    out = []
    for r in range(rows):
        a = grid[r][4]
        if re.fullmatch(r"\d+", a or ""):  # CN 값(숫자)이 있는 데이터행만
            out.append({"l1": grid[r][0], "l2": grid[r][1], "code": grid[r][2], "lu": grid[r][3],
                        "a": int(grid[r][4]), "b": int(grid[r][5]), "c": int(grid[r][6]), "d": int(grid[r][7])})
    return out


def parse_docs_subclasses(md: Path) -> list[tuple[str, str]]:
    """docs md의 '### 세분류' 색상표에서 (코드, 이름) 41개를 순서대로 추출."""
    text = md.read_text(encoding="utf-8")
    seg = text.split("### 세분류", 1)[1]
    pairs = []
    for line in seg.splitlines():
        m = re.match(r"\|\s*([^|]+?)\s*\|\s*(\d{3})\s*\|", line)
        if m:
            pairs.append((m.group(2), m.group(1)))
    return pairs


def parse_docs_midclass_codes(md: Path) -> set[str]:
    """docs 중분류 코드집합(110 등)."""
    text = md.read_text(encoding="utf-8")
    seg = text.split("### 중분류", 1)[1].split("### 세분류", 1)[0]
    # 행 시작 기준 2번째 열(코드)만 추출 (RGB 3자리 오매칭 방지)
    return set(re.findall(r"^\|\s*[^|]+?\s*\|\s*(\d{3})\s*\|", seg, re.MULTILINE))


def main() -> int:
    ref = extract_ref_table(SAMPLE)
    subs = parse_docs_subclasses(DOCS_MD)
    mid_codes = parse_docs_midclass_codes(DOCS_MD)

    print(f"CN 기준표 데이터행: {len(ref)}")
    print(f"docs 세분류: {len(subs)} / docs 중분류 코드: {len(mid_codes)}")

    problems, warns = [], []
    if len(ref) != len(subs):
        problems.append(f"행수 불일치: CN표 {len(ref)} vs docs 세분류 {len(subs)} (위치매핑 불가)")
    else:
        for i, row in enumerate(ref):
            l3_code, l3_name = subs[i]
            row["l3_code"], row["l3_name"] = l3_code, l3_name
            if row["code"] not in mid_codes:
                problems.append(f"[{i}] 중분류코드 {row['code']} 가 docs 중분류집합에 없음")
            # 세분류 코드의 앞 2자리는 중분류 코드의 앞 2자리와 같아야(같은 계열)
            if l3_code[:2] != row["code"][:2]:
                problems.append(f"[{i}] 계열 불일치: code={row['code']} vs l3_code={l3_code} ({row['lu']}/{l3_name})")
            if row["lu"] != l3_name:
                warns.append(f"[{i}] 이름변형: CN표'{row['lu']}' ↔ docs'{l3_name}' (코드 {l3_code})")

    print(f"\n중분류코드 ⊆ docs: {'OK' if not any('중분류코드' in p for p in problems) else 'FAIL'}")
    print(f"계열(앞2자리) 정합: {'OK' if not any('계열' in p for p in problems) else 'FAIL'}")
    print(f"이름변형(정상, 참고): {len(warns)}건")
    for w in warns[:8]:
        print("   ·", w)

    if problems:
        print(f"\n❌ 차단 문제 {len(problems)}건:")
        for p in problems:
            print("   -", p)
        return 1

    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(json.dumps(ref, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"\n✅ 저장: {OUT}  ({len(ref)}행, 8열 + l3_code/l3_name 매핑)")
    print("샘플:", json.dumps(ref[0], ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
