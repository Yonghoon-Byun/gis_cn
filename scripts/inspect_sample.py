"""추출된 cn_report_sample.hwpx의 구조 검증."""
import zipfile
from pathlib import Path

SAMPLE = Path(__file__).resolve().parent.parent / "gis_cn" / "templates" / "v1.0" / "cn_report_sample.hwpx"

z = zipfile.ZipFile(SAMPLE)
print(f"Sample file: {SAMPLE}")
print(f"Total size: {SAMPLE.stat().st_size:,} bytes")
print()

print("Entries:")
for n in sorted(z.namelist()):
    print(f"  {z.getinfo(n).file_size:>10,}  {n}")
print()

sec_names = [n for n in z.namelist() if n.startswith("Contents/section") and n.endswith(".xml")]
print(f"sections: {sec_names}")
data = "".join(z.open(n).read().decode("utf-8", errors="replace") for n in sec_names)
print(f"total section xml length: {len(data):,} chars")

keywords = [
    "유출곡선지수", "CN", "홍수량", "토지이용",
    "매개변수", "소유역", "산정결과", "TYPE A",
    "Clark", "임계지속기간", "토양통", "토사",
]
print()
print("Keyword occurrences in section1.xml:")
for kw in keywords:
    print(f"  {kw:15s}: {data.count(kw)}")

# 표 개수
import re
n_tables = len(re.findall(r"<hp:tbl\s", data))
n_paragraphs = len(re.findall(r"<hp:p\s", data))
n_pics = len(re.findall(r"<hp:pic\s", data))
print()
print(f"Tables: {n_tables}")
print(f"Paragraphs: {n_paragraphs}")
print(f"Pictures: {n_pics}")
