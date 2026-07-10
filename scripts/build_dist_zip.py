"""dist/gis_cn.zip 재빌드 — 플러그인 소스 변경 후 항상 실행.

소스 루트의 gis_cn/ 전체를 zip으로 묶어 dist/gis_cn.zip 생성.
__pycache__, *.pyc, .DS_Store 등 부산물은 제외.
"""
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "gis_cn"
OUT = ROOT / "dist" / "gis_cn.zip"

EXCLUDE_DIRS = {"__pycache__", ".pytest_cache", ".mypy_cache"}
EXCLUDE_EXTS = {".pyc", ".pyo"}
EXCLUDE_NAMES = {".DS_Store", "Thumbs.db"}


def should_skip(p: Path) -> bool:
    if p.name in EXCLUDE_NAMES:
        return True
    if p.suffix in EXCLUDE_EXTS:
        return True
    parts = set(p.parts)
    return bool(parts & EXCLUDE_DIRS)


def main():
    OUT.parent.mkdir(parents=True, exist_ok=True)
    count = 0
    with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as zf:
        for path in sorted(SRC.rglob("*")):
            if not path.is_file() or should_skip(path):
                continue
            arc = path.relative_to(SRC.parent)  # e.g. gis_cn/core/...
            zf.write(path, arc.as_posix())
            count += 1
    size = OUT.stat().st_size
    print(f"wrote: {OUT}  ({count} files, {size:,} bytes)")


if __name__ == "__main__":
    main()
