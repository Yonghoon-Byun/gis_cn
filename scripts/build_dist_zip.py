"""dist/gis_cn.zip 재빌드 — 플러그인 소스 변경 후 항상 실행.

소스 루트의 gis_cn/ 전체를 zip으로 묶어 dist/gis_cn.zip 생성.
__pycache__, *.pyc, .DS_Store 등 부산물은 제외.
"""
import os
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "gis_cn"
# 메인 프로젝트(worktree 아님)의 dist 폴더에 항상 zip 저장.
# worktree의 dist에도 관례상 복사.
MAIN_OUT = Path(
    r"D:\DATA\연구원 연도별 업무사항\06_2026년 개발 자료\04_인프라부문 AI 개발\gis_cn\dist\gis_cn.zip"
)
WT_OUT = ROOT / "dist" / "gis_cn.zip"

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
    MAIN_OUT.parent.mkdir(parents=True, exist_ok=True)
    WT_OUT.parent.mkdir(parents=True, exist_ok=True)
    count = 0
    with zipfile.ZipFile(MAIN_OUT, "w", zipfile.ZIP_DEFLATED) as zf:
        for path in sorted(SRC.rglob("*")):
            if not path.is_file() or should_skip(path):
                continue
            arc = path.relative_to(SRC.parent)  # e.g. gis_cn/core/...
            zf.write(path, arc.as_posix())
            count += 1
    size = MAIN_OUT.stat().st_size
    print(f"wrote: {MAIN_OUT}  ({count} files, {size:,} bytes)")

    # worktree 의 dist 에도 동일하게 복사 (관례 유지)
    import shutil
    shutil.copy2(MAIN_OUT, WT_OUT)
    print(f"mirrored: {WT_OUT}")


if __name__ == "__main__":
    main()
