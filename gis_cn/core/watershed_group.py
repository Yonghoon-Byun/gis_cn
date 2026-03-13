import json
import os
import logging

logger = logging.getLogger(__name__)

# 플러그인 루트에 유역합성 그룹 파일 저장
GROUPS_PATH = os.path.join(
    os.path.dirname(os.path.dirname(__file__)), 'watershed_groups.json'
)


def load_groups(path: str = None) -> dict:
    """
    저장된 유역합성 그룹 로드.
    반환: {그룹명: [소유역1, 소유역2, ...]}
    """
    p = path or GROUPS_PATH
    if not os.path.exists(p):
        return {}
    try:
        with open(p, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        logger.error(f"유역합성 그룹 로드 오류: {e}")
        return {}


def save_groups(groups: dict, path: str = None):
    """유역합성 그룹을 JSON으로 저장."""
    p = path or GROUPS_PATH
    with open(p, 'w', encoding='utf-8') as f:
        json.dump(groups, f, ensure_ascii=False, indent=2)
    logger.info(f"유역합성 그룹 저장: {len(groups)}개 그룹 ({p})")
