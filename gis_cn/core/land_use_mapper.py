import json
import os
import logging

logger = logging.getLogger(__name__)

# 플러그인 루트에 매핑 파일 저장
MAPPING_PATH = os.path.join(
    os.path.dirname(os.path.dirname(__file__)), 'land_use_mapping.json'
)


def load_mapping(path: str = None) -> dict:
    """저장된 매핑 로드. 반환: {원본명: 재분류명}"""
    p = path or MAPPING_PATH
    if not os.path.exists(p):
        return {}
    try:
        with open(p, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        logger.error(f"매핑 로드 오류: {e}")
        return {}


def save_mapping(mapping: dict, path: str = None):
    """매핑을 JSON으로 저장."""
    p = path or MAPPING_PATH
    with open(p, 'w', encoding='utf-8') as f:
        json.dump(mapping, f, ensure_ascii=False, indent=2)
    logger.info(f"매핑 저장: {len(mapping)}개 항목 ({p})")


def apply_mapping_to_layer(layer, mapping: dict) -> int:
    """
    CN값_input 레이어의 '토지이용' 필드에 매핑 적용.
    provider.changeAttributeValues() 사용 — 편집 모드 없이 스레드 안전.
    반환: 변경된 피처 수
    """
    if not mapping:
        return 0
    field_idx = layer.fields().indexFromName('토지이용')
    if field_idx < 0:
        return 0

    provider = layer.dataProvider()
    changes = {}
    for feat in layer.getFeatures():
        original = feat['토지이용']
        if original and str(original) in mapping:
            changes[feat.id()] = {field_idx: mapping[str(original)]}

    if changes:
        provider.changeAttributeValues(changes)
    logger.info(f"토지이용 재분류 적용: {len(changes)}건")
    return len(changes)
