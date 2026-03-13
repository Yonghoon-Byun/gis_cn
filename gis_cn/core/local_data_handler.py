import os
import logging
from qgis.core import QgsVectorLayer, QgsFeature, QgsField, QgsFields, QgsGeometry
from qgis.PyQt.QtCore import QVariant

logger = logging.getLogger(__name__)


class ValidationError(Exception):
    """로컬 파일 검증 오류."""
    pass


# ── 컬럼 자동 감지용 별칭 매핑 ──────────────────────────────────────────────

SOIL_COLUMN_ALIASES = {
    'hydro_type': ['hydro_type', 'HYDRO_TYPE', 'HYDGRP', 'hydgrp',
                   '수문학토양군', '수문토양군', '토양군', 'HSG', 'hsg'],
}

LC_COLUMN_ALIASES = {
    'l1_code': ['l1_code', 'L1_CODE', '대분류코드', '대분류_코드'],
    'l1_name': ['l1_name', 'L1_NAME', '대분류명', '대분류_명', '대분류'],
    'l2_code': ['l2_code', 'L2_CODE', '중분류코드', '중분류_코드'],
    'l2_name': ['l2_name', 'L2_NAME', '중분류명', '중분류_명', '중분류'],
    'l3_code': ['l3_code', 'L3_CODE', '세분류코드', '세분류_코드'],
    'l3_name': ['l3_name', 'L3_NAME', '세분류명', '세분류_명', '세분류'],
}


def _resolve_columns(layer, alias_map):
    """
    레이어 필드에서 canonical 컬럼명을 자동 감지.
    반환: {canonical_name: actual_field_name}
    감지 실패 시 ValidationError 발생.
    """
    field_names = [f.name() for f in layer.fields()]
    resolved = {}
    missing = []
    for canonical, aliases in alias_map.items():
        found = None
        for alias in aliases:
            if alias in field_names:
                found = alias
                break
        if found:
            resolved[canonical] = found
        else:
            missing.append(canonical)
    if missing:
        supported = []
        for m in missing:
            supported.extend(alias_map[m])
        raise ValidationError(
            f"필수 컬럼을 찾을 수 없습니다.\n"
            f"기대 컬럼: {missing}\n"
            f"실제 컬럼: {field_names}\n"
            f"지원되는 이름: {supported}"
        )
    return resolved


def _rename_columns(layer, column_map):
    """
    actual → canonical 리네이밍. 메모리 레이어로 복사하며 컬럼명 변경.
    column_map: {canonical_name: actual_field_name}
    모든 canonical == actual이면 원본 레이어를 그대로 반환 (복사 없음).
    """
    # Check if renaming is needed
    needs_rename = any(canonical != actual for canonical, actual in column_map.items())
    if not needs_rename:
        return layer

    # Build reverse map: actual_name -> canonical_name
    reverse_map = {actual: canonical for canonical, actual in column_map.items()}

    # Create new fields with canonical names
    src_fields = layer.fields()
    new_fields = QgsFields()
    for i in range(src_fields.count()):
        field = src_fields.at(i)
        name = field.name()
        if name in reverse_map:
            new_field = QgsField(reverse_map[name], field.type(), field.typeName(),
                                 field.length(), field.precision())
            new_fields.append(new_field)
        else:
            new_fields.append(QgsField(field))

    # Create memory layer
    geom_type = layer.geometryType()
    geom_str = {0: 'Point', 1: 'LineString', 2: 'Polygon'}.get(geom_type, 'Polygon')
    crs = layer.crs().authid()
    mem_layer = QgsVectorLayer(f"{geom_str}?crs={crs}", layer.name(), "memory")
    provider = mem_layer.dataProvider()
    provider.addAttributes(new_fields)
    mem_layer.updateFields()

    # Copy features with remapped attributes
    features = []
    for src_feat in layer.getFeatures():
        feat = QgsFeature(mem_layer.fields())
        feat.setGeometry(QgsGeometry(src_feat.geometry()))
        # Copy attributes by index
        attrs = []
        for i in range(src_fields.count()):
            attrs.append(src_feat.attributes()[i])
        feat.setAttributes(attrs)
        features.append(feat)

    provider.addFeatures(features)
    mem_layer.updateExtents()
    logger.info(f"컬럼 리네이밍 완료: {dict((v, k) for k, v in column_map.items() if k != v)}")
    return mem_layer


def _get_required_lc_aliases(level):
    """분류 수준에 따라 필요한 토지피복도 컬럼 alias만 반환."""
    from .spatial_ops import LEVEL_COLUMNS
    code_col, name_col = LEVEL_COLUMNS[level]
    required = {}
    if code_col in LC_COLUMN_ALIASES:
        required[code_col] = LC_COLUMN_ALIASES[code_col]
    if name_col in LC_COLUMN_ALIASES:
        required[name_col] = LC_COLUMN_ALIASES[name_col]
    return required


def load_local_soil(file_path, clip_layer):
    """
    로컬 토양군 파일 로드 → 컬럼 자동 감지 → canonical 리네이밍 → Clip.

    Args:
        file_path: 토양군 SHP/GPKG 파일 경로
        clip_layer: 클리핑 마스크 레이어 (소유역 전체 영역 dissolve)
    Returns:
        QgsVectorLayer: canonical 컬럼명을 가진 클립된 토양군 레이어
    Raises:
        ValidationError: 필수 컬럼 감지 실패
    """
    from .spatial_ops import clip_layer as do_clip

    # Load file
    layer = QgsVectorLayer(file_path, "토양군_원본", "ogr")
    if not layer.isValid():
        raise ValidationError(f"파일을 열 수 없습니다: {file_path}")

    # Auto-detect and rename columns
    column_map = _resolve_columns(layer, SOIL_COLUMN_ALIASES)
    renamed = _rename_columns(layer, column_map)

    # Clip to study area
    clipped = do_clip(renamed, clip_layer, "토양군_clip")
    logger.info(f"로컬 토양군 로드 완료: {clipped.featureCount()} features")
    return clipped


def load_local_land_cover(file_path, clip_layer, level):
    """
    로컬 토지피복도 파일 로드 → 컬럼 자동 감지 → canonical 리네이밍 → Clip → Dissolve.

    Args:
        file_path: 토지피복도 SHP/GPKG 파일 경로
        clip_layer: 클리핑 마스크 레이어 (소유역 전체 영역 dissolve)
        level: 분류 수준 ('l1', 'l2', 'l3')
    Returns:
        QgsVectorLayer: canonical 컬럼명을 가진 클립+디졸브된 토지피복도 레이어
    Raises:
        ValidationError: 필수 컬럼 감지 실패
    """
    from .spatial_ops import clip_layer as do_clip, dissolve_land_cover

    # Load file
    layer = QgsVectorLayer(file_path, "토지피복도_원본", "ogr")
    if not layer.isValid():
        raise ValidationError(f"파일을 열 수 없습니다: {file_path}")

    # Auto-detect required columns for this level
    required_aliases = _get_required_lc_aliases(level)
    column_map = _resolve_columns(layer, required_aliases)
    renamed = _rename_columns(layer, column_map)

    # Clip to study area
    clipped = do_clip(renamed, clip_layer, "토지피복도_clip")

    # Dissolve (l1, l2 only — l3 skips dissolve)
    if level in ('l1', 'l2'):
        result = dissolve_land_cover(clipped, level)
    else:
        result = clipped

    logger.info(f"로컬 토지피복도 로드 완료: {result.featureCount()} features (level={level})")
    return result


def get_local_soil_lc_intersection(soil_layer, lc_layer, level):
    """
    이미 canonical 리네이밍된 토양군 × 토지피복도 Intersection.

    Args:
        soil_layer: load_local_soil()의 반환값
        lc_layer: load_local_land_cover()의 반환값
        level: 분류 수준 ('l1', 'l2', 'l3')
    Returns:
        QgsVectorLayer: 토양군_토지피복_교차 레이어
    """
    from .spatial_ops import intersect_layers

    result = intersect_layers(soil_layer, lc_layer, "토양군_토지피복_교차")
    logger.info(f"로컬 Intersection 완료: {result.featureCount()} features")
    return result
