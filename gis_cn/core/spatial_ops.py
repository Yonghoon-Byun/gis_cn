import processing
from qgis.core import (
    QgsVectorLayer, QgsFeature, QgsField, QgsFields,
    QgsGeometry
)
from qgis.PyQt.QtCore import QVariant
import logging

logger = logging.getLogger(__name__)

# 분류 레벨별 (코드 컬럼, 이름 컬럼) 매핑
LEVEL_COLUMNS = {
    'l1': ('l1_code', 'l1_name'),
    'l2': ('l2_code', 'l2_name'),
    'l3': ('l3_code', 'l3_name'),
}


def clip_layer(source_layer: QgsVectorLayer,
               mask_layer: QgsVectorLayer,
               output_name: str) -> QgsVectorLayer:
    """source_layer를 mask_layer 경계로 Clip. 모든 원본 컬럼 보존."""
    result = processing.run("native:clip", {
        'INPUT': source_layer,
        'OVERLAY': mask_layer,
        'OUTPUT': 'memory:'
    })
    layer = result['OUTPUT']
    layer.setName(output_name)
    logger.info(f"Clip 완료: {output_name} ({layer.featureCount()} features)")
    return layer


def dissolve_land_cover(layer: QgsVectorLayer, level: str) -> QgsVectorLayer:
    """
    대분류(l1) 또는 중분류(l2) 선택 시, Clip된 토지피복도에서
    동일 분류값을 가진 인접 폴리곤을 병합(Dissolve)한다.
    세분류(l3)는 호출하지 않음.

    dissolve 키: (l1_code, l1_name) 또는 (l2_code, l2_name)
    결과 레이어명은 입력 레이어명 그대로 유지.
    """
    code_col, name_col = LEVEL_COLUMNS[level]
    before = layer.featureCount()
    result = processing.run("native:dissolve", {
        'INPUT': layer,
        'FIELD': [code_col, name_col],
        'OUTPUT': 'memory:'
    })
    dissolved = result['OUTPUT']
    dissolved.setName(layer.name())
    logger.info(
        f"Dissolve 완료: {level}분류 기준 ({code_col}+{name_col}), "
        f"{before} → {dissolved.featureCount()} features"
    )
    return dissolved


def intersect_layers(layer_a: QgsVectorLayer,
                     layer_b: QgsVectorLayer,
                     output_name: str) -> QgsVectorLayer:
    """두 레이어 Intersection. 양쪽 속성 모두 보존."""
    result = processing.run("native:intersection", {
        'INPUT': layer_a,
        'OVERLAY': layer_b,
        'INPUT_FIELDS': [],
        'OVERLAY_FIELDS': [],
        'OVERLAY_FIELDS_PREFIX': '',
        'OUTPUT': 'memory:'
    })
    layer = result['OUTPUT']
    layer.setName(output_name)
    logger.info(f"Intersection 완료: {output_name} ({layer.featureCount()} features)")
    return layer


def build_cn_input_layer(input_polygon_layer: QgsVectorLayer,
                         soil_layer: QgsVectorLayer,
                         land_cover_layer: QgsVectorLayer,
                         name_field: str,
                         level: str = 'l1'):
    """
    3개 레이어 교차 후 CN값_input 레이어 반환.

    처리 순서:
      1. Clip(soil, 소유역)
      2. Clip(land_cover, 소유역)
      3. Intersection(soil_clipped × lc_clipped)
      4. Intersection(결과 × 소유역계)  ← 소유역명 속성 부여
      5. 필요 컬럼만 추출하여 CN값_input 구성

    반환: (cn_input_layer, soil_clipped, lc_clipped)
    """
    _, name_col = LEVEL_COLUMNS[level]

    logger.info("Step 1: 토양군 Clip")
    soil_clipped = clip_layer(soil_layer, input_polygon_layer, "soil_clipped")

    logger.info("Step 2: 토지피복도 Clip")
    lc_clipped = clip_layer(land_cover_layer, input_polygon_layer, "lc_clipped")

    logger.info("Step 3: soil × land_cover Intersection")
    soil_lc = intersect_layers(soil_clipped, lc_clipped, "soil_lc_intersect")

    logger.info("Step 4: (soil×lc) × 소유역계 Intersection")
    final_intersect = intersect_layers(soil_lc, input_polygon_layer, "final_intersect")

    logger.info("Step 5: CN값_input 레이어 구성")
    cn_input = _build_result_layer(final_intersect, name_field, name_col)

    return cn_input, soil_clipped, lc_clipped


def _build_result_layer(source_layer: QgsVectorLayer,
                        name_field: str,
                        land_use_col: str) -> QgsVectorLayer:
    """
    교차 결과에서 지정 컬럼만 추출하여 CN값_input 레이어 생성.
    컬럼: 소유역명, 토지이용, 유역면적(㎡), 토양군, cn값
    """
    crs = source_layer.crs()
    uri = f"Polygon?crs={crs.authid()}"
    result_layer = QgsVectorLayer(uri, "CN값_input", "memory")
    provider = result_layer.dataProvider()

    fields = QgsFields()
    fields.append(QgsField("소유역명", QVariant.String))
    fields.append(QgsField("토지이용", QVariant.String))
    fields.append(QgsField("유역면적", QVariant.Double))
    fields.append(QgsField("토양군", QVariant.String))
    fields.append(QgsField("cn값", QVariant.Int))

    provider.addAttributes(fields)
    result_layer.updateFields()

    src_field_names = [f.name() for f in source_layer.fields()]

    features = []
    for src_feat in source_layer.getFeatures():
        feat = QgsFeature(result_layer.fields())
        geom = src_feat.geometry()
        feat.setGeometry(geom)

        watershed_name = _get_field_value(src_feat, src_field_names, name_field)
        land_use = _get_field_value(src_feat, src_field_names, land_use_col)
        area = geom.area() if geom else 0.0
        hydro_type = _get_field_value(src_feat, src_field_names, "hydro_type")

        feat.setAttributes([watershed_name, land_use, area, hydro_type, None])
        features.append(feat)

    provider.addFeatures(features)
    result_layer.updateExtents()
    logger.info(f"CN값_input 생성: {result_layer.featureCount()} features")
    return result_layer


def _get_field_value(feature, field_names: list, target: str):
    """
    필드명으로 값 조회.
    Intersection 후 중복 컬럼에 접두어가 붙는 경우도 탐색.
    예: "hydro_type" → "2_hydro_type" 등
    """
    if target is None:
        return None
    # 정확 매칭
    if target in field_names:
        return feature[target]
    # 접미어 매칭 (예: "_hydro_type" 포함)
    for name in field_names:
        if name == target or name.endswith(f"_{target}"):
            return feature[name]
    return None
