import psycopg2
from qgis.core import QgsVectorLayer, QgsFeature, QgsGeometry, QgsField, QgsFields
from qgis.PyQt.QtCore import QVariant
import logging

logger = logging.getLogger(__name__)

DB_CONFIG = {
    "host": "geo-spatial-hub-prod.postgres.database.azure.com",
    "port": 6432,
    "dbname": "dde-water",
    "user": "waterviewer",
    "password": "water123!@#",
}

GEOM_COLUMN = "geom"

# psycopg2 OID → QVariant 타입 매핑
_INT_OIDS = {20, 21, 23}  # int8, int2, int4
_FLOAT_OIDS = {700, 701, 1700}  # float4, float8, numeric

# level별 (코드 컬럼, 이름 컬럼) — dissolve 기준
_LC_DISSOLVE_COLS = {
    "l1": ("l1_code", "l1_name"),
    "l2": ("l2_code", "l2_name"),
    "l3": ("l3_code", "l3_name"),
}


def get_connection():
    """PostGIS DB 연결 반환"""
    return psycopg2.connect(**DB_CONFIG)


def _rows_to_memory_layer(rows, col_names, col_oids, geom_col, layer_name, srid):
    """
    psycopg2 결과 행을 QgsVectorLayer (메모리) 로 변환.
    geometry 컬럼은 ST_AsText()로 WKT 형태로 전달받아야 함.
    """
    geom_idx = col_names.index(geom_col)

    fields = QgsFields()
    for name, oid in zip(col_names, col_oids):
        if name == geom_col:
            continue
        if oid in _INT_OIDS:
            fields.append(QgsField(name, QVariant.Int))
        elif oid in _FLOAT_OIDS:
            fields.append(QgsField(name, QVariant.Double))
        else:
            fields.append(QgsField(name, QVariant.String))

    uri = f"Polygon?crs=EPSG:{srid}"
    layer = QgsVectorLayer(uri, layer_name, "memory")
    provider = layer.dataProvider()
    provider.addAttributes(fields)
    layer.updateFields()

    features = []
    for row in rows:
        feat = QgsFeature(layer.fields())
        wkt = row[geom_idx]
        if wkt:
            feat.setGeometry(QgsGeometry.fromWkt(wkt))
        attr_idx = 0
        for i, name in enumerate(col_names):
            if name == geom_col:
                continue
            feat.setAttribute(attr_idx, row[i])
            attr_idx += 1
        features.append(feat)

    provider.addFeatures(features)
    layer.updateExtents()
    return layer


def get_soil_layer(polygon_geom_wkt: str, srid: int = 5186) -> QgsVectorLayer:
    """
    소유역 폴리곤으로 Clip된 토양군 레이어 반환 (PostGIS ST_Intersection 적용).
    컬럼: soil_code, hydro_type, hydro_ty_1, k
    """
    sql = """
        WITH w AS (
            SELECT ST_GeomFromText(%(wkt)s, %(srid)s) AS geom
        )
        SELECT soil_code, hydro_type, hydro_ty_1, k,
               ST_AsText(clipped) AS geom
        FROM (
            SELECT soil_code, hydro_type, hydro_ty_1, k,
                   ST_CollectionExtract(
                       ST_Intersection(ST_Transform(s.geom, %(srid)s), w.geom),
                       3
                   ) AS clipped
            FROM public.soil s, w
            WHERE ST_Intersects(s.geom, ST_Transform(w.geom, ST_SRID(s.geom)))
        ) t
        WHERE clipped IS NOT NULL
          AND NOT ST_IsEmpty(clipped)
          AND ST_Area(clipped) > 0
    """
    try:
        conn = get_connection()
        cur = conn.cursor()
        cur.execute(sql, {"wkt": polygon_geom_wkt, "srid": srid})
        rows = cur.fetchall()
        col_names = [desc[0] for desc in cur.description]
        col_oids = [desc[1] for desc in cur.description]
        cur.close()
        conn.close()
        logger.info(f"토양군 clip: {len(rows)}개 피처")
        return _rows_to_memory_layer(
            rows, col_names, col_oids, "geom", "토양군_clip", srid
        )
    except Exception as e:
        logger.error(f"토양군 추출 오류: {e}")
        raise


def get_land_cover_layer(
    polygon_geom_wkt: str, srid: int = 5186, level: str = "l1"
) -> QgsVectorLayer:
    """
    소유역 폴리곤으로 Clip된 토지피복도 레이어 반환 (PostGIS ST_Intersection 적용).
    컬럼: gid, l1_code, l1_name, l2_code, l2_name, l3_code, l3_name
    """
    sql = """
        WITH w AS (
            SELECT ST_GeomFromText(%(wkt)s, %(srid)s) AS geom
        )
        SELECT gid, l1_code, l1_name, l2_code, l2_name, l3_code, l3_name,
               ST_AsText(clipped) AS geom
        FROM (
            SELECT gid, l1_code, l1_name, l2_code, l2_name, l3_code, l3_name,
                   ST_CollectionExtract(
                       ST_Intersection(ST_Transform(l.geom, %(srid)s), w.geom),
                       3
                   ) AS clipped
            FROM public.land_cover_yangju l, w
            WHERE ST_Intersects(l.geom, ST_Transform(w.geom, ST_SRID(l.geom)))
        ) t
        WHERE clipped IS NOT NULL
          AND NOT ST_IsEmpty(clipped)
          AND ST_Area(clipped) > 0
    """
    try:
        conn = get_connection()
        cur = conn.cursor()
        cur.execute(sql, {"wkt": polygon_geom_wkt, "srid": srid})
        rows = cur.fetchall()
        col_names = [desc[0] for desc in cur.description]
        col_oids = [desc[1] for desc in cur.description]
        cur.close()
        conn.close()
        logger.info(f"토지피복도 clip: {len(rows)}개 피처 (level={level})")
        return _rows_to_memory_layer(
            rows, col_names, col_oids, "geom", "토지피복도_clip", srid
        )
    except Exception as e:
        logger.error(f"토지피복도 추출 오류: {e}")
        raise


def get_soil_lc_intersection(
    polygon_geom_wkt: str, srid: int = 5186, level: str = "l1"
) -> QgsVectorLayer:
    """
    PostGIS에서 토양군 × 토지피복도 Intersection 일괄 처리.
    - 소유역으로 clip
    - level 기준 토지피복도 ST_Union(Dissolve)
    - 토양군 × 토지피복도 ST_Intersection
    반환 컬럼: soil_code, hydro_type, hydro_ty_1, k, {code_col}, {name_col}
    """
    code_col, name_col = _LC_DISSOLVE_COLS[level]
    sql = f"""
        WITH w AS (
            SELECT ST_GeomFromText(%(wkt)s, %(srid)s) AS geom
        ),
        sc AS (
            SELECT soil_code, hydro_type, hydro_ty_1, k,
                   ST_CollectionExtract(
                       ST_Intersection(ST_Transform(s.geom, %(srid)s), w.geom),
                       3
                   ) AS geom
            FROM public.soil s, w
            WHERE ST_Intersects(s.geom, ST_Transform(w.geom, ST_SRID(s.geom)))
        ),
        sc_valid AS (
            SELECT soil_code, hydro_type, hydro_ty_1, k, geom
            FROM sc
            WHERE geom IS NOT NULL AND NOT ST_IsEmpty(geom) AND ST_Area(geom) > 0
        ),
        lc_raw AS (
            SELECT {code_col}, {name_col},
                   ST_CollectionExtract(
                       ST_Intersection(ST_Transform(l.geom, %(srid)s), w.geom),
                       3
                   ) AS geom
            FROM public.land_cover_yangju l, w
            WHERE ST_Intersects(l.geom, ST_Transform(w.geom, ST_SRID(l.geom)))
        ),
        lc AS (
            SELECT {code_col}, {name_col},
                   ST_Union(geom) AS geom
            FROM lc_raw
            WHERE geom IS NOT NULL AND NOT ST_IsEmpty(geom) AND ST_Area(geom) > 0
            GROUP BY {code_col}, {name_col}
        ),
        ix AS (
            SELECT sc.soil_code, sc.hydro_type, sc.hydro_ty_1, sc.k,
                   lc.{code_col}, lc.{name_col},
                   ST_CollectionExtract(ST_Intersection(sc.geom, lc.geom), 3) AS geom
            FROM sc_valid sc
            JOIN lc ON ST_Intersects(sc.geom, lc.geom)
        )
        SELECT soil_code, hydro_type, hydro_ty_1, k,
               {code_col}, {name_col},
               ST_AsText(geom) AS geom
        FROM ix
        WHERE geom IS NOT NULL
          AND NOT ST_IsEmpty(geom)
          AND ST_Area(geom) > 0
    """
    try:
        conn = get_connection()
        cur = conn.cursor()
        cur.execute(sql, {"wkt": polygon_geom_wkt, "srid": srid})
        rows = cur.fetchall()
        col_names = [desc[0] for desc in cur.description]
        col_oids = [desc[1] for desc in cur.description]
        cur.close()
        conn.close()
        logger.info(f"토양군×토지피복 교차: {len(rows)}개 피처 (level={level})")
        return _rows_to_memory_layer(
            rows, col_names, col_oids, "geom", "토양군_토지피복_교차", srid
        )
    except Exception as e:
        logger.error(f"토양군×토지피복 교차 오류: {e}")
        raise
