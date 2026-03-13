import os
import pandas as pd
import logging

logger = logging.getLogger(__name__)

# cn_value.xlsx 기본 경로 (플러그인 루트)
XLSX_PATH = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'cn_value.xlsx')

HYDRO_COLUMNS = {'A', 'B', 'C', 'D'}


def load_cn_table(xlsx_path: str = None) -> pd.DataFrame:
    """
    cn_value.xlsx 로드.
    컬럼: 토지이용분류, A, B, C, D
    """
    path = xlsx_path or XLSX_PATH
    if not os.path.exists(path):
        raise FileNotFoundError(f"CN값 파일을 찾을 수 없습니다: {path}")
    df = pd.read_excel(path, dtype={'토지이용분류': str})
    df['토지이용분류'] = df['토지이용분류'].str.strip()
    logger.info(f"CN 테이블 로드: {len(df)}개 항목 ({path})")
    return df


def match_cn(cn_table: pd.DataFrame, land_use_name: str, hydro_type: str):
    """
    토지이용분류 이름 + 수문학적 토양군으로 CN값 매칭.
    반환: int 또는 None (매칭 실패 시)
    """
    if not land_use_name or not hydro_type:
        return None

    land_use_name = str(land_use_name).strip()
    hydro_type = str(hydro_type).strip().upper()

    if hydro_type not in HYDRO_COLUMNS:
        logger.warning(f"알 수 없는 토양군 코드: '{hydro_type}'")
        return None

    matched = cn_table[cn_table['토지이용분류'] == land_use_name]
    if matched.empty:
        return None

    val = matched.iloc[0][hydro_type]
    if pd.isna(val):
        return None
    return int(val)


def apply_cn_to_layer(layer, cn_table: pd.DataFrame, level: str = 'l1'):
    """
    레이어의 각 피처에 cn값 컬럼 값을 채워 넣음.
    반환: 매칭 실패 목록 [(소유역명, 토지이용, 토양군), ...]
    """
    fail_list = []
    layer.startEditing()

    cn_idx = layer.fields().indexOf("cn값")
    land_use_idx = layer.fields().indexOf("토지이용")
    hydro_idx = layer.fields().indexOf("토양군")
    name_idx = layer.fields().indexOf("소유역명")

    for feat in layer.getFeatures():
        land_use = feat.attributes()[land_use_idx]
        hydro_type = feat.attributes()[hydro_idx]
        watershed_name = feat.attributes()[name_idx]

        cn_val = match_cn(cn_table, land_use, hydro_type)
        if cn_val is None:
            fail_list.append((watershed_name, land_use, hydro_type))
            logger.warning(
                f"CN 매칭 실패: 소유역={watershed_name}, "
                f"토지이용={land_use}, 토양군={hydro_type}"
            )
        else:
            layer.changeAttributeValue(feat.id(), cn_idx, cn_val)

    layer.commitChanges()
    logger.info(f"CN값 적용 완료. 실패: {len(fail_list)}건")
    return fail_list
