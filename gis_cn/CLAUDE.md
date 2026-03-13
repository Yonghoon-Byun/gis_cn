# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 프로젝트 개요

QGIS 3.x 플러그인. 소유역 폴리곤(SHP/GPKG)을 입력받아 PostGIS DB에서 수문학적 토양군(`public.soil`)과 토지피복도(`public.land_cover_yangju`)를 추출한 뒤, 공간 교차 연산(Clip → Intersection)으로 분할된 폴리곤에 CN값을 매칭하여 `CN값_input` 레이어를 생성한다.

## 플러그인 등록

QGIS Plugin Manager에서 로드하거나, 플러그인 경로에 심볼릭 링크로 연결한다.

```
mklink /D "%APPDATA%\QGIS\QGIS3\profiles\default\python\plugins\gis_cn" "S:\11_QGIS\07_gis_water\gis_cn"
```

또는 QGIS → 플러그인 관리 → 설정 → 플러그인 경로에 `S:\11_QGIS\07_gis_water` 추가 후 `GIS CN값 계산기` 활성화.

## 코드 구조와 데이터 흐름

```
plugin.py                    → 툴바/메뉴 등록, dialog.py 호출
dialog.py                    → UI 컨트롤러 (CnCalculatorDialog + CnWorker QThread)
core/db_manager.py           → PostGIS psycopg2 쿼리 → QgsVectorLayer(memory) 반환
core/spatial_ops.py          → QGIS Processing (native:intersection) → 레이어 반환
core/cn_matcher.py           → pandas로 cn_value.xlsx 로드, hydro_type+토지이용명으로 CN값 매칭
core/result_calculator.py    → CN값_input 레이어 → result1/result2 + 유역합성 계산 및 openpyxl 내보내기
core/land_use_mapper.py      → land_use_mapping.json 로드/저장, 레이어에 매핑 적용
core/local_data_handler.py   → 로컬 SHP/GPKG 로드, 한국어 컬럼 자동 감지 + canonical 리네이밍, Clip/Intersection
core/watershed_group.py      → 유역합성 그룹 JSON (watershed_groups.json) 로드/저장
reference/cn_calculator.ui   → Qt Designer XML (4탭 구조)
```

### 전체 워크플로우 (탭별 순서)

```
Tab 0 레이어 불러오기  → CnWorker ①②③④ → 토양군_clip, 토지피복도_clip, 토양군_토지피복_교차 캔버스 표시
                         final_intersect를 self._final_intersect_layer에 저장
Tab 1 토지이용 재분류  → land_use_mapping.json 편집/저장
Tab 2 CN값 편집       → cn_value.xlsx 기본값 확인, 커스텀 CN표 불러오기/내보내기
Tab 3 CN값 계산       → ⑤ _build_result_layer → ⑥ 매핑 적용 → ⑦ CN매칭 → CN값_input 생성
                         결과 내보내기: results.xlsx (result1 시트 + result2 시트)
```

### CnWorker 실행 순서 (백그라운드 스레드, Tab 0)

각 단계 완료 시 `layer_ready(layer, name)` 시그널로 메인 스레드에 즉시 전달 → 캔버스에 즉시 표시.

```
① PostGIS 토양군 Clip     → ST_Intersection으로 잘린 geometry 반환 → 토양군_clip 즉시 표시
② PostGIS 토지피복도 Clip → ST_Intersection으로 잘린 geometry 반환 → 토지피복도_clip 즉시 표시
③ PostGIS Intersection    → get_soil_lc_intersection() 단일 쿼리 (Dissolve 포함) → 토양군_토지피복_교차 즉시 표시
④ QGIS Intersection(③×소유역계) → final_intersect → finished 시그널로 dialog에 전달
   (CN값_input 생성은 Tab 3에서 수행)
```

**중요:** `finished(final_intersect_layer, [])` 시그널로 교차 레이어를 전달한다. `_on_finished()`에서 `self._final_intersect_layer`, `self._last_level`, `self._last_name_field`에 저장한다.

## DB 연결 정보

`core/db_manager.py`의 `DB_CONFIG` 딕셔너리에 하드코딩되어 있다.

- Host: `geo-spatial-hub.postgres.database.azure.com:6432`
- DB: `dde-water`, Schema: `public`
- 토양군: `public.soil` — 매칭 컬럼: `hydro_type` (A/B/C/D)
- 토지피복: `public.land_cover_yangju` — 분류 컬럼: `l1_code/l1_name`, `l2_code/l2_name`, `l3_code/l3_name`
- 기본 SRID: `5186` (EPSG:5186, Korea TM)

## CN값 매칭 로직

`cn_value.xlsx` 컬럼 구조: `토지이용분류 | A | B | C | D`

- 행 매칭: `토지이용분류` == `l1_name`/`l2_name`/`l3_name` (분류 선택에 따라)
- 열 선택: `hydro_type` 값 (A/B/C/D)
- 매칭 실패 시 `cn값` = NULL, 실패 목록은 로그창에 출력
- **CN 계산 시 CN표 우선순위**: Tab 2(CN값 편집) 위젯 테이블 → 비어있으면 `cn_value.xlsx` 폴백

## PostGIS 공간 처리 구조 (db_manager.py)

DB에서 공간 처리를 최대한 수행하여 QGIS Processing 단계를 최소화한다.

### `get_soil_layer(wkt, srid)` / `get_land_cover_layer(wkt, srid, level)`
- `ST_Intersects`(공간 인덱스 활용)로 필터 후 `ST_Intersection`으로 소유역 내부만 반환
- `ST_CollectionExtract(..., 3)`으로 Polygon 부분만 추출
- 반환 레이어가 곧 clip 결과 — QGIS `native:clip` 단계 불필요

### `get_soil_lc_intersection(wkt, srid, level)`
PostGIS CTE 단일 쿼리로 아래를 일괄 처리:
1. `sc` CTE: 토양군 clip (ST_Intersection)
2. `lc_raw` CTE: 토지피복도 clip (ST_Intersection)
3. `lc` CTE: level 기준 ST_Union(Dissolve) — l1:(l1_code,l1_name), l2:(l2_code,l2_name), l3:(l3_code,l3_name)
4. `ix` CTE: sc × lc ST_Intersection
- QGIS `native:dissolve` + `native:intersection` 2단계를 대체

`_LC_DISSOLVE_COLS` 딕셔너리가 level별 dissolve 컬럼을 정의한다.

## Intersection 후 컬럼 중복 처리

`native:intersection` 실행 후 동일 컬럼명에 숫자 접두어가 붙을 수 있다 (예: `hydro_type` → `2_hydro_type`). `spatial_ops._get_field_value()`에서 `endswith(f"_{target}")` 패턴으로 탐색하여 처리한다.

## UI 구조 (4탭)

탭 0·2는 `.ui` 파일에서, 탭 1은 `_setup_mapping_tab()`으로 동적 삽입(`insertTab(1,...)`), 탭 3은 `.ui` 파일의 원래 탭 2(인덱스 이동).

| 인덱스 | 탭명 | 위젯명/생성 | 주요 위젯 |
|--------|------|------------|---------|
| Tab 0 | 레이어 불러오기 | `tabCnCalc` (.ui) | rbFile/rbLayer, leFilePath, cmbNameField, rbL1/L2/L3, progressBar, txtLog, btnRun, btnClose, btnNextStep0 |
| Tab 1 | 토지이용 재분류 | `_setup_mapping_tab()` (동적) | tblMapping, btnMappingLoadLayer, btnMappingAddRow, btnMappingDeleteRow, btnMappingSave, btnMappingClear, btnNextStep1 |
| Tab 2 | CN값 편집 | `tabCnEdit` (.ui) | tblCnValues, btnAddRow, btnDeleteRow, btnAddColumn, btnReloadCn(기본값), btnImportCn(불러오기), btnSaveCn(내보내기), btnNextStep2 |
| Tab 3 | CN값 계산 | `tabRecalc` (.ui, 동적 삽입 후 index 3) | btnApplyCn(CN값 계산 실행, 동적), cmbRecalcLayer, leOutputDir, btnOutputDir, txtRecalcLog, btnExportResult1(결과 내보내기) |

### 탭 인덱스 상수 (dialog.py)
```python
TAB_CALC    = 0   # 레이어 불러오기
TAB_MAPPING = 1   # 토지이용 재분류 (insertTab(1,...) 으로 동적 삽입)
TAB_CN_EDIT = 2   # CN값 편집 (삽입 후 원래 index 1이 2로 이동)
TAB_RECALC  = 3   # CN값 계산 (원래 .ui의 tab index 2가 3으로 이동)
```

`_init_ui()` 실행 순서:
1. `_setup_mapping_tab()` → `insertTab(1, ...)` + `setTabText(3, "CN값 계산")`
2. `_enhance_recalc_tab()` → `widget(TAB_RECALC).layout().insertWidget(0, card)` 로 CN값 계산 카드 추가

### 탭별 로직
- **Tab 1 (토지이용 재분류)**: 최초 진입 시 `_mapping_load_saved()` → `land_use_mapping.json` 자동 로드. `btnMappingLoadLayer`는 캔버스의 **토양군_토지피복_교차** 레이어에서 분류 수준(`_get_level()`)에 따라 `l1_name`/`l2_name`/`l3_name` 컬럼의 고유값을 불러와 원본 열에 채움.
- **Tab 2 (CN값 편집)**: `cn_value.xlsx`를 기본값으로 보호. `btnReloadCn`(기본값) → cn_value.xlsx 초기화. `btnImportCn`(불러오기) → 파일 선택 로드. `btnSaveCn`(내보내기) → 파일 선택 저장. CN 계산 시 위젯 테이블 데이터 우선 사용(`_get_cn_table_from_widget()`).
- **Tab 3 (CN값 계산)**: `btnApplyCn`(CN값 계산 실행) → `_apply_cn_calc()` → `self._final_intersect_layer` 기반 ⑤⑥⑦ 실행 → `CN값_input` 생성. `btnExportResult1`(결과 내보내기) → `results.xlsx` (result1+result2 시트 통합).

`tblCnValues`는 탭 최초 진입 시 `cn_value.xlsx`를 자동 로드한다 (`_cn_table_loaded` 플래그로 중복 로드 방지).

### 다음 단계 버튼
탭 0·1·2 하단에 `btnNextStep0/1/2` 버튼이 있어 클릭 시 다음 탭으로 이동한다. Tab 3에는 없음.

### Enter키 → 다음 셀 이동
`_NextCellDelegate` (QStyledItemDelegate 서브클래스)가 `tblCnValues`와 `tblMapping`에 설정되어 있다. Enter/Return 키 입력 시 현재 셀 편집을 완료하고 오른쪽 다음 셀(행 끝이면 다음 행 첫 셀)로 자동 이동한다.

## 토지이용 재분류 (land_use_mapper.py)

`core/land_use_mapper.py` — 매핑 파일: `<plugin_root>/land_use_mapping.json`

- `load_mapping()` → `{원본명: 재분류명}` dict
- `save_mapping(mapping)` → JSON 저장
- `apply_mapping_to_layer(layer, mapping)` → `provider.changeAttributeValues()` 사용 (스레드 안전), 변경 건수 반환

매핑 적용 시점:
- `_apply_cn_calc()` ⑥단계: `load_mapping()` 후 `apply_mapping_to_layer()` → CN매칭 전 적용

## UI 디자인 스타일 가이드

레퍼런스: `reference/ui/main_dialog.py`, `reference/ui/region_tab.py`, `reference/ui/statistics_tab.py`

- 글로벌 폰트: Pretendard → Malgun Gothic 폴백
- 배경: `#f9fafb` (다이얼로그), `white` (카드)
- 색상 토큰: `#374151`(기본 텍스트), `#6b7280`(보조 텍스트), `#9ca3af`(비활성), `#1f2937`(액션 버튼)
- 카드 패턴: `background-color: white; border: 1px solid #e5e7eb; border-radius: 8px;` + 내부 padding 16/14px
- 헤더: `.ui`에 `frameHeader` QFrame, `white` bg, `border-bottom: 1px solid #e5e7eb`, height 56px
- 탭바: 언더라인 스타일 (`border-bottom: 2px solid #1f2937` for selected)
- 주요 버튼: `background-color: #1f2937; color: white; font-weight: 600;` (btnRun, btnSaveCn, btnExportResult1, btnApplyCn, btnMappingSave)
- 다음 단계 버튼: `border: 1px solid #374151; color: #374151; background: white; font-weight: 600;`
- 보조 버튼: `border: 1px solid #d1d5db; color: #374151; background: white;`
- 진행바: height 6px, textVisible false, chunk색 `#1f2937`
- 테이블 행 높이: 32px (`verticalHeader().setDefaultSectionSize(32)`)

## result_calculator.py — AMC2/AMC3 계산 로직

`CN값_input` 레이어(컬럼: 소유역명, 토지이용, 유역면적, 토양군, cn값)에서 계산.

**토지이용별 (result1 행 단위):**
- `AMC2_lu` = Σ(area_type × CN_type) / 총면적_lu (가중평균)
- `AMC3_lu` = 79 (논/답 고정) | `trunc(23×AMC2_lu / (10+0.13×AMC2_lu))` (기타)

**소유역 요약 (result1 합계행, result2):**
- `AMC2_ws` = Σ(area_lu × CN_lu) / 소유역총면적 (전체 가중평균)
- `AMC3_ws` = Σ(AMC3_lu × area_lu) / 소유역총면적 (토지이용별 AMC3의 면적 가중평균 → float)

**토지이용 목록**: 고정 13개 대신 **실제 데이터가 있는 토지이용만** 가나다순 정렬 (`sorted(lu_map.keys())`)

**결과 내보내기**: `export_results(result1_data, result2_data, path)` → `results.xlsx` 단일 파일
- `result1` 시트: 소유역별 토지이용×토양군 상세표 (TYPE A/B/C/D 열쌍)
- `result2` 시트: 유역별 CN값 요약 (유역, 총면적, AMC2 CN, AMC3 CN)

## result_calculator.py — NULL/NaN 처리

- `_is_null(v)`: Python `None`, QVariant NULL, `float NaN` 모두 null로 판정
- `layer_to_dataframe()`: `_is_null()`로 안전 변환 — NaN 면적은 `0.0`, NaN cn값은 `None`
- `_amc3()`: `math.isnan(amc2)` 가드 추가 — NaN 입력 시 `0` 반환 (`math.trunc(NaN)` 크래시 방지)
- `calculate_results()`: 반환값 `(result1_data, result2_data, null_cn_rows)` 3-tuple
  - `null_cn_rows`: 면적 > 0이면서 cn값 NULL인 행의 `[소유역명, 토지이용, 토양군]` 목록
- `dialog._warn_null_cn()`: null_cn_rows가 있으면 `QMessageBox.warning` 팝업 표시

## 로컬 데이터 입력 (local_data_handler.py)

DB 대신 로컬 SHP/GPKG 파일에서 토양군/토지피복도를 로드하는 기능.

### 데이터 소스 전환

Tab 0에 "데이터 소스" 카드가 동적 삽입된다. `rbSourceDB` (기본) / `rbSourceLocal` 라디오 버튼으로 전환.
- DB 모드: 기존 PostGIS 워크플로우 (변경 없음)
- 로컬 모드: `leSoilPath` + `leLcPath` 파일 브라우저 활성화

### CnWorker 분기

`CnWorker.__init__(input_layer, name_field, level, data_source='db', soil_path=None, lc_path=None)`

`run()` 내부에서 `data_source` 값에 따라 분기:
- `'db'`: 기존 `db_manager` 함수 호출 (get_soil_layer, get_land_cover_layer, get_soil_lc_intersection)
- `'local'`: `local_data_handler` 함수 호출 (load_local_soil, load_local_land_cover, get_local_soil_lc_intersection)

두 모드 모두 dissolve → ①②③④ 4단계 동일 구조, 동일 시그널 발행.

### 한국어 컬럼 자동 감지

`SOIL_COLUMN_ALIASES` / `LC_COLUMN_ALIASES` 딕셔너리로 canonical ↔ 한국어 변형 매핑.
- `_resolve_columns(layer, alias_map)` → `{canonical: actual_field_name}` 자동 감지
- `_rename_columns(layer, column_map)` → canonical 이름으로 리네이밍된 메모리 레이어 반환
- 감지 실패 시 `ValidationError` + 기대/실제 컬럼명 안내

### 알려진 토양군 별칭
`hydro_type`, `HYDRO_TYPE`, `HYDGRP`, `수문학토양군`, `토양군`, `HSG`

### 알려진 토지피복도 별칭
`대분류코드/대분류명`, `중분류코드/중분류명`, `세분류코드/세분류명` 등

## 유역합성 (watershed_group.py + result_calculator.py)

소유역을 그룹화하여 합성 CN값을 계산하는 기능.

### 그룹 관리

`watershed_groups.json` 파일로 그룹 정의 영속 저장.
```json
{"00천1": ["H1", "H2", "H3"], "본류": ["H1", "H2", "H3", "H4", "H5"]}
```
- `load_groups(path=None)` → `dict[str, list[str]]`
- `save_groups(groups, path=None)` → JSON 저장

### Tab 3 유역합성 설정 카드

Tab 3 (CN값 계산) 내부에 "유역합성 설정" 카드가 동적 삽입된다 (CN값 계산 카드 바로 아래).
- `QTableWidget` 21열: 그룹명 + 소유역1~20
- 소유역 열은 `QComboBox` (편집 가능) — Tab 0 실행 후 입력 레이어의 소유역명이 드롭다운에 자동 표시
- `_get_watershed_names()`: `self.input_layer`의 `self._last_name_field` 고유값 목록 반환
- 행 추가/삭제, 저장/불러오기 버튼
- 저장 시 `watershed_groups.json`에 저장

### 합성 CN 계산

`_calculate_watershed_cn(ws_name, ws_df)` 헬퍼: 단일 유역(또는 합성)의 CN값 계산.
- `calculate_results()`: 기존 시그니처 유지, 내부적으로 헬퍼 호출
- `calculate_grouped_results(layer, groups)`: 그룹별 합성 CN 계산 (동일 헬퍼 재활용)

### results.xlsx 합성 결과

`export_results(r1, r2, path, *, grouped_result1=None, grouped_result2=None)`:
- keyword-only 파라미터로 기존 호출 호환 유지
- result1 시트: 개별 소유역 블록 → "【유역합성】" 구분 → 합성 블록 (연두색 fill)
- result2 시트: 개별 행 → "【유역합성】" 구분 → 합성 행 (연두색 fill)

## 외부 의존성

- `psycopg2`: PostGIS 연결
- `pandas` + `openpyxl`: cn_value.xlsx 읽기/쓰기
- QGIS Processing framework: `native:dissolve`, `native:intersection` (clip은 PostGIS로 대체됨)

## 주의사항

- Processing 알고리즘은 반드시 QGIS 환경 내에서 실행해야 하며, 독립 Python 스크립트로 실행 불가
- `build_cn_input_layer()`는 현재 `dialog.py`에서 직접 사용하지 않고, 내부 함수(`intersect_layers`, `_build_result_layer`)를 개별 호출한다. `build_cn_input_layer()`는 보존하되 레거시로 간주
- `_build_result_layer()`는 접두어(`_`)가 붙어있지만 `dialog.py`에서 직접 import하여 사용 중
- `clip_layer()`, `dissolve_land_cover()`는 `spatial_ops.py`에 보존되어 있으나 `dialog.py`에서는 더 이상 호출하지 않음 (PostGIS로 대체)
- `export_result1()`, `export_result2()`는 `result_calculator.py`에 보존되어 있으나 `dialog.py`에서는 `export_results()`를 사용

---

## 개발 진행 현황

### 완료된 작업

| 날짜 | 작업 | 상태 |
|------|------|------|
| 2026-03-11 | 플러그인 기능 분석 및 교차 분석 (기획안 vs 구현) | 완료 |
| 2026-03-11 | GIS_CN값_계산기_활용매뉴얼.md 작성 (8장 775줄) | 완료 |
| 2026-03-11 | 사용법_가이드.md 작성 (간략 가이드) | 완료 |
| 2026-03-12 | 현재 레이어 콤보박스 실시간 갱신 버그 수정 (plugin.py, dialog.py) | 완료 |
| 2026-03-12 | cn_cal.xlsm VBA 매크로 분석 (AMC2/AMC3 계산 공식 검증) | 완료 |
| 2026-03-12 | CN산정 V5 xlsm 분석 (유역합성 구조 파악) | 완료 |
| 2026-03-12 | QGIS 설치 및 플러그인 설치 매뉴얼 작성 | 완료 |
| 2026-03-12 | gis_cn.zip 최신 버전 재생성 (버그 수정 반영) | 완료 |

### 완료 (2026-03-12)

계획서: `.omc/plans/local-data-and-watershed-grouping.md` (v3)

#### Feature 1: 로컬 데이터 입력 기능
- [x] Task 1: `core/local_data_handler.py` 신규 — 로컬 SHP/GPKG 로드 + 한국어 컬럼 자동 감지 + canonical 리네이밍
- [x] Task 2: `CnWorker`에 `data_source='local'` 분기 추가 (dialog.py)
- [x] Task 3: Tab 0 UI에 "DB에서 가져오기 / 로컬 파일 사용" 라디오 + 파일 찾아보기

#### Feature 2: 유역합성(Watershed Grouping) 기능
- [x] Task 4: `core/watershed_group.py` 신규 — 그룹 JSON 관리
- [x] Task 5: Tab 3 내부에 "유역합성 설정" 카드 삽입
- [x] Task 6: `result_calculator.py` — `_calculate_watershed_cn()` 헬퍼 추출 + `calculate_grouped_results()` 추가
- [x] Task 7: `export_results()` 확장 — results.xlsx에 합성 결과 포함
- [x] Task 8: dialog.py 통합 연결

### 알려진 이슈

- **로컬 데이터 컬럼명 불일치**: 한국 정부 SHP 파일은 DB와 다른 컬럼명 사용 (수문학토양군 vs hydro_type 등). 자동 감지 + alias 매핑으로 해결 예정. 실제 기관 데이터 확보 후 alias 목록 검증 필요.

### 미구현 기능 (기획안 대비)

- 삽도(지도 이미지) 자동 생성
- 한글(.hwp) 내보내기
