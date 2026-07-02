# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 프로젝트 개요

QGIS 3.x 플러그인. 소유역 폴리곤(SHP/GPKG)을 입력받아 PostGIS DB에서 수문학적 토양군(`public.soil`)과 토지피복도(`public.land_cover`)를 추출한 뒤, 공간 교차 연산(Clip → Intersection)으로 분할된 폴리곤에 CN값을 매칭하여 `CN값_input` 레이어를 생성한다.

## ⚠️ 긴급 해결 과제 (2026-07-02 갱신) — 도시부 토양군 결손 (원본 데이터 특성)

**증상:** 특정 유역 SHP로 실행 시 `토양군_clip`에 대면적 빈 구멍 → CN 산정이 빈값/실패로 들어옴. (사용자 제보 + 스크린샷 IMG_8774, 광주 인근 유역)

**근본 원인(전면 재검증 완료):** 원본 토양자료(`Soil_Type_0905_E5186_Join_K`, 국가 토양유형도)가 **도시(시가화)·수역·비경작지엔 토양 폴리곤이 없다.** 도시가 많은 유역은 그만큼 토양군이 비어 들어온다.
- **원본 SHP ↔ DB soil 대조: 피처 1,169 = 1,169, 범위·컬럼 동일, 광주 유역 커버리지 42.7% = 42.7% (1 m² 오차) → ETL 무결(누락 없음).**
- 빈 공간 구성(광주 유역): 시가화 40% + 농업 28% + 초지 19% … (수역은 1.5%뿐).
- 결론: **DB·ETL·좌표계·플러그인 계산 모두 정상. 원인은 원본 토양유형도의 도시 미매핑(데이터 본질).**

**❌ 폐기된 오진(2026-06-30자 CRS 진단은 틀림):** "입력 `.prj` 미정의로 5186 오인" 아님. 입력은 **5186(as-is)이 정답**(브이월드 육안 확인). 5174로 바꾸면 오히려 약 100km(전남↔전북) 어긋남. (5174에서 토양군 100%가 나온 건 '엉뚱한 위치에 마침 토양군이 있어서' — 커버리지를 CRS 정확도 판단 근거로 쓰면 안 됨.)

**해결할 일(원본은 못 바꾸므로 CN 계산에서 갭 처리):**
1. **[권장] 최근접 토양군 채움(nearest-fill)**: 토양군 없는 조각마다 가장 가까운 A/B/C/D 부여 → 도시부도 CN 산정 가능.
2. **[대안] 도시=불투수 기본 CN** 처리 또는 면적 임계 미만 무시.
3. **[부차] 커버리지 안내**: 토양군 덮음이 낮으면 "토양군 X% 미커버(도시부 등)" 사용자 안내.

**참고:** `.prj` 없는 입력의 CRS 미정의 검증은 별개의 일반 개선 항목(경고)일 뿐, 이번 사고 원인은 아님.

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

- Host: `geo-spatial-hub-prod.postgres.database.azure.com:6432`
- DB: `dde-water`, Schema: `public`
- 토양군: `public.soil` — 매칭 컬럼: `hydro_type` (A/B/C/D)
- 토지피복: `public.land_cover` — 분류 컬럼: `l1_code/l1_name`, `l2_code/l2_name`, `l3_code/l3_name`
- 기본 SRID: `5186` (EPSG:5186, Korea TM)

## CN값 매칭 로직

`cn_value.xlsx` 컬럼 구조: `토지이용분류 | A | B | C | D`

- 행 매칭: `토지이용분류` == `l1_name`/`l2_name`/`l3_name` (분류 선택에 따라)
- 열 선택: `hydro_type` 값 (A/B/C/D)
- 매칭 실패 시 `cn값` = NULL, 실패 목록은 로그창에 출력
- **CN 계산 시 CN표 우선순위**: Tab 2(CN값 편집) 위젯 테이블 → 비어있으면 `cn_value.xlsx` 폴백

> **⚠ CN표 출처 차이 (정합성 주의, 2026-06-05 확인):** 실제 CN 매칭에 쓰는 기본 폴백표
> `cn_value.xlsx`(13종 단일계층)는 보고서 표4-18에 인쇄되는 국가표준 **[표 4-2] 유출곡선지수
> (AMC-Ⅱ)**(`data/cn_reference_std.json`, 41행 대/중/세분류)와 **분류체계·일부 값이 다르다.**
> 예: `밭` 63/**75/83/87** vs 표4-2 63/74/82/85; `임야` 45/66/77/83(표4-2 산림 55/72/82/85≠);
> `초지` 39/61/74/80(표4-2 자연초지 30/58/71/78≠); `광장·주차장` 98 등은 표4-2에 없는 분류.
> **현 방침: `cn_value.xlsx` 값 보존**(기존 CN산정 V5 검증값일 수 있음) — 국가표준으로 재정합하려면
> 모든 산정 CN이 바뀌므로 별도 결정 필요. 보고서 표4-18은 국가표준(`cn_reference_std.json`)으로
> 별도 인쇄되며 실제 산정표와 다를 수 있음에 유의.

## PostGIS 공간 처리 (db_manager.py)

DB에서 Clip+Intersection을 처리 → QGIS Processing 최소화. Geometry는 WKB(바이너리)로 전송.

- **`get_all_layers(wkt, srid, level)`** (권장): 단일 DB 연결 + 임시 테이블로 ①②③ 일괄 처리. clip 중복 제거로 ~40% 속도 향상.
- `get_soil_layer`/`get_land_cover_layer`/`get_soil_lc_intersection`: 개별 함수 (하위호환용)
- `_LC_DISSOLVE_COLS`가 level별 dissolve 컬럼 정의
- **공간필터는 인덱스 친화 형태 필수**: WHERE는 `ST_Intersects(t.geom, w.geom_n)` — `w.geom_n`은 입력 윈도우를 테이블 native SRID(5186)로 **1회 변환한 상수**(CTE). `ST_Transform(w.geom, ST_SRID(t.geom))`처럼 행마다 `ST_SRID(t.geom)`를 부르면 상수로 안 잡혀 GIST 인덱스를 못 쓰고 **전체 Seq Scan** → land_cover(약 4,180만 행)에서 타임아웃. 출력 SELECT는 `ST_Transform(t.geom, %(srid)s)`로 입력 좌표계 출력 유지. (2026-06-29 수정 — land_cover_yangju→land_cover 교체로 표가 커지며 잠복 안티패턴 발현, opus 4에이전트 검증)

## Intersection 후 컬럼 중복 처리

`native:intersection` 실행 후 동일 컬럼명에 숫자 접두어가 붙을 수 있다 (예: `hydro_type` → `2_hydro_type`). `spatial_ops._get_field_value()`에서 `endswith(f"_{target}")` 패턴으로 탐색하여 처리한다.

## UI 구조 (5탭, 2026-06-29 갱신)

탭 0·1·3은 `.ui`, 탭 2(토지이용 재분류)는 `_setup_mapping_tab()`이 `insertTab(2,...)`로 동적 삽입, **탭 4(보고서 출력)는 `_setup_report_tab()`이 `addTab`으로 신설**. 탭 라벨에 단계 번호(①~⑤)와 완료 배지(✓)를 표시(`_refresh_tab_titles`/`_mark_step_done`, `_tab_done` 상태).

| 인덱스 | 탭명 | 위젯명/생성 | 주요 위젯 |
|--------|------|------------|---------|
| Tab 0 ① | 레이어 불러오기 | `tabCnCalc` (.ui) | rbFile/rbLayer, leFilePath, cmbNameField, rbL1/L2/L3, progressBar, txtLog, btnRun, btnClose, btnNextStep0(→②) |
| Tab 1 ② | CN값 편집 | `tabCnEdit` (.ui) | tblCnValues, btnAddRow, btnDeleteRow, btnAddColumn, btnReloadCn(기본값), btnImportCn(불러오기), btnSaveCn(내보내기), btnNextStep2(→③) |
| Tab 2 ③ | 토지이용 재분류 | `_setup_mapping_tab()` (동적 insertTab(2)) | tblMapping + 버튼, btnCnRefPopup(CN표 참조 팝업), btnMappingLoadLayer, btnMappingSave, btnMappingClear, btnNextStep1(→④) |
| Tab 3 ④ | CN값 계산 | `tabRecalc` (.ui) + 동적 | btnApplyCn, **leLayerName**, cmbRecalcLayer, txtRecalcLog, **CN값 메모리 관리 카드**, **btnNextStep3(→⑤, 2026-06-29 신설)** |
| Tab 4 ⑤ | **보고서 출력** | `_setup_report_tab()` (동적, addTab, `TAB_REPORT=4`) | leOutputDir/btnOutputDir, chkExportExcel/chkExportHwp, chkStaged, _stage_combos(메모리 선택), btnStagedPreview, tblStaged, **btnReportExport**(+진행바 `_export_progress`/상태 `_export_status`) |

### 탭 인덱스 상수 (dialog.py)
```python
TAB_CALC    = 0   # 레이어 불러오기
TAB_CN_EDIT = 1   # CN값 편집 (.ui 원본 위치 유지)
TAB_MAPPING = 2   # 토지이용 재분류 (insertTab(2,...) 으로 동적 삽입)
TAB_RECALC  = 3   # CN값 계산
TAB_REPORT  = 4   # 보고서 출력 (addTab 으로 추가)
```
> 시각 순서 = 레이어(0)→CN값 편집(1)→토지이용 재분류(2)→CN값 계산(3)→보고서(4). '다음 단계' 버튼(btnNextStep0/2/1/3)이 이 순서대로 0→1→2→3→4 이동.

`_init_ui()` 실행 순서:
1. `_setup_mapping_tab()` → `insertTab(2, ...)`
2. `_enhance_recalc_tab()` → `widget(TAB_RECALC).layout().insertWidget(0, card)` 로 CN값 계산 카드 추가
3. (init 말미) `_refresh_tab_titles()`(번호/배지) + `_add_recalc_next_button()`(④→⑤ 버튼)

### 탭별 핵심 로직
- **Tab 1**: `_mapping_load_saved()` → json 로드. `_MappingComboDelegate`로 CN표 드롭다운. `_CnRefDialog` 팝업(검색+더블클릭 자동입력). `_cn_ref_dirty` 플래그로 Tab 2 편집 시 자동 동기화.
- **Tab 2**: `cn_value.xlsx` 기본값 보호. CN 계산 시 위젯 테이블 우선(`_get_cn_table_from_widget()`). `_cn_table_loaded` 플래그로 중복 로드 방지.
- **Tab 3**: `btnApplyCn` → `_apply_cn_calc()` → ⑤⑥⑦ → `CN값_input` 생성. `btnExportResult1` → `results.xlsx`.
- **셀 이동**: `_NextCellDelegate`(Enter→다음셀), `_MappingComboDelegate`(Enter→다음행)
- **다음 단계**: `btnNextStep0/1/2` → 다음 탭 이동 (Tab 3에는 없음)

## 토지이용 재분류 (land_use_mapper.py)

`land_use_mapping.json` 기반. `load_mapping()` → dict, `apply_mapping_to_layer()` → `provider.changeAttributeValues()`. `_apply_cn_calc()` ⑥단계에서 CN매칭 전 적용.

## UI 스타일 가이드

레퍼런스: `reference/ui/main_dialog.py`, `reference/ui/region_tab.py`, `reference/ui/statistics_tab.py`
- 폰트: Pretendard → Malgun Gothic / 배경: `#f9fafb`, 카드: `white`, 카드 border: `#e5e7eb` radius 8px
- 색상: `#374151`(텍스트), `#1f2937`(액션버튼/탭선택), `#6b7280`(보조), `#9ca3af`(비활성)
- 버튼: 주요(`#1f2937 bg, white text`), 다음단계(`#374151 border outline`), 보조(`#d1d5db border`)
- 탭바: 언더라인(`border-bottom: 2px solid #1f2937`), 헤더: 56px white, 진행바: 6px `#1f2937`, 테이블 행: 32px

## result_calculator.py — AMC2/AMC3 계산 로직

`CN값_input` 레이어(컬럼: 소유역명, 토지이용, 유역면적, 토양군, cn값)에서 계산.

**토지이용별 (result1 행 단위):**
- `AMC2_lu` = Σ(area_type × CN_type) / 총면적_lu (가중평균)
- `AMC3_lu` = 79 (논 고정) | `trunc(23×AMC2_lu / (10+0.13×AMC2_lu))` (답 포함 기타)

**소유역 요약 (result1 합계행, result2):**
- `AMC2_ws` = Σ(area_lu × CN_lu) / 소유역총면적 (전체 가중평균)
- `AMC3_ws` = Σ(AMC3_lu × area_lu) / 소유역총면적 (토지이용별 AMC3의 면적 가중평균 → float)

**토지이용 목록**: 고정 13개 대신 **실제 데이터가 있는 토지이용만** 가나다순 정렬 (`sorted(lu_map.keys())`)

**결과 내보내기**: `export_results(result1_data, result2_data, path)` → `results.xlsx` 단일 파일
- `result1` 시트: 소유역별 토지이용×토양군 상세표 (TYPE A/B/C/D 열쌍)
- `result2` 시트: 유역별 CN값 요약 (유역, 총면적, AMC2 CN, AMC3 CN)

**NULL/NaN 처리**: `_is_null()` 통합 판정(None/QVariant NULL/NaN). `calculate_results()` → 3-tuple `(result1, result2, null_cn_rows)`. null_cn_rows 있으면 `_warn_null_cn()` 팝업.

## 로컬 데이터 입력 (local_data_handler.py)

Tab 0 `rbSourceDB`/`rbSourceLocal` 라디오로 데이터 소스 전환. `CnWorker(data_source='db'|'local')`로 분기, 동일 ①②③④ 구조.

**한국어 컬럼 자동 감지**: `SOIL_COLUMN_ALIASES`/`LC_COLUMN_ALIASES`로 canonical↔한국어 매핑. `_resolve_columns()` → 자동 감지, `_rename_columns()` → canonical 리네이밍. 실패 시 `ValidationError`.
- 토양군 별칭: `hydro_type`, `HYDGRP`, `수문학토양군`, `토양군`, `HSG`
- 토지피복도 별칭: `대분류코드/대분류명`, `중분류코드/중분류명`, `세분류코드/세분류명` 등

## 유역합성 (watershed_group.py + result_calculator.py)

`watershed_groups.json`으로 소유역 그룹 관리(`load_groups`/`save_groups`). Tab 3에 유역합성 설정 카드(21열 테이블, ComboBox 드롭다운).

**계산**: `_calculate_watershed_cn()` 헬퍼 → `calculate_results()`/`calculate_grouped_results()` 재활용.
**내보내기**: `export_results(..., *, grouped_result1=None, grouped_result2=None)` — keyword-only로 기존 호환. result1/result2 시트에 "【유역합성】" 구분 + 연두색 fill.

## 외부 의존성

- `psycopg2`: PostGIS 연결
- `pandas` + `openpyxl`: cn_value.xlsx 읽기/쓰기
- QGIS Processing framework: `native:dissolve`, `native:intersection` (clip은 PostGIS로 대체됨)

## 배포 규칙

**플러그인 소스 수정 후 반드시 `dist/gis_cn.zip` 재생성할 것.** 배포 zip이 항상 최신 소스를 반영해야 한다.

```
프로젝트 구조:
  gis_cn/       → 플러그인 소스 (순수 배포 대상)
  docs/         → 매뉴얼, 기획안, 참고자료
  scripts/      → DB 마이그레이션 등 유틸리티
  dist/         → 배포용 zip (gis_cn.zip)
```

## 주의사항

- Processing 알고리즘은 반드시 QGIS 환경 내에서 실행해야 하며, 독립 Python 스크립트로 실행 불가
- `build_cn_input_layer()`는 현재 `dialog.py`에서 직접 사용하지 않고, 내부 함수(`intersect_layers`, `_build_result_layer`)를 개별 호출한다. `build_cn_input_layer()`는 보존하되 레거시로 간주
- `_build_result_layer()`는 접두어(`_`)가 붙어있지만 `dialog.py`에서 직접 import하여 사용 중
- `clip_layer()`, `dissolve_land_cover()`는 `spatial_ops.py`에 보존되어 있으나 `dialog.py`에서는 더 이상 호출하지 않음 (PostGIS로 대체)
- `export_result1()`, `export_result2()`는 `result_calculator.py`에 보존되어 있으나 `dialog.py`에서는 `export_results()`를 사용

---

## 알려진 이슈

- **로컬 데이터 컬럼명 불일치**: 한국 정부 SHP 파일은 DB와 다른 컬럼명 사용. alias 매핑으로 해결 중, 실제 기관 데이터 확보 후 검증 필요.
- **CN값 계산 검증 필요**: 유역합성 계산값이 정확한지 기존 엑셀(CN산정 V5)과 비교 검증 예정.
- **CN 매칭표 ≠ 국가표준 표4-2**: 실제 계산용 `cn_value.xlsx`(13종)가 보고서 표4-18 국가표준표(`data/cn_reference_std.json`, 41행)와 분류·일부 값이 다름(밭 B/C/D, 임야, 초지 등). 현재는 `cn_value.xlsx` 값 보존 방침이며 차이만 문서화(위 "CN값 매칭 로직" 경고 박스 참조). 국가표준 재정합은 고영향 변경이라 보류.

## 미구현 기능

- 토지피복도 커스텀 분류 (L1/L2/L3 혼합 분류 — A안 확정, 미구현)
- 삽도(지도 이미지) 자동 생성
- ~~개발 전/중/후 3단계 dialog UI~~ → **구현 완료**(2026-06-05, 아래 "최근 주요 변경 (2026-06-05)" 참조). 잔여 미구현: 단계별 SHP 자동입력 UI, meta 표지 누름틀.
- 표4-19 **자동 페이지 분할 + 단계/소유역 진짜 셀병합 + 정수포맷 구현 완료**(2026-06-05): ①단계별 섹터 분리 + 행수 예산(`RES_ROWS_PER_TABLE=28`) 초과 시 소유역 블록 경계 물리 분할(`_plan_table_groups`는 단계 바뀌면 새 표, `_render_split_res`는 표 포함 `<hp:p>` 복제 + `pageBreak='1'` + repeatHeader). ②단계(col0)/소유역(col1) **진짜 cellSpan rowSpan 세로 병합**(`_merge_res_columns`: 구간 첫 행 rowSpan=k + 나머지 행 해당 tc 제거, 표분할이 단계/소유역 경계라 한 표 안에서 병합 완결). ③숫자 **정수+천단위 콤마**(`_fmt_num`, 엑셀 `#,##0` 일치 — 표4-19만, 표4-17 증감은 소수 유지). 골든 `test_merge`/`test_staged_split` GREEN, 원본 대조 통과. 한글 육안 게이트 필요.

## HWP/Excel 공통 입력 모델

`core/analysis_result.py`의 `AnalysisResult` dataclass가 **Single Source of Truth**. Excel/HWP 양쪽 렌더러는 이 모델만 입력으로 받는다.

- `core/analysis_result.py` — dataclass 정의 (`AnalysisResult`, `WatershedBlock`, `LandUseRow`, `WatershedSummary`, `CnReferenceRow`, `ProjectMeta`, `MapImage`, `NullRow`)
- `core/result_calculator.py::build_analysis_result()` — 기존 `calculate_results()` 반환(list[dict])을 `AnalysisResult`로 변환
- `core/result_calculator.py::export_excel(result, path)` — 기존 `export_results()`의 얇은 래퍼 (HWP 렌더러와 동일 시그니처)
- `core/hwp_renderer.py::render_hwp(result, template, out)` — pyhwpx(OLE) 래퍼. 한글 미설치/미등록 시 `HwpRendererError`
- `templates/v1.0/README.md` — HWP 템플릿 제작 스펙(누름틀/책갈피/셀필드 ID)

UI는 결과 저장 버튼 위에 "출력 포맷" 체크박스 + HWP 템플릿 경로를 추가(`_setup_output_format_row()`). Excel/HWP 동시 선택 시 순차 실행하며, HWP 실패해도 Excel은 보존(부분 실패 허용).

## 한글(.hwpx) 보고서 출력 — 순수 Python 렌더러 (한컴오피스 불필요)

`core/hwp_renderer.py`(pyhwpx/OLE)는 **레거시**다(한컴오피스 필요·COM 불안정). 현재 출력 경로는
**`core/hwpx_writer.py`** — lxml 만으로 HWPX(OWPML zip+XML)를 직접 작성한다. `dialog.py` 는
`hwpx_writer.render_hwpx` 를 호출(.hwp → .hwpx).

- `core/hwpx_writer.py::render_hwpx(result, template, out)` — 단일단계.
- `core/hwpx_writer.py::render_staged_report(report, template, out)` — **개발 전/중/후 3단계**(표4-17 증감 비교 + 표4-19 단계 컬럼).
- 동적 표: 프로토타입 행을 결과 수만큼 복제 + `rowAddr` 재번호(`_render_dynamic_table`). 사전할당-trim 방식 폐기.
- `scripts/build_template.py` — 샘플 hwpx → 클린 템플릿(`templates/v1.0/cn_report.hwpx`, 표4-17/18/19 3개) 빌더. 쓰레기 표 제거·res 축소·표 이식·15mm 여백·secPr 보정·검증.
- 데이터모델: `StagedReport{stages, stage_order, cn_delta_rows}`, `CnDeltaRow`(표4-17 소유역별 전/중/후 CN+증감+사유). 계산엔진과 **직교**(옵션 C) — `result_calculator.assemble_staged_report()`/`build_cn_delta()`.

### 한글 HWPX 렌더링 핵심 함정 (반드시 준수 — "구조 정합 ≠ 한글 렌더 성공")

zip OK + XML well-formed + id unique 라도 한글이 거부할 수 있다. 아래 불변식을 지켜야 한다:

1. **secPr-run = `[secPr, ctrl(colPr)]`**: 섹션정의 run 은 secPr 뒤에 colPr(단 설정) 컨트롤을 동반해야 한다. secPr 만 든 run, 또는 표(tbl)와 같은 run 에 secPr 가 끼이면 → 한글이 섹션정의를 폐기하고 **전 페이지가 기본 여백**(위20/머리15/왼30/오30/꼬리15/아래15)으로 떨어진다. (XML 의 margin 값은 정상이어도 한글이 안 읽음.)
2. **linesegarray 불변식**: `<hp:linesegarray>`(라인 세그먼트 맵)가 있는 단락에서 `<hp:run>`을 **삭제·이동하면** 라인맵이 없는 run 을 참조해 한글이 **"문서가 손상되었거나 변조" 경고**를 띄운다. → 다른 단락의 요소(colPr 등)를 옮기지 말고 **그 자리에 새로 합성(synthesize)** 하라. (`relocate_secpr` 가 colPr 를 이동→삭제하지 않고 표준 단일 단 colPr 를 새로 만드는 이유.)
3. **fieldBegin/fieldEnd 짝**: 고아 fieldEnd(짝 없는 beginIDRef) → 동일한 "손상/변조" 경고. 누름틀 제거 시 begin/end 쌍을 반드시 함께. `build_template._validate()` 가 짝·id중복을 검사.
4. **borderFill ID 해석**: 표 이식 시 `borderFillIDRef` 는 ID 가 아니라 **정의로 해석**된다(같은 ID 라도 문서마다 정의가 다름). 원본 정의를 fresh ID 로 복사 + remap 해야 테두리가 유지(`_copy_table_styles`). 안 하면 테두리가 NONE→빨간 점선.
5. **mimetype**: zip **첫 엔트리** + **STORED(비압축)**. 쪽 여백 단위 HWPUNIT=1/7200inch (15mm=4252, 12.5mm=3543).

### 검증 도구 (한글 없이 회귀 차단)

- **`scripts/compare_to_original.py [출력.hwpx]`** — 원본 hwpx(`제4장 재해영향…최종.hwpx`)와 쪽여백·secPr 배치(colPr 동반 여부)·표 테두리 스타일을 **정량 비교**하는 회귀 게이트. secPr-run 에 colPr 없거나 표가 섞이면 ★불일치★.
- `scripts/test_hwpx_writer.py` — 골든 구조 테스트(필드·행수·id유니크·그리드 타일링). 한글/QGIS 불필요.
- `scripts/diag_integrity.py` / `diag_para_diff.py` — 손상/변조 원인(필드 짝·run 구조) 진단, 두 hwpx 단락별 diff.
- `scripts/render_sample.py [출력]` — 한글 육안 검증용 3단계 샘플 렌더.
- **시각/열림 확인은 자동화 불가** — 한글에서 직접 열어보는 사용자 육안 게이트 필수(개발 PC COM 은 -2147221005 로 실패).

## 최근 주요 변경 (2026-06-29 추가) — UI 사용성·DB 명세서

- **워크플로우 가독성**: 탭 라벨 단계 번호 ①~⑤ + 완료 배지 ✓(`_refresh_tab_titles`/`_mark_step_done`/`_tab_done`). 레이어 로드·재분류 저장·CN계산·보고서 출력 완료 시 배지, 초기화 시 리셋. **CN값 계산(④)→보고서 출력(⑤) '다음 단계' 버튼(`btnNextStep3`) 신설**.
- **보고서 출력 스레드화**: `ExportWorker(QThread)` 신설 — 한글/엑셀 렌더(순수 lxml/openpyxl)를 UI 스레드에서 분리해 **출력 중 QGIS 프리징 제거**. QGIS 의존 계산(`calculate_*`/`_staged_build_report`)은 UI 스레드 유지, 렌더만 워커로. 진행바(`_export_progress`, 불확정)+상태 라벨+대기커서, 부분실패 허용 유지. 단일/3단계(`_export_results`/`_export_staged`) 모두 `_run_export_jobs` 경유. (워커 참조는 done 직후 None 처리 안 함 — QThread 조기 파괴 방지)
- (정정) "UI 구조" 탭 인덱스: 실제는 `TAB_CN_EDIT=1`/`TAB_MAPPING=2`(과거 문서가 1↔2 뒤바뀜).
- **DB 명세서**: `docs/DB_접속정보_및_테이블_명세서.md`(soil/land_cover 접속·컬럼·쿼리패턴·성능함정). **읽기전용 DB 비밀번호 포함 → PUBLIC 저장소 커밋 금지**, zip 등으로 별도 배포.

## 최근 주요 변경 (2026-06-29) — 답 AMC3·land_cover 교체+인덱스·안내문구

- **답 AMC3 일반식 전환**: `AMC3_FIXED_79 = {'논', '답'}` → `{'논'}`. 논만 AMC3=79 하드코딩 유지, **답은 일반식**(`trunc(23×AMC2/(10+0.13×AMC2))`)으로 계산. 단 `cn_value.xlsx`의 답 CN이 A~D 전부 79라 AMC2=79 → 결과적으로 AMC3=89(고정값 아님, 표값에 종속). 논·답은 동일 지목(벼논)임에 유의.
- **토지피복 테이블 교체 `land_cover_yangju` → `land_cover`**(운영DB, 전국 약 4,180만 행). `db_manager.py` 3개 SQL + CLAUDE.md + docs 매뉴얼 반영.
- **clip 쿼리 공간인덱스 복구(중요)**: 기존 `ST_Intersects(t.geom, ST_Transform(w.geom, ST_SRID(t.geom)))`는 행마다 `ST_SRID`라 상수가 아니어서 GIST 인덱스 미사용→전체 Seq Scan. land_cover가 작던(297k) 시절엔 잠복하다 41.8M 교체로 발현(2km 박스도 타임아웃). **수정**: 윈도우를 native 5186으로 1회 변환한 상수 `w.geom_n`을 WHERE에 사용(4함수 6술어). 출력 SELECT는 `ST_Transform(t.geom,%(srid)s)`로 입력 좌표계 보존. opus 4에이전트로 결과동일성·인덱스사용·정합성 검증. (위 "PostGIS 공간 처리" 경고 참조)
- **재분류 안내 문구**: 토지이용 재분류 탭 안내(`dialog.py` info_lbl)를 불릿 줄구분으로 가독성 개선 + "CN 산정 시 '논'은 79로 고정됩니다." 문구 추가.
- (참고) DB 진단: soil은 1,169행이지만 폴리곤당 평균 5만/최대 181만 정점(복잡도 병목), land_cover는 41.8M 볼륨 병목. 큰 유역계 최적화(soil ST_Subdivide / 규모 가드)는 보류.

## 최근 주요 변경 (2026-06-05) — 3단계 보고서/탭 대규모 개편

- **탭 분리(5탭)**: "보고서 출력" 탭 신설(`TAB_REPORT=4`, `_setup_report_tab`). CN값 계산 탭(3)=계산 + **CN값 메모리 관리**(`_setup_memory_list_card`), 보고서 출력 탭(4)=출력 설정·3단계 미리보기/사유·내보내기. 출력 위젯(leOutputDir·내보내기 버튼)을 탭4로 reparent, `_wrap_tab_in_scroll`가 `_recalc_content_layout` 보관.
- **메모리풀 모델**: CN값 계산 시 `leLayerName`(레이어 이름) 입력 → 그 이름으로 레이어 생성 + `_save_to_memory`로 계산결과 스냅샷을 `self._memory_pool[name]` 저장. 보고서 탭 개발 전/중/후 = `_stage_combos` 드롭다운으로 메모리 선택(`_staged_build_report`가 풀에서 조립). "단계로 저장" 버튼 폐기. `_refresh_recalc_layer_list`는 'cn값' 필드 보유 레이어로 필터(사용자 명명 포함).
- **표4-19 페이지/병합/포맷**: ①단계별 섹터 분리 + 행수예산(`RES_ROWS_PER_TABLE=28`) 초과 시 소유역 경계 물리 분할(`_plan_table_groups`/`_render_split_res`, 표 포함 `<hp:p>` 복제 + `pageBreak='1'` + repeatHeader, 복제표 누름틀 id 재배정). ②단계(col0)/소유역(col1) **진짜 cellSpan rowSpan 세로 병합**(`_merge_res_columns`). ③단계 컬럼 **세로쓰기**(`textDirection="VERTICAL"`). ④숫자 **정수+천단위 콤마**(`_fmt_num`=엑셀 `#,##0`, 표4-19만; 표4-17 증감은 소수 유지). ⑤캡션 **표 번호 제거**("[표 4-19]"→"[표]", autoNum 필드 제거, `_strip_table_numbers`).
- **Excel**: 3단계는 **단일 파일·단계별 시트**(`export_excel_staged`, `export_results`에 `wb`/`sheet_title` 옵션, `_ares_to_results` 추출). **열너비 13·배율 85%**(`defaultColWidth=13`, `zoomScale=85`). HWP 템플릿 입력 UI 제거(항상 내장 `cn_report.hwpx`).
- 골든 `scripts/test_hwpx_writer.py`(`test_merge`/`test_staged_split`)·`scripts/test_staged_dialog.py` + 원본대조 게이트 통과. **한글/엑셀 육안 게이트**(세로쓰기·캡션·손상경고·시트·서식) 필수.

## 최근 주요 변경 (2026-06-05 추가) — 유역합성·표 선스타일·UI 버그픽스

- **유역합성(composite) 출력 누락 수정 — 원인 3겹**: ①`_save_to_memory`가 합성 결과를 스냅샷에 안 담음(`calculate_grouped_results` + `build_analysis_result(grouped_result1/2)` 추가) → 3단계 Excel 누락 해소. ②`render_hwpx`/`render_staged_report`가 `composite_detail`을 아예 렌더 안 함(템플릿에 유역합성 전용 표 없음) → **표4-19(res.*)에 단계별로 합성 블록을 detail 뒤에 이어붙임**. ③3단계는 스냅샷만 사용 → **CN계산 *후* 그룹 정의 시 조용히 누락** → 내보내기 시점 스냅샷 레이어 + 현재 그룹으로 재계산 주입(`_refresh_composite_from_groups`, 단일단계 경로와 대칭). 레이어 부재/그룹 없음 시 graceful.
- **표 선스타일(행간 이중선) 수정**: 템플릿 프로토타입 1행은 헤더 바로 아래라 윗변이 이중선(`DOUBLE_SLIM`)인데, 이를 모든 데이터행에 복제하면 **행마다 이중선**이 그려져 원본과 달라짐. → **첫 데이터행만 이중선(헤더 구분) 유지, 2행부터 동일 좌/우/아래·동일 배경의 실선(`SOLID`)으로 교체**(`_Doc.border_solid_map` = header.xml borderFill 이중선→실선 매핑(좌·우·아래·fill 동일 보장), `_fix_row_separators`를 `_render_dynamic_table` 끝에서 호출 — 표4-17·표4-19 공통). 셀 배경(fill)까지 키에 포함해 음영 변형 방지.
- **Excel '합계' 라벨**: `export_results` 합계 행의 토지이용 칸이 비어 있던 것을 `'합계'`로 기입(개별 소유역 + 유역합성 블록 모두). 한글 렌더러(`_res_summary_row`)와 일치.
- **CN 계산 버튼 UI 씹힘(위젯 겹침) 수정**: `progressBarCalc`에 `RetainSizeWhenHidden(True)`(숨김 시에도 20px 자리 유지 → 계산 중 표시될 때 레이아웃 reflow 제거) + 계산 시작/종료 시 카드 강제 `repaint()/update()`(`processEvents` 반복 시 스테일 픽셀 잔상 정리).
- **출력 파일명 통일**: 단일·3단계 모두 `result.xlsx` / `result.hwpx`(주의: 같은 폴더에 동일 파일명 → 단일/3단계 둘 다 돌리면 덮어씀). **체크박스 라벨** '개발 전/중/후 3단계 비교…' → **'비교 검토 기능'**.
- 어드버서리얼 리뷰(4 에이전트)·골든 8종·py_compile 통과. `export_result1`/`export_result2`는 미사용 레거시(미수정).

## 최근 주요 변경 (2026-06-04~05)

- **한글 출력을 순수 Python(lxml) 렌더러로 전환** — 한컴오피스 불필요(`hwpx_writer.py`). `dialog.py` → `render_hwpx` 연결, UI `.hwp`→`.hwpx`.
- 개발 전/중/후 **3단계 비교 파이프라인**(옵션 C): `StagedReport`/`CnDeltaRow` 데이터모델 + `assemble_staged_report`/`build_cn_delta` + `render_staged_report`(표4-17 증감 비교, 표4-19 단계 컬럼).
- 동적 표(결과 수만큼 행 생성), 표4-18 국가표준 CN 기준표(8열) 이식, 15mm 쪽 여백.
- **한글 렌더 버그 3종 해결**: ① 테두리 NONE/빨간점선 → borderFill 정의 이식(`_copy_table_styles`) ② 쪽 여백 무시 → secPr-run 에 colPr 동반 ③ 손상/변조 경고 → linesegarray 가 있는 단락의 run 삭제 금지(colPr 합성). 위 "핵심 함정" 참조.
- 회귀 게이트 `scripts/compare_to_original.py`(원본 대조) + 진단 스크립트 추가.

## 최근 주요 변경 (2026-04-22)

- 버전 0.1 → 0.2 (Excel/HWP 공통 출력 모델 도입)
- `core/analysis_result.py` 신규 — `AnalysisResult` dataclass (SSOT)
- `core/hwp_renderer.py` 신규 — pyhwpx 기반 HWP 렌더러 스켈레톤 (지연 import)
- `core/result_calculator.py` — `build_analysis_result()`, `export_excel()` 추가 (기존 API 보존)
- `dialog.py` — 출력 포맷 선택 UI(Excel/HWP 체크박스 + 템플릿 경로) 추가
- `templates/v1.0/README.md` — HWP 템플릿 제작 스펙 추가

## 최근 주요 변경 (2026-04-08~09)

- DB 접속정보: 개발DB → 운영DB(geo-spatial-hub-prod) waterviewer 계정
- PostGIS 최적화: `get_all_layers()` 단일 연결 + 임시 테이블 + WKB 전송
- 소유역 입력 시 좌표계 자동 변환 (non-5186 → EPSG:5186)
- 초기화 버튼, CN계산 프로그레스바, Tab 3 QScrollArea 레이아웃
- 유역합성 카드: 체크박스 토글(기본 접힘) + CN계산 위로 배치
- CN참조 팝업: 360px + 새로고침 버튼 (Tab 2 편집 즉시 반영)
