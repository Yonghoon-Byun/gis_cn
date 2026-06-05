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

- Host: `geo-spatial-hub-prod.postgres.database.azure.com:6432`
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

## PostGIS 공간 처리 (db_manager.py)

DB에서 Clip+Intersection을 처리 → QGIS Processing 최소화. Geometry는 WKB(바이너리)로 전송.

- **`get_all_layers(wkt, srid, level)`** (권장): 단일 DB 연결 + 임시 테이블로 ①②③ 일괄 처리. clip 중복 제거로 ~40% 속도 향상.
- `get_soil_layer`/`get_land_cover_layer`/`get_soil_lc_intersection`: 개별 함수 (하위호환용)
- `_LC_DISSOLVE_COLS`가 level별 dissolve 컬럼 정의

## Intersection 후 컬럼 중복 처리

`native:intersection` 실행 후 동일 컬럼명에 숫자 접두어가 붙을 수 있다 (예: `hydro_type` → `2_hydro_type`). `spatial_ops._get_field_value()`에서 `endswith(f"_{target}")` 패턴으로 탐색하여 처리한다.

## UI 구조 (4탭)

탭 0·2는 `.ui` 파일에서, 탭 1은 `_setup_mapping_tab()`으로 동적 삽입(`insertTab(1,...)`), 탭 3은 `.ui` 파일의 원래 탭 2(인덱스 이동).

| 인덱스 | 탭명 | 위젯명/생성 | 주요 위젯 |
|--------|------|------------|---------|
| Tab 0 | 레이어 불러오기 | `tabCnCalc` (.ui) | rbFile/rbLayer, leFilePath, cmbNameField, rbL1/L2/L3, progressBar, txtLog, btnRun, btnClose, btnNextStep0 |
| Tab 1 | 토지이용 재분류 | `_setup_mapping_tab()` (동적) | tblMapping + 버튼, btnCnRefPopup(CN표 참조 팝업), btnMappingLoadLayer, btnMappingAddRow, btnMappingDeleteRow, btnMappingSave, btnMappingClear, btnNextStep1 |
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
- `AMC3_lu` = 79 (논/답 고정) | `trunc(23×AMC2_lu / (10+0.13×AMC2_lu))` (기타)

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

## 미구현 기능

- 토지피복도 커스텀 분류 (L1/L2/L3 혼합 분류 — A안 확정, 미구현)
- 삽도(지도 이미지) 자동 생성
- **개발 전/중/후 3단계 dialog UI**: 데이터모델(`StagedReport`)·렌더러(`render_staged_report`)는 완성, dialog 는 아직 단일단계 `render_hwpx`만 호출. 단계 선택·누적 UI, 단계별 SHP 입력, 적정성 사유 입력, meta 표지 누름틀 미구현.
- 표4-19 단계 컬럼 rowSpan 병합(현재 행마다 표시 — 미관용 선택)

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
