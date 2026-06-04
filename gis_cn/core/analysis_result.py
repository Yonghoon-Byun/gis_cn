"""CN 분석 결과의 Single Source of Truth.

Excel·HWP 양쪽 렌더러는 이 dataclass만 입력으로 받아 출력합니다.
기존 `result_calculator.calculate_results()`가 반환하는 list[dict] 구조는
`build_analysis_result()`를 통해 이 dataclass로 변환됩니다.

스키마 버전이 변경되면 `templates/vX.Y/`도 맞춰 갱신해야 합니다.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from pathlib import Path
from typing import Optional

SCHEMA_VERSION = "1.0"


@dataclass
class CnReferenceRow:
    """표 4-18 국가표준 CN 기준표 한 행 (8열: 대분류·중분류·코드·세분류 + A/B/C/D).

    국가표준 고정표(재해영향평가 실무지침, 행안부 고시)이며 `data/cn_reference_std.json`에서
    로드한다. land_use/a/b/c/d 는 기존 5열 호환을 위해 앞에 두고, l1/l2/code(중분류)/
    l3_code(세분류) 분류 계층을 뒤에 추가(기본값 "")하여 기존 위치인자 호출과 호환.
    """
    land_use: str            # 세분류명(보고서 표시)
    a: Optional[int]
    b: Optional[int]
    c: Optional[int]
    d: Optional[int]
    l1: str = ""             # 대분류명
    l2: str = ""             # 중분류명
    code: str = ""           # 중분류 코드(110 등, 보고서 표시용)
    l3_code: str = ""        # 세분류 코드(111 등, land_cover.l3_code 연결용)
    l3_name: str = ""        # 세분류 표준명(docs 분류체계 기준)

    @classmethod
    def from_std(cls, d: dict) -> "CnReferenceRow":
        """cn_reference_std.json 한 항목(dict) → CnReferenceRow."""
        return cls(
            land_use=d.get("lu", ""), a=d.get("a"), b=d.get("b"), c=d.get("c"), d=d.get("d"),
            l1=d.get("l1", ""), l2=d.get("l2", ""), code=d.get("code", ""),
            l3_code=d.get("l3_code", ""), l3_name=d.get("l3_name", ""),
        )


@dataclass
class LandUseRow:
    """p.26~ 산정결과표의 한 행 (토지이용 × 토양군 A/B/C/D)."""
    land_use: str
    a_area: Optional[float]; a_cn: Optional[int]
    b_area: Optional[float]; b_cn: Optional[int]
    c_area: Optional[float]; c_cn: Optional[int]
    d_area: Optional[float]; d_cn: Optional[int]
    total_area: Optional[float]
    amc2_cn: Optional[float]
    amc3_cn: Optional[int]


@dataclass
class WatershedBlock:
    """소유역(또는 유역합성) 1개의 산정결과 블록."""
    name: str
    is_composite: bool = False
    rows: list[LandUseRow] = field(default_factory=list)
    total_a: float = 0.0
    total_b: float = 0.0
    total_c: float = 0.0
    total_d: float = 0.0
    total_area: float = 0.0
    amc2_cn: float = 0.0
    amc3_cn: float = 0.0
    member_names: list[str] = field(default_factory=list)


@dataclass
class WatershedSummary:
    """유역별 CN정리 요약 한 행."""
    name: str
    total_area: float
    amc2_cn: float
    amc3_cn: float
    is_composite: bool = False


@dataclass
class CnDeltaRow:
    """표 4-17 한 행: 소유역별 개발 전/중/후 CN 변화 + 증감 + 적정성 검토 사유.

    증감은 AMC-Ⅲ CN 기준(②-①, ③-①, float 차). reason_* 는 사용자 입력(또는 LLM 초안).
    단계 간 소유역은 동일명으로 정렬(1차 정책) — 이름이 같은 행끼리 전/중/후를 가로 결합.
    """
    watershed: str                          # 소유역명(단계 간 정렬 키)
    cn_pre: Optional[float] = None          # 개발 전 ① (AMC-Ⅲ)
    cn_during: Optional[float] = None       # 개발 중 ②
    delta_during: Optional[float] = None    # ②-①
    reason_during: str = ""                 # 개발 중 적정성 검토 사유
    cn_post: Optional[float] = None         # 개발 후 ③
    delta_post: Optional[float] = None      # ③-①
    reason_post: str = ""                   # 개발 후 적정성 검토 사유


@dataclass
class MapImage:
    """HWP 본문에 삽입할 지도 이미지 (Excel은 사용 안 함)."""
    path: Path
    caption: str
    bookmark_id: str


@dataclass
class NullRow:
    """CN값 매칭 실패 피처."""
    watershed: str
    land_use: str
    hydro_type: str


@dataclass
class ProjectMeta:
    project_name: str = ""
    site_name: str = ""
    author: str = ""
    organization: str = "도화엔지니어링"
    analysis_date: date = field(default_factory=date.today)
    srid: int = 5186
    data_source: str = "db"
    land_cover_level: str = "l3"
    development_stage: str = "현 상태"


@dataclass
class AnalysisResult:
    """렌더러(Excel/HWP)가 공유하는 Single Source of Truth."""
    schema_version: str = SCHEMA_VERSION
    meta: ProjectMeta = field(default_factory=ProjectMeta)

    cn_reference: list[CnReferenceRow] = field(default_factory=list)

    detail_blocks: list[WatershedBlock] = field(default_factory=list)
    summary_rows: list[WatershedSummary] = field(default_factory=list)

    composite_detail: list[WatershedBlock] = field(default_factory=list)
    composite_summary: list[WatershedSummary] = field(default_factory=list)

    map_images: list[MapImage] = field(default_factory=list)
    null_cn_rows: list[NullRow] = field(default_factory=list)
    notes: str = ""


# ─────────────────────────────────────────────────────────────────────────────
# 다단계(개발 전/중/후) 비교 — 옵션 C: 단계는 컨테이너 레벨, 계산엔진은 단계와 직교
# ─────────────────────────────────────────────────────────────────────────────

STAGE_PRE = "개발 전"
STAGE_DURING = "개발 중"
STAGE_POST = "개발 후"
DEFAULT_STAGE_ORDER = [STAGE_PRE, STAGE_DURING, STAGE_POST]


@dataclass
class StagedReport:
    """개발 전/중/후 다단계 비교 보고서 컨테이너.

    각 단계는 독립 `AnalysisResult`(단계 1개분)이며, CN 계산엔진은 단계와 직교한다
    (같은 계산을 단계별 입력으로 N회 호출). 단계 결합·증감 비교(표 4-17)는
    `cn_delta_rows` 에 조립한다. 단일단계 사용 시 `stages` 에 1개만 담으면 기존 경로와 동일.
    """
    meta: ProjectMeta = field(default_factory=ProjectMeta)
    stages: dict[str, AnalysisResult] = field(default_factory=dict)    # 단계키 → 단계 결과
    stage_order: list[str] = field(default_factory=list)              # 출력 순서(채워진 단계만)
    cn_delta_rows: list[CnDeltaRow] = field(default_factory=list)     # 표 4-17 (소유역별 단계간 증감)
    cn_reference: list[CnReferenceRow] = field(default_factory=list)  # 표 4-18 (단계 무관 국가표준표)

    def ordered_stages(self) -> list[tuple[str, "AnalysisResult"]]:
        """채워진 단계를 출력 순서대로 (단계명, 결과) 튜플 리스트로 반환."""
        order = self.stage_order or [s for s in DEFAULT_STAGE_ORDER if s in self.stages]
        return [(s, self.stages[s]) for s in order if s in self.stages]
