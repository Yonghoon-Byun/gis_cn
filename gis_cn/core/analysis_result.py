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
    """p.25 기준 분류표 한 행 (cn_value.xlsx 또는 Tab 2 편집표의 한 행)."""
    land_use: str
    a: Optional[int]
    b: Optional[int]
    c: Optional[int]
    d: Optional[int]


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
