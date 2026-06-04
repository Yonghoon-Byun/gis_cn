"""순수 Python HWPX(OWPML) 렌더러 — 한컴오피스/COM 불필요(lxml only).

검증된 PoC(개발 PC 한글 실개봉 확인, 2026-06-04)를 제품화한 모듈.
`AnalysisResult`(단계 1개분)를 빈 양식 hwpx 템플릿에 채워 `.hwpx`로 저장한다.
표의 데이터 행은 프로토타입 1행을 결과 행수만큼 deepcopy하여 동적 생성한다.

핵심 규칙(한글 크래시 회피 — PoC에서 규명):
  · 헤더 = 첫 데이터행(첫 누름틀 행) 이전 행들. 끝 빈 행은 데이터영역으로 간주해 교체.
  · 행 재구성 후 모든 <hp:cellAddr rowAddr> 를 위치순 0..rowCnt-1 로 재번호 + rowCnt 갱신.
  · 복제 행의 fieldBegin id 를 유니크 재배정(짝 fieldEnd beginIDRef 동기화).
  · 출하 전 자체검증: rowAddr 연속·범위, 그리드 완전 타일링, fieldBegin/End 짝.
  · zip 재패킹 시 mimetype 을 STORED(비압축)로 먼저 기록.

스키마 정합: `analysis_result.py` 의 res.* 필드 ID 및 `templates/v1.0` 템플릿과 동기화.
"""
from __future__ import annotations

import logging
import zipfile
from copy import deepcopy
from pathlib import Path
from typing import Optional

from .analysis_result import (
    AnalysisResult, WatershedBlock, LandUseRow, StagedReport, CnDeltaRow,
)

logger = logging.getLogger(__name__)

SECTION = "Contents/section0.xml"

# 원본 재해영향평가 보고서 쪽 여백 (HWPUNIT = 1/7200 inch)
ORIG_MARGIN = {"top": "4252", "bottom": "4252", "left": "4252", "right": "4252",
               "header": "3543", "footer": "3543", "gutter": "0"}

# 산정결과표(표 4-19) res.* 셀필드 (prefix 'res')
RES_PREFIX = "res"
META_FIELDS = (
    "meta.project_name", "meta.site_name", "meta.author",
    "meta.organization", "meta.analysis_date", "meta.development_stage",
)


class HwpxRenderError(RuntimeError):
    """HWPX 렌더링 실패 (lxml 미설치, 템플릿 오류, 자체검증 실패 포함)."""


def _import_lxml():
    try:
        from lxml import etree  # type: ignore
    except ImportError as e:
        raise HwpxRenderError(
            "lxml 이 설치되어 있지 않습니다. QGIS 번들 Python 에 lxml 이 필요합니다 "
            "(`pip install lxml`). 한글 보고서 출력에 한컴오피스는 불필요합니다."
        ) from e
    return etree


# ─────────────────────────────────────────────────────────────────────────────
# 숫자 포맷
# ─────────────────────────────────────────────────────────────────────────────

def _fmt_area(v: Optional[float]) -> str:
    return "" if v is None else f"{v:,.3f}"


def _fmt_cn(v) -> str:
    if v is None:
        return ""
    if isinstance(v, float):
        return f"{v:.2f}"
    return str(v)


# ─────────────────────────────────────────────────────────────────────────────
# 저수준 OWPML 헬퍼
# ─────────────────────────────────────────────────────────────────────────────

class _Doc:
    """로드된 hwpx 1개분 (zip 원본 + section0 lxml 트리)."""

    def __init__(self, etree, template: Path):
        self.etree = etree
        zin = zipfile.ZipFile(template)
        self.names = zin.namelist()
        self.raw = {n: zin.read(n) for n in self.names}
        zin.close()
        if SECTION not in self.raw:
            raise HwpxRenderError(f"템플릿에 {SECTION} 없음: {template}")
        self.root = etree.fromstring(self.raw[SECTION], etree.XMLParser(huge_tree=True))
        self.ns = self.root.nsmap.get("hp")
        if not self.ns:
            raise HwpxRenderError("템플릿 네임스페이스(hp) 누락 — OWPML 형식 아님")
        self._uid = 900_000_000

    def hp(self, tag: str) -> str:
        return f"{{{self.ns}}}{tag}"

    def fresh_id(self) -> str:
        self._uid += 1
        return str(self._uid)

    def tables(self):
        return list(self.root.iter(self.hp("tbl")))

    def find_table_with_field(self, prefix: str):
        """누름틀 name 이 `prefix.` 로 시작하는 첫 표."""
        for tbl in self.root.iter(self.hp("tbl")):
            if any((fb.get("name") or "").startswith(prefix + ".")
                   for fb in tbl.iter(self.hp("fieldBegin"))):
                return tbl
        return None

    def set_margins(self, margins: dict) -> int:
        n = 0
        for mg in self.root.iter(self.hp("margin")):
            if mg.getparent().tag == self.hp("pagePr"):
                for k, v in margins.items():
                    mg.set(k, v)
                n += 1
        return n

    def fill_field(self, scope, name: str, value: str) -> bool:
        """scope(표 또는 행) 안에서 CLICK_HERE 누름틀(name) 사이에 <hp:t>value</hp:t> 주입."""
        for fb in scope.iter(self.hp("fieldBegin")):
            if fb.get("name") != name:
                continue
            run = fb.getparent().getparent()      # ctrl -> run
            end_ctrl = None
            for ctrl in run.findall(self.hp("ctrl")):
                fe = ctrl.find(self.hp("fieldEnd"))
                if fe is not None and fe.get("beginIDRef") == fb.get("id"):
                    end_ctrl = ctrl
                    break
            for tt in run.findall(self.hp("t")):
                run.remove(tt)
            t = self.etree.SubElement(run, self.hp("t"))
            t.text = value
            if end_ctrl is not None:
                run.remove(t)
                end_ctrl.addprevious(t)
            return True
        return False

    def reassign_field_ids(self, tr) -> None:
        for fb in tr.iter(self.hp("fieldBegin")):
            old = fb.get("id")
            nid, nfid = self.fresh_id(), self.fresh_id()
            fb.set("id", nid)
            fb.set("fieldid", nfid)
            run = fb.getparent().getparent()
            for fe in run.iter(self.hp("fieldEnd")):
                if fe.get("beginIDRef") == old:
                    fe.set("beginIDRef", nid)
                    fe.set("fieldid", nfid)

    def serialize(self) -> None:
        self.raw[SECTION] = self.etree.tostring(
            self.root, xml_declaration=True, encoding="UTF-8", standalone=True)

    def save(self, out: Path) -> None:
        self.serialize()
        out = Path(out)
        if out.exists():
            out.unlink()
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
            zi = zipfile.ZipInfo("mimetype")
            zi.compress_type = zipfile.ZIP_STORED
            zout.writestr(zi, self.raw.get("mimetype", b"application/hwp+zip"))
            for n in self.names:
                if n != "mimetype":
                    zout.writestr(n, self.raw[n])


# ─────────────────────────────────────────────────────────────────────────────
# 동적 행 생성 (프로토타입 deepcopy × N)
# ─────────────────────────────────────────────────────────────────────────────

def _render_dynamic_table(doc: _Doc, table, field_prefix: str,
                          data_rows: "list[dict[str, str]]") -> None:
    """표의 데이터행을 data_rows(필드명→값 dict) 개수에 맞춰 동적 생성.

    헤더(첫 데이터행 이전)는 보존, 기존 데이터영역은 전부 교체. 각 행은 프로토타입을
    deepcopy → 필드 id 재배정 → 값 주입. 후처리로 전체 rowAddr 재번호 + rowCnt 갱신.
    """
    hp = doc.hp
    trs = table.findall(hp("tr"))

    def is_data(tr):
        return any((fb.get("name") or "").startswith(field_prefix + ".")
                   for fb in tr.iter(hp("fieldBegin")))

    first_data = next((i for i, tr in enumerate(trs) if is_data(tr)), None)
    if first_data is None:
        raise HwpxRenderError(f"템플릿 표에 '{field_prefix}.*' 프로토타입 행 없음")
    header_rows = trs[:first_data]
    prototype = deepcopy(trs[first_data])

    for tr in trs[first_data:]:      # 기존 데이터영역(끝 빈 행 포함) 제거
        table.remove(tr)

    prev = header_rows[-1] if header_rows else None
    for rowdata in data_rows:
        nr = deepcopy(prototype)
        doc.reassign_field_ids(nr)
        for name, value in rowdata.items():
            doc.fill_field(nr, name, value)
        if prev is not None:
            prev.addnext(nr)
        else:
            table.append(nr)
        prev = nr

    # 전체 행 rowAddr 위치순 재번호 + rowCnt
    all_rows = table.findall(hp("tr"))
    for pos, tr in enumerate(all_rows):
        for ca in tr.iter(hp("cellAddr")):
            ca.set("rowAddr", str(pos))
    table.set("rowCnt", str(len(all_rows)))


def _res_row(prefix: str, stage: str, ws: str, row: LandUseRow) -> "dict[str, str]":
    p = lambda local: f"{prefix}.{local}"
    return {
        p("stage"): stage, p("ws"): ws, p("lu"): row.land_use,
        p("a_area"): _fmt_area(row.a_area), p("a_cn"): _fmt_cn(row.a_cn),
        p("b_area"): _fmt_area(row.b_area), p("b_cn"): _fmt_cn(row.b_cn),
        p("c_area"): _fmt_area(row.c_area), p("c_cn"): _fmt_cn(row.c_cn),
        p("d_area"): _fmt_area(row.d_area), p("d_cn"): _fmt_cn(row.d_cn),
        p("total_area"): _fmt_area(row.total_area),
        p("amc2_cn"): _fmt_cn(row.amc2_cn), p("amc3_cn"): _fmt_cn(row.amc3_cn),
    }


def _res_summary_row(prefix: str, stage: str, block: WatershedBlock) -> "dict[str, str]":
    p = lambda local: f"{prefix}.{local}"
    return {
        p("stage"): stage, p("ws"): block.name, p("lu"): "합계",
        p("a_area"): _fmt_area(block.total_a or None), p("a_cn"): "",
        p("b_area"): _fmt_area(block.total_b or None), p("b_cn"): "",
        p("c_area"): _fmt_area(block.total_c or None), p("c_cn"): "",
        p("d_area"): _fmt_area(block.total_d or None), p("d_cn"): "",
        p("total_area"): _fmt_area(block.total_area),
        p("amc2_cn"): _fmt_cn(block.amc2_cn), p("amc3_cn"): _fmt_cn(block.amc3_cn),
    }


def _detail_rows(blocks: "list[WatershedBlock]", prefix: str, stage: str = "") -> "list[dict[str, str]]":
    """소유역 블록들 → res 표 데이터행 dict 리스트 (블록별 토지이용행 + 합계행).

    stage 가 주어지면 각 행의 단계(res.stage) 셀에 채운다(다단계 보고서의 단계 구분).
    템플릿에 res.stage 누름틀이 없으면 fill_field 가 조용히 무시(단일단계 호환).
    """
    out: list[dict] = []
    for block in blocks:
        for row in block.rows:
            out.append(_res_row(prefix, stage, block.name, row))
        out.append(_res_summary_row(prefix, stage, block))
    return out


# ─────────────────────────────────────────────────────────────────────────────
# 자체검증
# ─────────────────────────────────────────────────────────────────────────────

def _validate(doc: _Doc) -> None:
    hp = doc.hp
    problems: list[str] = []
    for ti, tbl in enumerate(doc.tables()):
        rc, cc = int(tbl.get("rowCnt")), int(tbl.get("colCnt"))
        grid = {}
        over = False
        for tr in tbl.findall(hp("tr")):
            for tc in tr.findall(hp("tc")):
                ca, sp = tc.find(hp("cellAddr")), tc.find(hp("cellSpan"))
                r, c = int(ca.get("rowAddr")), int(ca.get("colAddr"))
                rs = int(sp.get("rowSpan")) if sp is not None else 1
                cs = int(sp.get("colSpan")) if sp is not None else 1
                if r >= rc or c >= cc:
                    problems.append(f"표{ti}: cellAddr({r},{c}) 범위초과 {rc}x{cc}")
                    over = True
                for dr in range(rs):
                    for dc in range(cs):
                        key = (r + dr, c + dc)
                        if key in grid:
                            problems.append(f"표{ti}: 셀 중복 {key}")
                        grid[key] = 1
        if not over:
            missing = [(r, c) for r in range(rc) for c in range(cc) if (r, c) not in grid]
            if missing:
                problems.append(f"표{ti}: 그리드 미충족 {len(missing)}칸 예{missing[:3]}")
    begins = {}
    for fb in doc.root.iter(hp("fieldBegin")):
        begins[fb.get("id")] = begins.get(fb.get("id"), 0) + 1
    for fid, cnt in begins.items():
        if cnt > 1:
            problems.append(f"fieldBegin id 중복: {fid} x{cnt}")
    for fe in doc.root.iter(hp("fieldEnd")):
        if fe.get("beginIDRef") not in begins:
            problems.append(f"fieldEnd 고아 참조: {fe.get('beginIDRef')}")
    if problems:
        raise HwpxRenderError("자체검증 실패(한글 크래시 위험): " + "; ".join(problems[:8]))


# ─────────────────────────────────────────────────────────────────────────────
# 메인 엔트리
# ─────────────────────────────────────────────────────────────────────────────

def render_hwpx(result: AnalysisResult, template, out, *, apply_orig_margin: bool = True) -> Path:
    """`AnalysisResult`(단계 1개분) → HWPX. 한컴오피스/COM 불필요.

    현재 구현 범위: 쪽 여백 + 메타 누름틀 + 산정결과표(표 4-19) 동적 행.
    (기준표 표4-18·다단계 표4-19·비교표 표4-17 은 클린 템플릿 완성 후 확장.)

    Raises:
        HwpxRenderError: lxml 미설치, 템플릿 오류, 자체검증 실패.
    """
    etree = _import_lxml()
    template, out = Path(template), Path(out)
    if not template.exists():
        raise HwpxRenderError(f"템플릿을 찾을 수 없습니다: {template}")

    doc = _Doc(etree, template)

    if apply_orig_margin:
        doc.set_margins(ORIG_MARGIN)

    # 메타 누름틀 (템플릿에 있을 때만; graceful)
    m = result.meta
    meta_vals = {
        "meta.project_name": m.project_name, "meta.site_name": m.site_name,
        "meta.author": m.author, "meta.organization": m.organization,
        "meta.analysis_date": m.analysis_date.isoformat(),
        "meta.development_stage": m.development_stage,
    }
    filled_meta = sum(int(doc.fill_field(doc.root, k, v)) for k, v in meta_vals.items())

    # 산정결과표(표 4-19)
    res_tbl = doc.find_table_with_field(RES_PREFIX)
    if res_tbl is None:
        raise HwpxRenderError(f"템플릿에 '{RES_PREFIX}.*' 산정결과표 없음")
    data_rows = _detail_rows(result.detail_blocks, RES_PREFIX, stage=result.meta.development_stage)
    if not data_rows:
        raise HwpxRenderError("산정결과 데이터가 비어 있습니다 (detail_blocks 없음)")
    _render_dynamic_table(doc, res_tbl, RES_PREFIX, data_rows)

    _validate(doc)
    doc.save(out)
    logger.info("HWPX 저장: %s (res %d행, meta %d필드)", out, len(data_rows), filled_meta)
    return out


# ─────────────────────────────────────────────────────────────────────────────
# 다단계 비교 보고서 (StagedReport) — 표4-17 비교표 + 표4-19 단계 결과
# ─────────────────────────────────────────────────────────────────────────────

DELTA_PREFIX = "delta"


def _fmt_increment(v) -> str:
    """표 4-17 증감 표기: 증가 '증) X.XX' / 감소 '감) X.XX' / 변화없음·없음 '-'."""
    if v is None:
        return "-"
    if v > 0:
        return f"증) {v:.2f}"
    if v < 0:
        return f"감) {abs(v):.2f}"
    return "-"


def _delta_rows_data(rows: "list[CnDeltaRow]") -> "list[dict[str, str]]":
    out: list[dict] = []
    for r in rows:
        out.append({
            "delta.ws": r.watershed,
            "delta.cn_pre": _fmt_cn(r.cn_pre),
            "delta.cn_during": _fmt_cn(r.cn_during),
            "delta.delta_during": _fmt_increment(r.delta_during),
            "delta.reason_during": r.reason_during or "",
            "delta.cn_post": _fmt_cn(r.cn_post),
            "delta.delta_post": _fmt_increment(r.delta_post),
            "delta.reason_post": r.reason_post or "",
        })
    return out


def render_staged_report(report: StagedReport, template, out, *, apply_orig_margin: bool = True) -> Path:
    """`StagedReport`(개발 전/중/후) → HWPX. 한컴/COM 불필요.

    채움: 메타 + 표4-17 비교표(cn_delta_rows) + 표4-19 산정결과표.
    표4-18 기준표는 정적(템플릿 내장)이라 손대지 않는다.
    (현재 표4-19는 전체 단계 블록을 한 표에 누적 — 단계별 페이지 분리 B-1은 후속 F3b.)
    """
    etree = _import_lxml()
    template, out = Path(template), Path(out)
    if not template.exists():
        raise HwpxRenderError(f"템플릿을 찾을 수 없습니다: {template}")

    doc = _Doc(etree, template)
    if apply_orig_margin:
        doc.set_margins(ORIG_MARGIN)

    m = report.meta
    meta_vals = {
        "meta.project_name": m.project_name, "meta.site_name": m.site_name,
        "meta.author": m.author, "meta.organization": m.organization,
        "meta.analysis_date": m.analysis_date.isoformat(),
        "meta.development_stage": " / ".join(s for s, _ in report.ordered_stages()) or m.development_stage,
    }
    for k, v in meta_vals.items():
        doc.fill_field(doc.root, k, v)

    # 표 4-17 비교표 (delta.*)
    delta_tbl = doc.find_table_with_field(DELTA_PREFIX)
    if delta_tbl is not None and report.cn_delta_rows:
        _render_dynamic_table(doc, delta_tbl, DELTA_PREFIX, _delta_rows_data(report.cn_delta_rows))

    # 표 4-19 산정결과표 (res.*) — 단계별 행에 res.stage 로 단계 표시(전체 단계 누적).
    res_tbl = doc.find_table_with_field(RES_PREFIX)
    res_data: list[dict] = []
    for stage_name, ares in report.ordered_stages():
        res_data.extend(_detail_rows(ares.detail_blocks, RES_PREFIX, stage=stage_name))
    if res_tbl is not None and res_data:
        _render_dynamic_table(doc, res_tbl, RES_PREFIX, res_data)

    _validate(doc)
    doc.save(out)
    logger.info("HWPX(다단계) 저장: %s (delta %d, res %d행)",
                out, len(report.cn_delta_rows), len(res_data))
    return out
