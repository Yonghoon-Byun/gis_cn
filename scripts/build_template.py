# -*- coding: utf-8 -*-
"""F2: 클린 HWPX 템플릿 빌더 (한컴오피스/COM 불필요, lxml only).

현 cn_report.hwpx의 깨진 부분(852cb76 '알려진 한계')을 정리한다:
  · 빈 "(계속)" res 쓰레기 표 5개 제거 (조부모 <hp:p> 통째)
  · res 산정결과표를 헤더 + 프로토타입 데이터 1행으로 축소(런타임에 동적 복제)
  · 원본 재해영향평가 보고서 쪽 여백(사방 15mm / 머리·꼬리 12.5mm) 적용
  · 전체 행 rowAddr 위치순 재번호 + 검증(한글 크래시 회피)

후속(F2b): ref 8열 국가표준 기준표(표4-18)·meta 표지·표4-17 증감 비교표 추가.

사용: python scripts/build_template.py [출력경로]   (생략 시 템플릿 덮어쓰기)
"""
from __future__ import annotations

import sys
import zipfile
from copy import deepcopy
from pathlib import Path

from lxml import etree

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

ROOT = Path(__file__).resolve().parent.parent
TEMPLATE = ROOT / "gis_cn" / "templates" / "v1.0" / "cn_report.hwpx"
SAMPLE = ROOT / "gis_cn" / "templates" / "v1.0" / "cn_report_sample.hwpx"
SECTION = "Contents/section0.xml"
HEADER = "Contents/header.xml"
ORIG_MARGIN = {"top": "4252", "bottom": "4252", "left": "4252", "right": "4252",
               "header": "3543", "footer": "3543", "gutter": "0"}
RES_PREFIX = "res"


def _hh(root):
    ns = None
    for k, v in root.nsmap.items():
        if k == "hh":
            ns = v
    return ns, (lambda t: f"{{{ns}}}{t}")


def add_standard_ref_table(section_root, raw, hp) -> dict:
    """표 4-18 국가표준 CN 기준표(8열, 정적)를 샘플에서 템플릿으로 이식.

    누름틀 없는 고정표(샘플에 이미 표준값 포함). 누락 charPr(14,15)만 샘플 header에서
    복사(참조 폰트/테두리는 템플릿에 이미 존재). 기준표 단락을 res 표 단락 앞에 삽입.
    OWPML 네임스페이스(hp/hh)는 샘플·템플릿 동일(한컴 2011 표준)하여 deepcopy 호환.
    """
    info = {"charpr_added": 0, "ref_rows": 0}
    # idempotency: 이미 8열 기준표가 있으면 중복 삽입 방지
    for t in section_root.iter(hp("tbl")):
        if t.get("colCnt") == "8" and "AMC-II" in "".join(
                x.text or "" for x in (t.find(hp("caption")) if t.find(hp("caption")) is not None else t).iter(hp("t"))):
            info["skipped"] = True
            return info
    smp_sec = etree.fromstring(zipfile.ZipFile(SAMPLE).read(SECTION), etree.XMLParser(huge_tree=True))
    sns = smp_sec.nsmap.get("hp")
    shp = lambda t: f"{{{sns}}}{t}"
    ref_tbl = next(
        t for t in smp_sec.iter(shp("tbl"))
        if "AMC-II" in "".join(x.text or "" for x in
                              (t.find(shp("caption")) if t.find(shp("caption")) is not None else t).iter(shp("t")))
        and t.get("colCnt") == "8")
    ref_p = ref_tbl.getparent().getparent()      # tbl -> run -> p
    info["ref_rows"] = len(ref_tbl.findall(shp("tr")))

    # header.xml: charPr 14,15 복사 (참조 폰트/테두리는 기존재 — 사전 실측 확인)
    hdr = etree.fromstring(raw[HEADER], etree.XMLParser(huge_tree=True))
    _, hh = _hh(hdr)
    char_container = next(cp for cp in hdr.iter(hh("charPr"))).getparent()
    existing = {cp.get("id") for cp in char_container.findall(hh("charPr"))}
    smp_hdr = etree.fromstring(zipfile.ZipFile(SAMPLE).read(HEADER), etree.XMLParser(huge_tree=True))
    _, shh = _hh(smp_hdr)
    for cid in ("14", "15"):
        if cid in existing:
            continue
        src = next(cp for cp in smp_hdr.iter(shh("charPr")) if cp.get("id") == cid)
        char_container.append(deepcopy(src))
        info["charpr_added"] += 1
    if char_container.get("itemCnt") is not None:
        char_container.set("itemCnt", str(len(char_container.findall(hh("charPr")))))
    raw[HEADER] = etree.tostring(hdr, xml_declaration=True, encoding="UTF-8", standalone=True)

    # 기준표 단락을 res 표 단락 앞에 삽입 (보고서 순서: 표4-18 → 표4-19)
    res_tbl = next(t for t in section_root.iter(hp("tbl"))
                   if any((fb.get("name") or "").startswith(RES_PREFIX + ".")
                          for fb in t.iter(hp("fieldBegin"))))
    res_tbl.getparent().getparent().addprevious(deepcopy(ref_p))
    return info


def _ns(root):
    ns = root.nsmap.get("hp")
    return ns, (lambda t: f"{{{ns}}}{t}")


def _res_field_count(tbl, hp) -> int:
    return sum(1 for fb in tbl.iter(hp("fieldBegin"))
              if (fb.get("name") or "").startswith(RES_PREFIX + "."))


# 표 4-17 비교표 열(colAddr) → delta.* 누름틀 매핑
DELTA_COL_FIELDS = {
    0: "delta.ws", 1: "delta.cn_pre", 2: "delta.cn_during", 3: "delta.delta_during",
    4: "delta.reason_during", 5: "delta.cn_post", 6: "delta.delta_post", 7: "delta.reason_post",
}


def _field_prototype(section_root, hp):
    """템플릿 res 셀에서 CLICK_HERE 누름틀(fieldBegin/fieldEnd ctrl) 프로토타입 추출."""
    fb = next(fb for fb in section_root.iter(hp("fieldBegin"))
              if (fb.get("name") or "").startswith("res."))
    begin_ctrl = fb.getparent()              # <hp:ctrl> wrapping fieldBegin
    run = begin_ctrl.getparent()
    end_ctrl = next(c for c in run.findall(hp("ctrl")) if c.find(hp("fieldEnd")) is not None)
    return deepcopy(begin_ctrl), deepcopy(end_ctrl)


_FIELD_UID = [700_000_000]


def _fresh_field_id() -> str:
    _FIELD_UID[0] += 1
    return str(_FIELD_UID[0])


def _inject_field(cell, begin_proto, end_proto, name, hp):
    """셀의 첫 run 을 CLICK_HERE 누름틀(name)로 교체(기존 텍스트 제거, 빈 값 셀필드)."""
    begin, end = deepcopy(begin_proto), deepcopy(end_proto)
    nid, nfid = _fresh_field_id(), _fresh_field_id()
    fb = begin.find(hp("fieldBegin")); fb.set("name", name); fb.set("id", nid); fb.set("fieldid", nfid)
    fe = end.find(hp("fieldEnd")); fe.set("beginIDRef", nid); fe.set("fieldid", nfid)
    for t in cell.iter(hp("t")):             # 셀 내 모든 텍스트 비움
        t.text = ""
    run = cell.find(".//" + hp("run"))
    for ch in list(run):
        if etree.QName(ch).localname in ("ctrl", "t"):
            run.remove(ch)
    run.append(begin); run.append(end)
    etree.SubElement(run, hp("t")).text = ""


def add_delta_table(section_root, raw, hp) -> dict:
    """표 4-17 개발단계별 소유역 CN 변화 비교표를 샘플에서 이식 + delta.* 누름틀 주입.

    헤더(2행) + 프로토타입 데이터 1행으로 축소(런타임 동적 복제). 보고서 순서상
    맨 앞(표4-17 → 표4-18 → 표4-19)에 삽입. 스타일(charPr/borderFill) 템플릿 호환 실측.
    """
    info = {"delta_fields": 0}
    for t in section_root.iter(hp("tbl")):
        if any((fb.get("name") or "").startswith("delta.") for fb in t.iter(hp("fieldBegin"))):
            info["skipped"] = True
            return info
    smp = etree.fromstring(zipfile.ZipFile(SAMPLE).read(SECTION), etree.XMLParser(huge_tree=True))
    sns = smp.nsmap.get("hp"); shp = lambda t: f"{{{sns}}}{t}"
    full = lambda el: "".join(x.text or "" for x in el.iter(shp("t")))
    src_tbl = next(t for t in smp.iter(shp("tbl")) if t.find(shp("caption")) is not None
                   and "유출곡선지수" in full(t.find(shp("caption")))
                   and "변화" in full(t.find(shp("caption"))))
    new_p = deepcopy(src_tbl.getparent().getparent())   # tbl -> run -> p
    next(section_root.iter(hp("tbl"))).getparent().getparent().addprevious(new_p)

    new_tbl = next(new_p.iter(hp("tbl")))
    trs = new_tbl.findall(hp("tr"))
    header_n = 2                              # r0,r1 = 복합 헤더, r2 = 프로토타입 데이터
    prototype = trs[header_n]
    for tr in trs[header_n + 1:]:
        new_tbl.remove(tr)

    begin_proto, end_proto = _field_prototype(section_root, hp)
    for tc in prototype.findall(hp("tc")):
        col = int(tc.find(hp("cellAddr")).get("colAddr"))
        name = DELTA_COL_FIELDS.get(col)
        if name:
            _inject_field(tc, begin_proto, end_proto, name, hp)
            info["delta_fields"] += 1

    all_rows = new_tbl.findall(hp("tr"))
    for pos, tr in enumerate(all_rows):
        for ca in tr.iter(hp("cellAddr")):
            ca.set("rowAddr", str(pos))
    new_tbl.set("rowCnt", str(len(all_rows)))
    return info


def clean_template(src: Path, out: Path) -> dict:
    zin = zipfile.ZipFile(src)
    names = zin.namelist()
    raw = {n: zin.read(n) for n in names}
    zin.close()
    root = etree.fromstring(raw[SECTION], etree.XMLParser(huge_tree=True))
    ns, hp = _ns(root)
    stats = {"junk_removed": 0, "res_rows_before": 0, "res_rows_after": 0, "margins": 0}

    # 1) 쓰레기 res 표(필드 0) 제거 — 조부모 <hp:p> 통째
    tables = list(root.iter(hp("tbl")))
    res_tables = [t for t in tables
                  if "산정결과" in "".join(x.text or "" for x in
                                       (t.find(hp("caption")) if t.find(hp("caption")) is not None else t).iter(hp("t")))]
    real_res = next((t for t in res_tables if _res_field_count(t, hp) > 0), None)
    if real_res is None:
        raise RuntimeError("res.* 필드를 가진 산정결과표를 찾지 못함")
    for t in res_tables:
        if t is real_res:
            continue
        p = t.getparent().getparent()      # tbl -> run -> p
        p.getparent().remove(p)
        stats["junk_removed"] += 1

    # 2) real res 표: 헤더 + 프로토타입 1행으로 축소
    trs = real_res.findall(hp("tr"))
    stats["res_rows_before"] = len(trs)

    def is_data(tr):
        return any((fb.get("name") or "").startswith(RES_PREFIX + ".")
                   for fb in tr.iter(hp("fieldBegin")))

    first_data = next(i for i, tr in enumerate(trs) if is_data(tr))
    prototype = deepcopy(trs[first_data])
    for tr in trs[first_data:]:            # 데이터영역(끝 빈 행 포함) 제거
        real_res.remove(tr)
    if first_data > 0:
        trs[first_data - 1].addnext(prototype)   # 마지막 헤더 행 뒤에 프로토타입 삽입
    else:
        real_res.append(prototype)

    # 3) rowAddr 위치순 재번호 + rowCnt
    all_rows = real_res.findall(hp("tr"))
    for pos, tr in enumerate(all_rows):
        for ca in tr.iter(hp("cellAddr")):
            ca.set("rowAddr", str(pos))
    real_res.set("rowCnt", str(len(all_rows)))
    stats["res_rows_after"] = len(all_rows)

    # 3.5) res 표 c0(단계) 셀에 res.stage 누름틀 주입 (다단계 단계 표시)
    proto = next((tr for tr in real_res.findall(hp("tr"))
                  if any((fb.get("name") or "").startswith(RES_PREFIX + ".")
                         for fb in tr.iter(hp("fieldBegin")))), None)
    if proto is not None:
        c0 = next((tc for tc in proto.findall(hp("tc"))
                   if tc.find(hp("cellAddr")).get("colAddr") == "0"), None)
        if c0 is not None and not any(fb.get("name") == "res.stage" for fb in c0.iter(hp("fieldBegin"))):
            bp, ep = _field_prototype(real_res, hp)
            _inject_field(c0, bp, ep, "res.stage", hp)
            stats["res_stage_field"] = True

    # 3.6) 깨진 ref.* 누름틀(표 밖 본문 잔존) 제거 — 기준표는 정적 표4-18로 대체됨
    removed = 0
    for fb in list(root.iter(hp("fieldBegin"))):
        if not (fb.get("name") or "").startswith("ref."):
            continue
        ctrl_b = fb.getparent()
        run = ctrl_b.getparent()
        for ctrl in list(run.findall(hp("ctrl"))):
            fe = ctrl.find(hp("fieldEnd"))
            if fe is not None and fe.get("beginIDRef") == fb.get("id"):
                run.remove(ctrl)
        run.remove(ctrl_b)
        removed += 1
    stats["orphan_ref_removed"] = removed

    # 4) 원본 여백
    for mg in root.iter(hp("margin")):
        if mg.getparent().tag == hp("pagePr"):
            for k, v in ORIG_MARGIN.items():
                mg.set(k, v)
            stats["margins"] += 1

    # 4.5) 표 4-18 국가표준 CN 기준표(정적 8열) 이식
    stats.update(add_standard_ref_table(root, raw, hp))

    # 4.6) 표 4-17 개발단계별 CN 변화 비교표(delta.* 누름틀) 이식 — 맨 앞
    delta_info = add_delta_table(root, raw, hp)
    stats["delta_fields"] = delta_info.get("delta_fields", 0)
    stats["delta_skipped"] = delta_info.get("skipped", False)

    # 5) 검증 — 모든 표 rowAddr 범위·연속, 그리드 타일링
    _validate(root, hp)

    # 6) 저장
    raw[SECTION] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    out = Path(out)
    if out.exists():
        out.unlink()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        zi = zipfile.ZipInfo("mimetype")
        zi.compress_type = zipfile.ZIP_STORED
        zout.writestr(zi, raw.get("mimetype", b"application/hwp+zip"))
        for n in names:
            if n != "mimetype":
                zout.writestr(n, raw[n])
    return stats


def _validate(root, hp) -> None:
    problems = []
    for ti, tbl in enumerate(root.iter(hp("tbl"))):
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
                    problems.append(f"표{ti} cellAddr({r},{c}) 범위초과 {rc}x{cc}"); over = True
                for dr in range(rs):
                    for dc in range(cs):
                        k = (r + dr, c + dc)
                        if k in grid:
                            problems.append(f"표{ti} 셀 중복 {k}")
                        grid[k] = 1
        if not over:
            miss = [(r, c) for r in range(rc) for c in range(cc) if (r, c) not in grid]
            if miss:
                problems.append(f"표{ti} 그리드 미충족 {len(miss)}칸 예{miss[:3]}")
    if problems:
        raise RuntimeError("검증 실패(한글 크래시 위험): " + "; ".join(problems[:8]))


def main() -> int:
    out = Path(sys.argv[1]) if len(sys.argv) > 1 else TEMPLATE
    stats = clean_template(TEMPLATE, out)
    print(f"클린 템플릿 생성: {out}")
    print(f"  쓰레기 표 제거: {stats['junk_removed']}개")
    print(f"  res 행: {stats['res_rows_before']} → {stats['res_rows_after']} (헤더 + 프로토타입 1)")
    print(f"  여백 적용: pagePr margin {stats['margins']}개 (사방 15mm)")
    if stats.get("skipped"):
        print("  표4-18 기준표: 이미 존재(건너뜀)")
    else:
        print(f"  표4-18 기준표 이식: {stats.get('ref_rows', 0)}행, charPr +{stats.get('charpr_added', 0)}")
    if stats.get("delta_skipped"):
        print("  표4-17 비교표: 이미 존재(건너뜀)")
    else:
        print(f"  표4-17 비교표 이식: delta.* 누름틀 {stats.get('delta_fields', 0)}개")
    print("  검증 통과 ✓")
    return 0


if __name__ == "__main__":
    sys.exit(main())
