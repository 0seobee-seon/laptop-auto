#!/usr/bin/env python3
"""
update_kaechwal.py
──────────────────
HWPX 주간업무 보고서에서 오늘 날짜(또는 지정 날짜) 기준으로
개찰일이 도래한 '진행중' 항목을 '개찰' 섹션으로 자동 이동합니다.

Usage:
    python update_kaechwal.py --input 입력.hwpx --output 출력.hwpx [--date YYYY-MM-DD]

Output (stdout, JSON):
    {
        "success": true,
        "reference_date": "2026-05-21",
        "moved_items": [{"num":"4","name":"의정부법조타운","kaechwal_date":"5/21"}],
        "remaining_items": [{"num":"7","name":"쌍문1동","kaechwal_date":"5/27"}],
        "kaechwal_count": 7,
        "jinhang_count": 1,
        "output_path": "출력.hwpx"
    }
"""

import argparse
import json
import re
import sys
import zipfile
from datetime import date, datetime


# ────────────────────────────────────────────
# CLI
# ────────────────────────────────────────────

def parse_args():
    p = argparse.ArgumentParser(description="HWPX 개찰 현황 업데이트")
    p.add_argument("--input",  required=True, help="입력 HWPX 파일 경로")
    p.add_argument("--output", required=True, help="출력 HWPX 파일 경로")
    p.add_argument("--date",   help="기준 날짜 YYYY-MM-DD (기본: 오늘)")
    return p.parse_args()


# ────────────────────────────────────────────
# 날짜 파싱
# ────────────────────────────────────────────

def parse_korean_date(text: str, ref_year: int):
    """
    '5/21'  → date(ref_year, 5, 21)
    '추후'  → None
    ''      → None
    """
    text = text.strip()
    if not text or text in ("추후", "-", "미정"):
        return None
    m = re.match(r"^(\d{1,2})/(\d{1,2})$", text)
    if m:
        try:
            return date(ref_year, int(m.group(1)), int(m.group(2)))
        except ValueError:
            return None
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})$", text)
    if m:
        return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    return None


# ────────────────────────────────────────────
# XML 헬퍼
# ────────────────────────────────────────────

def get_cells(row_xml: str):
    return re.findall(r"<hp:tc\b.*?</hp:tc>", row_xml, re.DOTALL)

def get_texts(xml: str):
    return re.findall(r"<hp:t>(.*?)</hp:t>", xml)

def col_addr(cell: str) -> int:
    m = re.search(r'colAddr="(\d+)"', cell)
    return int(m.group(1)) if m else -1

def row_addr(cell: str) -> int:
    m = re.search(r'rowAddr="(\d+)"', cell)
    return int(m.group(1)) if m else -1

def row_span(cell: str) -> int:
    m = re.search(r'rowSpan="(\d+)"', cell)
    return int(m.group(1)) if m else 1

def cell_height(cell: str) -> int:
    m = re.search(r'<hp:cellSz width="\d+" height="(\d+)"', cell)
    return int(m.group(1)) if m else 0

def set_row_addr(cell: str, new_addr: int) -> str:
    old = row_addr(cell)
    return cell.replace(f'rowAddr="{old}"', f'rowAddr="{new_addr}"', 1)

def set_row_span(cell: str, new_span: int) -> str:
    old = row_span(cell)
    return cell.replace(f'rowSpan="{old}"', f'rowSpan="{new_span}"', 1)

def set_cell_height(cell: str, new_h: int) -> str:
    old_h = cell_height(cell)
    return cell.replace(f'height="{old_h}"', f'height="{new_h}"', 1)

def make_row(cells) -> str:
    return "<hp:tr>" + "".join(cells) + "</hp:tr>"

def row_height(row_xml: str) -> int:
    """행 높이: 레이블(col0) 제외한 첫 셀의 height"""
    for c in get_cells(row_xml):
        if col_addr(c) != 0:
            h = cell_height(c)
            if h > 0:
                return h
    return 2955  # 기본 높이


# ────────────────────────────────────────────
# 테이블 구조 분석
# ────────────────────────────────────────────

KAECHWAL_DATE_COL = 6   # 개찰일 열의 colAddr

def analyze_table(rows):
    """
    Returns dict:
      kaechwal_start, kaechwal_span, kaechwal_col0_cell
      jinhang_start,  jinhang_span,  jinhang_col0_cell
      kaechwal_items, jinhang_items   ← list of item dicts
    """
    result = {}

    for ri, row in enumerate(rows):
        for cell in get_cells(row):
            if col_addr(cell) == 0:
                txt = "".join(get_texts(cell))
                if "개찰" in txt and "kaechwal_start" not in result:
                    result.update(
                        kaechwal_start=ri,
                        kaechwal_span=row_span(cell),
                        kaechwal_col0_cell=cell,
                    )
                elif "진행중" in txt and "jinhang_start" not in result:
                    result.update(
                        jinhang_start=ri,
                        jinhang_span=row_span(cell),
                        jinhang_col0_cell=cell,
                    )

    def section_items(start, span):
        items = []
        for ri in range(start, start + span):
            cells = get_cells(rows[ri])
            data = [c for c in cells if col_addr(c) != 0]
            num_cell  = next((c for c in data if col_addr(c) == 1), None)
            name_cell = next((c for c in data if col_addr(c) == 2), None)
            kd_cell   = next((c for c in data if col_addr(c) == KAECHWAL_DATE_COL), None)
            items.append({
                "row_idx":       ri,
                "cells":         cells,
                "data_cells":    data,
                "num":           "".join(get_texts(num_cell))  if num_cell  else "",
                "name":          "".join(get_texts(name_cell)) if name_cell else "",
                "kaechwal_date": "".join(get_texts(kd_cell))   if kd_cell   else "",
            })
        return items

    if "kaechwal_start" in result:
        result["kaechwal_items"] = section_items(
            result["kaechwal_start"], result["kaechwal_span"])
    if "jinhang_start" in result:
        result["jinhang_items"] = section_items(
            result["jinhang_start"], result["jinhang_span"])

    return result


# ────────────────────────────────────────────
# 메인 변환 로직
# ────────────────────────────────────────────

def process(input_path: str, output_path: str, ref_date: date) -> dict:
    # 파일 읽기
    with zipfile.ZipFile(input_path) as z:
        files     = {n: z.read(n) for n in z.namelist()}
        section   = files["Contents/section0.xml"].decode("utf-8")

    tbl_start = section.find("<hp:tbl")
    tbl_end   = section.find("</hp:tbl>") + len("</hp:tbl>")
    tbl       = section[tbl_start:tbl_end]

    tr_pat = re.compile(r"(<hp:tr>.*?</hp:tr>)", re.DOTALL)
    rows   = tr_pat.findall(tbl)

    info = analyze_table(rows)

    if "jinhang_start" not in info:
        return {"success": False, "error": "진행중 섹션을 찾을 수 없습니다."}

    # 분류: 이동 vs 유지
    to_move, to_keep = [], []
    for item in info["jinhang_items"]:
        if not item["num"]:           # 빈 행 스킵
            continue
        kd = parse_korean_date(item["kaechwal_date"], ref_date.year)
        (to_move if (kd and kd <= ref_date) else to_keep).append(item)

    if not to_move:
        return {
            "success":        True,
            "reference_date": str(ref_date),
            "message":        "이동할 항목이 없습니다.",
            "moved_items":    [],
            "remaining_items":[{"num": i["num"], "name": i["name"],
                                 "kaechwal_date": i["kaechwal_date"]} for i in to_keep],
            "output_path":    output_path,
        }

    # 높이 계산
    rh = {i: row_height(rows[i]) for i in range(len(rows))}
    EMPTY_H = 2955

    # 새 개찰 섹션
    existing_kaechwal = [i for i in info["kaechwal_items"] if i["num"]]
    new_kaechwal_data  = existing_kaechwal + to_move
    new_kaechwal_span  = len(new_kaechwal_data) + 1   # +1 빈 행
    new_kaechwal_h     = sum(rh.get(i["row_idx"], EMPTY_H)
                              for i in new_kaechwal_data) + EMPTY_H

    # 새 진행중 섹션
    new_jinhang_data   = to_keep
    new_jinhang_span   = len(new_jinhang_data) + 1 if new_jinhang_data else 0
    new_jinhang_h      = (sum(rh.get(i["row_idx"], EMPTY_H)
                               for i in new_jinhang_data) + EMPTY_H
                           if new_jinhang_data else 0)

    # 빈 행 템플릿 (data_cells만)
    empty_data_cells = next(
        (i["data_cells"] for i in info["kaechwal_items"] + info["jinhang_items"]
         if not i["num"]),
        info["kaechwal_items"][-1]["data_cells"] if info["kaechwal_items"] else []
    )

    # ── 새 행 조립 ──
    new_rows    = []
    cur         = info["kaechwal_start"]
    k_start     = info["kaechwal_start"]

    # 헤더 행(들) 유지
    for i in range(k_start):
        new_rows.append(rows[i])

    # ── 개찰 섹션 ──
    # 개찰 라벨 셀 수정
    col0_k = info["kaechwal_col0_cell"]
    col0_k = set_row_span(col0_k, new_kaechwal_span)
    col0_k = set_cell_height(col0_k, new_kaechwal_h)
    col0_k = set_row_addr(col0_k, cur)

    first_k = new_kaechwal_data[0]
    new_rows.append(make_row(
        [col0_k] + [set_row_addr(c, cur) for c in first_k["data_cells"]]
    ))
    cur += 1

    for item in new_kaechwal_data[1:]:
        new_rows.append(make_row([set_row_addr(c, cur) for c in item["data_cells"]]))
        cur += 1

    # 개찰 빈 행
    new_rows.append(make_row([set_row_addr(c, cur) for c in empty_data_cells]))
    cur += 1

    # ── 진행중 섹션 ──
    if new_jinhang_data:
        col0_j = info["jinhang_col0_cell"]
        col0_j = set_row_span(col0_j, new_jinhang_span)
        col0_j = set_cell_height(col0_j, new_jinhang_h)
        col0_j = set_row_addr(col0_j, cur)

        first_j = new_jinhang_data[0]
        new_rows.append(make_row(
            [col0_j] + [set_row_addr(c, cur) for c in first_j["data_cells"]]
        ))
        cur += 1

        for item in new_jinhang_data[1:]:
            new_rows.append(make_row([set_row_addr(c, cur) for c in item["data_cells"]]))
            cur += 1

        # 진행중 빈 행
        new_rows.append(make_row([set_row_addr(c, cur) for c in empty_data_cells]))
        cur += 1

    # 테이블 재조립
    new_tbl = tbl.replace("".join(rows), "".join(new_rows), 1)
    new_xml = section[:tbl_start] + new_tbl + section[tbl_end:]

    # HWPX 저장
    with zipfile.ZipFile(output_path, "w") as zout:
        for name, data in files.items():
            if name == "Contents/section0.xml":
                zout.writestr(name, new_xml.encode("utf-8"),
                              compress_type=zipfile.ZIP_DEFLATED)
            elif name == "mimetype":
                zout.writestr(name, data, compress_type=zipfile.ZIP_STORED)
            else:
                zout.writestr(name, data, compress_type=zipfile.ZIP_DEFLATED)

    return {
        "success":         True,
        "reference_date":  str(ref_date),
        "moved_items":     [{"num": i["num"], "name": i["name"],
                              "kaechwal_date": i["kaechwal_date"]} for i in to_move],
        "remaining_items": [{"num": i["num"], "name": i["name"],
                              "kaechwal_date": i["kaechwal_date"]} for i in to_keep],
        "kaechwal_count":  len(new_kaechwal_data),
        "jinhang_count":   len(new_jinhang_data),
        "output_path":     output_path,
    }


# ────────────────────────────────────────────
# Entry point
# ────────────────────────────────────────────

def main():
    args = parse_args()
    ref = (datetime.strptime(args.date, "%Y-%m-%d").date()
           if args.date else date.today())
    result = process(args.input, args.output, ref)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    if not result.get("success"):
        sys.exit(1)


if __name__ == "__main__":
    main()
