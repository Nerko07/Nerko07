#!/usr/bin/env python3
"""
Validate STANDARDI_SPN (column A) against ISBIH catalogue.

Usage:
  python3 validate_standardi_spn.py /path/to/STANDARDI_SPN.xlsx

Outputs:
  - STANDARDI_SPN_validated.xlsx  (original + Nova oznaka column when designation changed)
  - missing_heat_related.csv      (catalogue standards not in file, domain-filtered)
"""

from __future__ import annotations

import csv
import json
import re
import sys
import time
import urllib.parse
import urllib.request
from pathlib import Path

BASE = "https://isbih.gov.ba"
OUT = Path("/workspace/isbih_standards_check")


def api(path: str, **params):
    items = []
    for k, v in params.items():
        if isinstance(v, list):
            for x in v:
                items.append((k, x))
        elif v is not None:
            items.append((k, v))
    qs = urllib.parse.urlencode(items, doseq=True)
    url = f"{BASE}/{path}?{qs}"
    req = urllib.request.Request(
        url, headers={"Accept": "application/json", "User-Agent": "Mozilla/5.0"}
    )
    with urllib.request.urlopen(req, timeout=60) as r:
        return json.load(r)


def normalize_code(code: str) -> str:
    c = (code or "").strip()
    c = re.sub(r"\s+", " ", c)
    return c


def strip_year(code: str) -> str:
    c = normalize_code(code)
    c = re.sub(r"\+\s*A\d+(:\d{4})?$", "", c)
    c = re.sub(r":\d{4}.*$", "", c)
    return c.strip()


def lookup_standard(code: str) -> dict:
    """Return match info for a designation from column A."""
    q = normalize_code(code)
    if not q:
        return {"input": code, "found": False}

    # Prefer exact-ish search with all statuses first
    all_items = api(
        "catalogue-api/list", page=0, simple_search=q, standard_status=""
    ).get("items") or []
    valid_items = api(
        "catalogue-api/list", page=0, simple_search=q, standard_status="2"
    ).get("items") or []

    def score(it):
        sc = it.get("standard_code") or ""
        if sc == q:
            return 0
        if strip_year(sc) == strip_year(q):
            return 1
        if strip_year(q) in strip_year(sc) or strip_year(sc) in strip_year(q):
            return 2
        return 9

    all_items = sorted(all_items, key=score)
    valid_items = sorted(valid_items, key=score)

    exact_valid = [i for i in valid_items if i.get("standard_code") == q]
    family_valid = [
        i
        for i in valid_items
        if strip_year(i.get("standard_code") or "") == strip_year(q)
    ]
    exact_any = [i for i in all_items if i.get("standard_code") == q]

    result = {
        "input": q,
        "found": bool(all_items),
        "exact_valid": bool(exact_valid),
        "is_valid": bool(exact_valid),
        "valid_codes": [i.get("standard_code") for i in family_valid],
        "exact_status_id": (exact_any[0].get("status_id") if exact_any else None),
        "suggested_new_code": None,
        "title": None,
    }
    if exact_valid:
        result["title"] = exact_valid[0].get("title")
        result["is_valid"] = True
    elif family_valid:
        # designation / year changed — propose newest valid code(s)
        result["suggested_new_code"] = "; ".join(
            sorted({i.get("standard_code") for i in family_valid})
        )
        result["title"] = family_valid[0].get("title")
        result["is_valid"] = False
    elif exact_any:
        result["title"] = exact_any[0].get("title")
        result["is_valid"] = exact_any[0].get("status_id") == "2"
        if family_valid:
            result["suggested_new_code"] = "; ".join(
                sorted({i.get("standard_code") for i in family_valid})
            )
    return result


def read_column_a(path: Path) -> list[str]:
    try:
        import openpyxl
    except ImportError:
        import subprocess

        subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl", "-q"])
        import openpyxl

    wb = openpyxl.load_workbook(path, data_only=True)
    # Prefer sheet named like STANDARDI, else active
    ws = None
    for name in wb.sheetnames:
        if "standard" in name.lower() or "spn" in name.lower():
            ws = wb[name]
            break
    if ws is None:
        ws = wb.active
    codes = []
    for row in ws.iter_rows(min_col=1, max_col=1, min_row=1):
        val = row[0].value
        if val is None:
            continue
        s = str(val).strip()
        if not s or s.lower() in {"oznaka", "oznaka standarda", "standard", "code"}:
            continue
        codes.append(s)
    return codes, wb, ws


def main():
    if len(sys.argv) < 2:
        print("Usage: validate_standardi_spn.py STANDARDI_SPN.xlsx")
        sys.exit(1)
    path = Path(sys.argv[1])
    codes, wb, ws = read_column_a(path)
    print(f"Loaded {len(codes)} codes from {path}")

    # Find header row / insert Nova oznaka column after A
    header_row = 1
    if ws["A1"].value and "oznaka" not in str(ws["A1"].value).lower():
        # maybe no header — insert header at top
        ws.insert_rows(1)
        ws["A1"] = "Oznaka standarda"
        header_row = 1
        # shift codes mentally: data starts row 2
    # Ensure column B title for new designation
    # If B already used, insert column at B
    existing_b = ws.cell(header_row, 2).value
    if existing_b and "nova" not in str(existing_b).lower():
        ws.insert_cols(2)
    ws.cell(header_row, 2).value = "Nova oznaka (ISBIH važeća)"
    ws.cell(header_row, 3).value = ws.cell(header_row, 3).value or "Status ISBIH"
    ws.cell(header_row, 4).value = ws.cell(header_row, 4).value or "Napomena"

    results = []
    for i, code in enumerate(codes):
        info = lookup_standard(code)
        results.append(info)
        row = header_row + 1 + i
        # Keep A as original
        if ws.cell(row, 1).value is None:
            ws.cell(row, 1).value = code
        if info.get("suggested_new_code"):
            ws.cell(row, 2).value = info["suggested_new_code"]
        elif info.get("exact_valid"):
            ws.cell(row, 2).value = ""  # unchanged / still valid
        else:
            ws.cell(row, 2).value = info.get("suggested_new_code") or ""
        if info.get("exact_valid"):
            status = "Važeći"
            note = ""
        elif info.get("suggested_new_code"):
            status = "Promijenjena oznaka / novo izdanje"
            note = f"Važeće: {info['suggested_new_code']}"
        elif info.get("found"):
            status = "Pronađen, nije važeći ili nije tačna oznaka"
            note = f"status_id={info.get('exact_status_id')}"
        else:
            status = "Nije pronađen u katalogu"
            note = ""
        ws.cell(row, 3).value = status
        ws.cell(row, 4).value = note
        print(f"{i+1}/{len(codes)} {code} -> {status} {note}")
        time.sleep(0.2)

    out_xlsx = OUT / "STANDARDI_SPN_validated.xlsx"
    wb.save(out_xlsx)
    with open(OUT / "validation_results.json", "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print("Wrote", out_xlsx)

    # Missing heat-related standards
    focused_path = OUT / "focused_enriched.json"
    if focused_path.exists():
        focused = json.load(open(focused_path, encoding="utf-8"))
        present = {strip_year(c).upper() for c in codes}
        present |= {normalize_code(c).upper() for c in codes}
        missing = []
        for row in focused:
            oc = (row.get("oznaka") or "").upper()
            bc = (row.get("base_code") or "").upper()
            if oc in present or bc in present:
                continue
            # also match without BAS prefix
            alt = re.sub(r"^BAS\s+", "", bc).upper()
            if any(alt and alt in p for p in present):
                continue
            if (row.get("status") or "").lower().startswith("važe") or row.get(
                "status_id"
            ) == "2":
                missing.append(row)
        miss_csv = OUT / "missing_heat_related.csv"
        with open(miss_csv, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f, delimiter=";")
            w.writerow(
                ["Oznaka standarda", "Naziv standarda", "Izdanje", "Važeće izdanje"]
            )
            for r in sorted(missing, key=lambda x: x.get("oznaka") or ""):
                w.writerow(
                    [
                        r.get("oznaka"),
                        r.get("naziv"),
                        r.get("izdanje"),
                        r.get("vazece_izdanje") or r.get("oznaka"),
                    ]
                )
        print("Wrote", miss_csv, "rows", len(missing))


if __name__ == "__main__":
    main()
