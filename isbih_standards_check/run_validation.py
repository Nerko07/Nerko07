#!/usr/bin/env python3
"""Validate STANDARDI_SPN against ISBIH and produce deliverables."""

from __future__ import annotations

import csv
import json
import re
import time
import urllib.parse
import urllib.request
from copy import copy
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

BASE = "https://isbih.gov.ba"
OUT = Path("/workspace/isbih_standards_check")
SRC = Path("/home/ubuntu/.cursor/projects/workspace/uploads/STANDARDI_SPN_95ee.xlsx")
CORE_JSON = OUT / "core_heat_delivery_measurement.json"
FOCUSED_JSON = OUT / "focused_enriched.json"


def api(path: str, **params):
    items = []
    for k, v in params.items():
        if isinstance(v, list):
            for x in v:
                items.append((k, x))
        elif v is not None:
            items.append((k, v))
    qs = urllib.parse.urlencode(items, doseq=True)
    req = urllib.request.Request(
        f"{BASE}/{path}?{qs}",
        headers={"Accept": "application/json", "User-Agent": "Mozilla/5.0"},
    )
    with urllib.request.urlopen(req, timeout=60) as r:
        return json.load(r)


def normalize(code: str) -> str:
    return re.sub(r"\s+", " ", (code or "").strip())


def strip_year(code: str) -> str:
    c = normalize(code)
    # Keep /Cor1 as part of family distinction optionally
    c = re.sub(r"\+\s*A\d+(:\d{4})?$", "", c)
    c = re.sub(r":\d{4}.*$", "", c)
    return c.strip()


def family_key(code: str) -> str:
    """Family without year, but keep Cor/Amd markers without year."""
    c = normalize(code)
    c = re.sub(r":\d{4}$", "", c)
    c = re.sub(r"(/(?:Cor|A)\d+):\d{4}$", r"\1", c)
    return c.strip()


def fetch_detail(page_link: str) -> dict:
    url = BASE + page_link if page_link.startswith("/") else page_link
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=60) as r:
        html = r.read().decode("utf-8", errors="ignore")

    def grab(label: str) -> str:
        m = re.search(rf"<span>\s*{re.escape(label)}\s*:</span>\s*([^<]+)", html, re.I)
        return m.group(1).strip() if m else ""

    return {
        "status": grab("Status"),
        "edition": grab("Izdanje").rstrip("."),
        "publication_date": grab("Datum objave"),
    }


def search_code(code: str) -> dict:
    q = normalize(code)
    # Broad search then filter client-side for better matching
    # Try exact-ish queries
    queries = [q, strip_year(q), family_key(q)]
    # Also without BAS prefix
    no_bas = re.sub(r"^BAS\s+", "", q)
    queries.append(no_bas)
    queries.append(strip_year(no_bas))

    all_hits = {}
    for qq in dict.fromkeys(queries):  # unique preserve order
        for status in ["", "2", "3"]:
            d = api(
                "catalogue-api/list",
                page=0,
                simple_search=qq,
                standard_status=status if status else None,
            )
            for it in d.get("items") or []:
                all_hits[it["id"]] = it
            time.sleep(0.08)

    items = list(all_hits.values())

    def score(it):
        sc = normalize(it.get("standard_code") or "")
        if sc == q:
            return 0
        if family_key(sc) == family_key(q):
            return 1
        if strip_year(sc) == strip_year(q):
            return 2
        # numeric series match e.g. 1434-3
        return 9

    items = sorted(items, key=score)
    exact = [i for i in items if normalize(i.get("standard_code") or "") == q]
    same_family = [
        i for i in items if family_key(i.get("standard_code") or "") == family_key(q)
    ]
    same_base = [
        i for i in items if strip_year(i.get("standard_code") or "") == strip_year(q)
    ]

    valid_family = [i for i in same_family if i.get("status_id") == "2"]
    valid_base = [i for i in same_base if i.get("status_id") == "2"]
    # Prefer family, else base (year/edition changes)
    valid = valid_family or valid_base

    result = {
        "input": q,
        "exact_found": bool(exact),
        "exact_status_id": exact[0].get("status_id") if exact else None,
        "exact_title": (exact[0].get("title") if exact else None),
        "exact_id": exact[0].get("id") if exact else None,
        "exact_link": exact[0].get("page_link") if exact else None,
        "valid_codes": sorted({normalize(i.get("standard_code")) for i in valid}),
        "valid_items": valid,
        "all_family_codes": sorted(
            {normalize(i.get("standard_code")) for i in (same_family or same_base)}
        ),
    }

    if exact and exact[0].get("status_id") == "2":
        result["verdict"] = "VAZECI"
        result["nova_oznaka"] = ""
        result["napomena"] = "Oznaka je važeća u ISBIH katalogu."
    elif valid and exact and exact[0].get("status_id") != "2":
        result["verdict"] = "PROMIJENJENA"
        result["nova_oznaka"] = "; ".join(result["valid_codes"])
        result["napomena"] = (
            f"Unesena oznaka nije važeća (status_id={exact[0].get('status_id')}). "
            f"Važeća oznaka: {result['nova_oznaka']}"
        )
    elif valid and not exact:
        result["verdict"] = "PROMIJENJENA"
        result["nova_oznaka"] = "; ".join(result["valid_codes"])
        result["napomena"] = (
            f"Tačna oznaka nije pronađena; važeća srodna oznaka: {result['nova_oznaka']}"
        )
    elif exact:
        result["verdict"] = "POVUCEN"
        result["nova_oznaka"] = ""
        result["napomena"] = (
            f"Pronađena u katalogu ali nije važeća (status_id={exact[0].get('status_id')}); "
            f"nema važeće zamjene u istoj porodici."
        )
    else:
        result["verdict"] = "NIJE_PRONADJEN"
        result["nova_oznaka"] = ""
        result["napomena"] = "Nije pronađen u ISBIH katalogu (pretraga po oznaci)."

    # Enrich edition for exact or first valid
    target = None
    if exact and exact[0].get("status_id") == "2":
        target = exact[0]
    elif valid:
        target = valid[0]
    elif exact:
        target = exact[0]

    result["izdanje"] = ""
    result["status_label"] = ""
    result["title"] = result.get("exact_title")
    if target:
        try:
            detail = fetch_detail(target.get("page_link") or f"/standard/{target['id']}")
            result["izdanje"] = detail.get("edition") or ""
            result["status_label"] = detail.get("status") or ""
            result["title"] = target.get("title") or result.get("title")
            result["detail_code"] = normalize(target.get("standard_code"))
        except Exception as e:
            result["detail_error"] = str(e)
        time.sleep(0.2)

    # If multiple valid, get editions for nova oznaka primary
    if result["verdict"] == "PROMIJENJENA" and valid:
        try:
            d = fetch_detail(valid[0].get("page_link"))
            result["vazece_izdanje"] = d.get("edition") or ""
            result["izdanje_nove"] = d.get("edition") or ""
        except Exception:
            result["vazece_izdanje"] = ""
        time.sleep(0.15)
    elif result["verdict"] == "VAZECI":
        result["vazece_izdanje"] = result["izdanje"]
    else:
        result["vazece_izdanje"] = result.get("izdanje") or ""

    return result


def main():
    wb = openpyxl.load_workbook(SRC)
    ws = wb["STANDARDI"]

    # Insert "Nova oznaka" as column B (after Oznaka)
    # Current: A Oznaka, B Naziv, C Izdanje, D Važeće izdanje
    ws.insert_cols(2)
    ws["B1"] = "Nova oznaka"
    ws["E1"] = "Važeće izdanje"  # was shifted: A oznaka, B nova, C naziv, D izdanje, E vazeće
    # After insert: A=Oznaka, B=Nova, C=Naziv, D=Izdanje, E=Važeće, F=old empty
    # Fix headers
    ws["A1"] = "Oznaka standarda"
    ws["B1"] = "Nova oznaka"
    ws["C1"] = "Naziv standarda"
    ws["D1"] = "Izdanje"
    ws["E1"] = "Važeće izdanje"
    ws["F1"] = "Status ISBIH"
    ws["G1"] = "Napomena"

    header_font = Font(bold=True)
    for col in range(1, 8):
        ws.cell(1, col).font = header_font

    fill_ok = PatternFill("solid", fgColor="C6EFCE")
    fill_chg = PatternFill("solid", fgColor="FFEB9C")
    fill_bad = PatternFill("solid", fgColor="FFC7CE")

    results = []
    data_rows = []
    for r in range(2, ws.max_row + 1):
        code = ws.cell(r, 1).value  # still in A before we filled B
        # After insert_cols, original B(naziv) is now C, etc.
        if code is None or not str(code).strip():
            continue
        if str(code).strip().lower().startswith("oznaka"):
            continue
        data_rows.append(r)

    print(f"Validating {len(data_rows)} standards...")
    for idx, r in enumerate(data_rows):
        code = str(ws.cell(r, 1).value).strip()
        print(f"[{idx+1}/{len(data_rows)}] {code}")
        info = search_code(code)
        results.append(info)

        ws.cell(r, 2).value = info.get("nova_oznaka") or None
        # Update title from catalog if empty? keep original naziv in C
        if info.get("izdanje"):
            # For changed codes, Izdanje = edition of entered (if known) else leave;
            # Važeće izdanje = edition of valid
            if info["verdict"] == "VAZECI":
                ws.cell(r, 4).value = (
                    int(info["izdanje"]) if info["izdanje"].isdigit() else info["izdanje"]
                )
                ws.cell(r, 5).value = ws.cell(r, 4).value
            elif info["verdict"] == "PROMIJENJENA":
                # keep/update entered edition if we had exact detail
                if info.get("exact_found") and info.get("izdanje") and info.get("detail_code") == code:
                    ws.cell(r, 4).value = (
                        int(info["izdanje"])
                        if str(info["izdanje"]).isdigit()
                        else info["izdanje"]
                    )
                v = info.get("vazece_izdanje") or info.get("izdanje_nove") or ""
                if v:
                    ws.cell(r, 5).value = int(v) if str(v).isdigit() else v
                # If we only have valid item detail, set both meaningfully
                if not ws.cell(r, 4).value and info.get("izdanje"):
                    # edition of old unknown; leave
                    pass
                if info.get("vazece_izdanje"):
                    ws.cell(r, 5).value = (
                        int(info["vazece_izdanje"])
                        if str(info["vazece_izdanje"]).isdigit()
                        else info["vazece_izdanje"]
                    )
            else:
                if info.get("izdanje"):
                    ws.cell(r, 4).value = (
                        int(info["izdanje"])
                        if str(info["izdanje"]).isdigit()
                        else info["izdanje"]
                    )

        # For PROMIJENJENA, fetch edition of valid code specifically already in vazece_izdanje
        # Also update izdanje of entered if withdrawn exact exists - need separate detail
        if info["verdict"] == "PROMIJENJENA" and info.get("exact_link"):
            try:
                d_old = fetch_detail(info["exact_link"])
                if d_old.get("edition"):
                    ed = d_old["edition"]
                    ws.cell(r, 4).value = int(ed) if ed.isdigit() else ed
                time.sleep(0.15)
            except Exception:
                pass

        status_map = {
            "VAZECI": "Važeći",
            "PROMIJENJENA": "Promijenjena oznaka",
            "POVUCEN": "Povučen / nije važeći",
            "NIJE_PRONADJEN": "Nije pronađen",
        }
        ws.cell(r, 6).value = status_map.get(info["verdict"], info["verdict"])
        ws.cell(r, 7).value = info.get("napomena")

        fill = {
            "VAZECI": fill_ok,
            "PROMIJENJENA": fill_chg,
            "POVUCEN": fill_bad,
            "NIJE_PRONADJEN": fill_bad,
        }.get(info["verdict"])
        if fill:
            for c in range(1, 8):
                ws.cell(r, c).fill = fill

        print(f"  -> {info['verdict']} nova={info.get('nova_oznaka')!r}")

    out_xlsx = OUT / "STANDARDI_SPN_validated.xlsx"
    wb.save(out_xlsx)
    with open(OUT / "validation_results.json", "w", encoding="utf-8") as f:
        # strip non-serializable
        serial = []
        for r in results:
            serial.append({k: v for k, v in r.items() if k != "valid_items"})
        json.dump(serial, f, ensure_ascii=False, indent=2)
    print("Wrote", out_xlsx)

    # Missing heat-related standards
    present_keys = set()
    for r in results:
        present_keys.add(strip_year(r["input"]).upper())
        present_keys.add(family_key(r["input"]).upper())
        present_keys.add(normalize(r["input"]).upper())
        for vc in r.get("valid_codes") or []:
            present_keys.add(strip_year(vc).upper())
            present_keys.add(family_key(vc).upper())
            present_keys.add(normalize(vc).upper())

    core = []
    if CORE_JSON.exists():
        core = json.load(open(CORE_JSON, encoding="utf-8"))
    elif FOCUSED_JSON.exists():
        core = json.load(open(FOCUSED_JSON, encoding="utf-8"))

    # Expand core with M-Bus / communication and other relevant from catalog searches
    # Also include EN 1434-3:2026 etc already in core
    missing = []
    for row in core:
        oc = normalize(row.get("oznaka") or "")
        bc = normalize(row.get("base_code") or strip_year(oc))
        keys = {
            oc.upper(),
            bc.upper(),
            strip_year(oc).upper(),
            family_key(oc).upper(),
            re.sub(r"^BAS\s+", "", bc, flags=re.I).upper(),
        }
        if keys & present_keys:
            continue
        # skip Cor1 if base present
        if "/COR" in oc.upper():
            base_wo = re.sub(r"/COR\d+", "", strip_year(oc), flags=re.I).upper()
            if any(base_wo in p for p in present_keys):
                continue
        if (row.get("status") or "").lower().startswith("važe") or row.get("status_id") == "2":
            missing.append(row)

    miss_csv = OUT / "standardi_nedostajuci_toplota.csv"
    with open(miss_csv, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["Oznaka standarda", "Naziv standarda", "Izdanje", "Važeće izdanje"])
        for row in sorted(missing, key=lambda x: x.get("oznaka") or ""):
            w.writerow(
                [
                    row.get("oznaka"),
                    row.get("naziv"),
                    row.get("izdanje"),
                    row.get("vazece_izdanje") or row.get("oznaka"),
                ]
            )

    # Also Excel for missing
    wb2 = openpyxl.Workbook()
    ws2 = wb2.active
    ws2.title = "Nedostajući"
    ws2.append(["Oznaka standarda", "Naziv standarda", "Izdanje", "Važeće izdanje"])
    for row in sorted(missing, key=lambda x: x.get("oznaka") or ""):
        iz = row.get("izdanje") or ""
        iz = iz.rstrip(".") if isinstance(iz, str) else iz
        vaz = row.get("vazece_izdanje") or iz
        if isinstance(vaz, str):
            # vazece_izdanje may be full code list; for missing list use edition number
            vaz = iz
        ws2.append([row.get("oznaka"), row.get("naziv"), iz, vaz])
    for col in ws2.columns:
        max_len = 0
        col_letter = col[0].column_letter
        for cell in col:
            max_len = max(max_len, len(str(cell.value or "")))
        ws2.column_dimensions[col_letter].width = min(max_len + 2, 80)
    miss_xlsx = OUT / "standardi_nedostajuci_toplota.xlsx"
    wb2.save(miss_xlsx)

    print("Wrote", miss_csv, "rows", len(missing))
    print("Wrote", miss_xlsx)

    # Summary
    from collections import Counter

    c = Counter(r["verdict"] for r in results)
    print("SUMMARY", dict(c))
    for r in results:
        if r["verdict"] != "VAZECI":
            print(f"  ! {r['input']}: {r['verdict']} -> {r.get('nova_oznaka')}")


if __name__ == "__main__":
    main()
