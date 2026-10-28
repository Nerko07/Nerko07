#!/usr/bin/env python3
"""Tighten candidate list and enrich with edition / replacement from detail pages."""

from __future__ import annotations

import csv
import json
import os
import re
import time
import urllib.request
from html import unescape

OUT = "/workspace/isbih_standards_check"
BASE = "https://isbih.gov.ba"

# Focused on: useful floor area, heat delivery, billing/allocation, heat measurement
INCLUDE_CODE = re.compile(
    r"""(?ix)
    \bEN\s*(ISO\s*)?(
        1434(?:-\d+)? |          # heat meters
        834 | 835 |              # heat cost allocators
        9836 |                   # floor area / volume of buildings
        12831(?:-\d+)? |         # design heat load calculation
        15316(?:-\d+(-\d+)?)? |  # energy performance / heating systems calc
        15378(?:-\d+)? |         # inspection of heating systems / boilers
        13941(?:-\d+)? |         # district heating pipes (delivery)
        14419 |                  # district heating pipe systems - surveillance
        15603 |                  # overall energy use / EPBD calculation (legacy)
        52000-\d+ | 52016-\d+ | 52017-\d+ | 52018-\d+ |  # EPB overarching
        15450 |                  # heating systems design for heat pumps (borderline)
        14336 |                  # installation/commissioning of water based heating
        12828 |                  # design of water-based heating systems
        12170 | 12171 |          # heating systems - operating instructions / procedure
        676 |                    # gas-fired boilers? skip later if needed
        442-\d+ |                # radiators (emitters - delivery in space)
        215 |                    # thermostatic radiator valves
        1264-\d+ |               # floor heating
        6700 |                   # ?
        16911 |                  # district heating guidelines TR
        15375 |                  # district heating stations?
        13941 | 14419 |
        253 |                    # district heating pipes PE?
        15665 |                  # ventilation related energy?
        16798 |                  # energy performance ventilation
        17463 |                  # ?
        ISO\s*17742 |            # energy efficiency / savings calculation?
        ISO\s*50015 |            # M&V energy performance
        ISO\s*50006             # energy performance indicators
    )\b
    """
)

INCLUDE_TITLE = re.compile(
    r"""(?ix)
    (
      mjeril[ao]\s+toplotne\s+energ |
      heat\s*meters? |
      kalorimet(?:ri|ar) |
      raspodjel[ae]\s+troškova\s+(grijanja|toplot) |
      heat\s*cost\s*allocat |
      indikator(?:i)?\s+utroška\s+topl |
      alokator(?:i)?\s+trošk |
      obra[cč]un.{0,40}(toplot|grijan|energij) |
      isporuk[ae].{0,40}(toplot|toplin|energij) |
      mjerenj[ea].{0,40}(toplotne\s+energ|toplote\b|toplinske\s+energ) |
      korisn[ae]\s+povr[sš]in |
      povr[sš]in[ae]\s+i\s+zapremin |
      floor\s*area |
      definicij[ae].{0,20}(povr|zapremin).{0,30}zgrad |
      prora[cč]un.{0,30}(toplotn.{0,15}optere|design\s*heat|grijan) |
      daljinsk[aoe].{0,15}grijan |
      district\s*heating |
      predajne\s+stanic |
      toplotne\s+podstanic |
      energetsk[ae]\s+svojstv[ao]\s+zgrad |
      energy\s*performance\s*of\s*buildings
    )
    """
)

EXCLUDE_TITLE = re.compile(
    r"""(?ix)
    (
      električn[ae]\s+instalac |
      kablov |
      solarne?\s+termoelektr |
      mineraln[ao]\s+izolacijsk |
      požar|vatru|pozar |
      PVC\s+cijev |
      diferencijaln.{0,20}skeniranj.{0,20}kalorimet |
      DSC\b |
      štapni\s+termometr |
      termometri\s+za\s+kalorimet |
      fenestrac |
      plime\s+i\s+oseke |
      energija\s+mora |
      komunikacijske\s+mreže |
      automatizacij[ae]\s+u\s+elektro |
      toplotno\s+izolacion[aei]\s+proizvod |
      proizvodi\s+za\s+toplotnu\s+izol |
      površinsk[aoe]\s+obrade? |
      zaštita\s+od\s+toplotnog\s+djelovanja |
      rashladni\s+sistem |
      toplotne?\s+pump |
      klimatizac |
      uređaji\s+za\s+klimatiz |
      solarne?\s+kolektor |
      pelet|biomasa\s+kotao |
      dimovod |
      plinski\s+aparat(?!\s+za\s+grijanje\s+prostora\s+sa\s+ventilator)
    )
    """
)


def fetch_detail(page_link: str) -> dict:
    url = BASE + page_link if page_link.startswith("/") else page_link
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=60) as r:
        html = r.read().decode("utf-8", errors="ignore")

    def grab(label: str) -> str:
        m = re.search(
            rf"<span>\s*{re.escape(label)}\s*:</span>\s*([^<]+)",
            html,
            re.I,
        )
        return unescape(m.group(1)).strip() if m else ""

    status = grab("Status")
    edition = grab("Izdanje")
    pub = grab("Datum objave")
    method = grab("Metoda usvajanja")

    # replacements / links with other BAS standards
    replacements = []
    for m in re.finditer(
        r'href="(/standard/\d+)"[^>]*>\s*(BAS[^<]+?)\s*<',
        html,
    ):
        code = unescape(m.group(2)).strip()
        if code and code not in replacements:
            replacements.append(code)

    # notes about withdrawn / replaced
    note = ""
    m = re.search(r"replacement-action.*?>(.*?)</div>", html, re.S | re.I)
    if m:
        note = re.sub(r"<[^>]+>", " ", m.group(1))
        note = unescape(re.sub(r"\s+", " ", note)).strip()

    title_en = ""
    m = re.search(r"Naslov EN.*?<[^>]+>([^<]+)", html, re.S | re.I)
    # simpler: look in page h1/h2
    m2 = re.search(r'<h1[^>]*>\s*(.*?)\s*</h1>', html, re.S | re.I)
    h1 = re.sub(r"<[^>]+>", " ", m2.group(1)) if m2 else ""
    h1 = unescape(re.sub(r"\s+", " ", h1)).strip()

    return {
        "status": status,
        "edition": edition,
        "publication_date": pub,
        "method": method,
        "related_codes": replacements[:20],
        "replacement_note": note,
        "page_title": h1,
    }


def base_code(standard_code: str) -> str:
    """BAS EN 1434-1:2023 -> BAS EN 1434-1"""
    c = standard_code or ""
    c = re.sub(r"\+\s*A\d+:\d{4}$", "", c)
    c = re.sub(r":\d{4}.*$", "", c)
    return c.strip()


def main():
    raw = json.load(open(f"{OUT}/relevant_candidates.json", encoding="utf-8"))
    focused = []
    for it in raw:
        code = it.get("standard_code") or ""
        title = " ".join([it.get("title") or "", it.get("title_ba") or ""])
        if EXCLUDE_TITLE.search(title) and not INCLUDE_CODE.search(code):
            continue
        if INCLUDE_CODE.search(code) or INCLUDE_TITLE.search(title):
            # extra hard exclude still
            if EXCLUDE_TITLE.search(title) and not re.search(
                r"1434|834|835|9836|12831|15316|13941|14419|5200|15603|12828|14336|12170|12171|15378|16911|442-|215\b|1264-",
                code,
                re.I,
            ):
                continue
            focused.append(it)

    print(f"focused count: {len(focused)}")
    enriched = []
    for i, it in enumerate(focused):
        link = it.get("page_link") or f"/standard/{it['id']}"
        try:
            detail = fetch_detail(link)
        except Exception as e:
            detail = {"error": str(e)}
        row = {
            "oznaka": it.get("standard_code"),
            "naziv": it.get("title") or it.get("title_ba"),
            "status_id": it.get("status_id"),
            "status": detail.get("status"),
            "izdanje": detail.get("edition"),
            "datum_objave": (it.get("publication_date") or "")[:10]
            or (detail.get("publication_date") or "")[:10],
            "page_link": BASE + link,
            "related_codes": detail.get("related_codes") or [],
            "replacement_note": detail.get("replacement_note") or "",
            "base_code": base_code(it.get("standard_code") or ""),
            "_via": it.get("_via"),
        }
        enriched.append(row)
        print(f"[{i+1}/{len(focused)}] {row['oznaka']} | izdanje={row['izdanje']} | {row['status']}")
        time.sleep(0.25)

    # For each base_code family, determine current valid designation
    by_base: dict[str, list] = {}
    for row in enriched:
        by_base.setdefault(row["base_code"], []).append(row)

    for base, rows in by_base.items():
        valid = [r for r in rows if (r.get("status") or "").lower().startswith("važe") or r.get("status_id") == "2"]
        valid_codes = sorted({r["oznaka"] for r in valid})
        for r in rows:
            r["vazece_izdanje"] = "; ".join(valid_codes) if valid_codes else ""

    with open(f"{OUT}/focused_enriched.json", "w", encoding="utf-8") as f:
        json.dump(enriched, f, ensure_ascii=False, indent=2)

    # CSV for user: Oznaka; Naziv; Izdanje; Važeće izdanje
    csv_path = f"{OUT}/standardi_toplota_kandidati.csv"
    with open(csv_path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(
            [
                "Oznaka standarda",
                "Naziv standarda",
                "Izdanje",
                "Važeće izdanje",
                "Status",
                "Datum objave",
                "Link",
            ]
        )
        for r in sorted(enriched, key=lambda x: x["oznaka"] or ""):
            w.writerow(
                [
                    r["oznaka"],
                    r["naziv"],
                    r["izdanje"],
                    r["vazece_izdanje"],
                    r["status"],
                    r["datum_objave"],
                    r["page_link"],
                ]
            )
    print("wrote", csv_path)


if __name__ == "__main__":
    main()
