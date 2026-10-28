#!/usr/bin/env python3
"""Fetch ISBIH catalogue standards related to heat metering, delivery, billing and floor area."""

from __future__ import annotations

import json
import os
import re
import time
import urllib.parse
import urllib.request
from collections import OrderedDict

BASE = "https://isbih.gov.ba"
OUT = "/workspace/isbih_standards_check"


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


def search_all(**params):
    page = 0
    out = []
    total = 0
    while True:
        d = api("catalogue-api/list", page=page, **params)
        batch = d.get("items") or []
        out.extend(batch)
        total = d.get("total") or 0
        if len(out) >= total or not batch:
            break
        page += 1
        if page > 300:
            break
        time.sleep(0.12)
    return out, total


def main():
    os.makedirs(OUT, exist_ok=True)

    ics_queries = [
        "toplot",
        "toplin",
        "grijan",
        "kalor",
        "energet",
        "zgrad",
        "površin",
        "mjeren",
        "instalac",
        "termoteh",
    ]
    ics_map = {}
    for q in ics_queries:
        d = api("catalogue-api/ics", query=q)
        for it in d.get("item") or []:
            ics_map[it["id"]] = it
        time.sleep(0.1)

    with open(f"{OUT}/ics_candidates.json", "w", encoding="utf-8") as f:
        json.dump(list(ics_map.values()), f, ensure_ascii=False, indent=2)

    prio = []
    for it in ics_map.values():
        desc = (it["description"] + " " + it["code"]).lower()
        if re.search(
            r"toplot|kalor|grijan|heating|heat\b|energetsk|energy performance|"
            r"zgrad.*energ|instalac.*zgrad|centraln|termoteh|mjeren.*temp|"
            r"termodinam",
            desc,
        ):
            prio.append(it)

    print(f"priority ICS: {len(prio)}")
    for it in sorted(prio, key=lambda x: x["code"]):
        print(f"  {it['code']} id={it['id']} {it['description']}")

    seen: OrderedDict = OrderedDict()

    for it in prio:
        items, total = search_all(standard_status="2", ics=it["id"])
        print(f"ICS {it['code']} id={it['id']} total={total}")
        for x in items:
            x["_via"] = f"ics:{it['code']}"
            seen[x["id"]] = x
        time.sleep(0.15)

    queries = [
        "mjerila toplotne",
        "kalorimet",
        "toplotne energije",
        "toplinske energije",
        "raspodjela troškova",
        "heat cost allocator",
        "daljinsko grijanje",
        "daljinskog grijanja",
        "sistemi centralnog grijanja",
        "korisna površina",
        "korisna povrsina",
        "površina i zapremina",
        "povrsina i zapremina",
        "floor area",
        "BAS EN 1434",
        "BAS EN 834",
        "BAS EN 835",
        "BAS EN ISO 9836",
        "BAS EN 12831",
        "BAS EN 15316",
        "BAS EN 15378",
        "BAS EN 13941",
        "BAS EN 14419",
        "BAS EN 15603",
        "BAS EN ISO 52000",
        "BAS EN ISO 52016",
        "BAS EN 12828",
        "BAS EN 14336",
        "BAS EN 1264",
        "BAS EN 442",
        "BAS EN 215",
        "BAS EN 12170",
        "BAS EN 15450",
        "BAS CEN/TR 16911",
        "isporuka toplotne",
        "obračun toplotne",
        "obracun toplotne",
        "mjerenje toplotne",
        "mjerenje toplote",
        "alokatori troškova",
        "indikatori utroška toplote",
    ]
    for q in queries:
        items, total = search_all(simple_search=q, standard_status="2")
        print(f"Q={q!r} total={total}")
        for x in items:
            x.setdefault("_via", f"q:{q}")
            seen[x["id"]] = x
        time.sleep(0.15)

    print("UNIQUE", len(seen))
    with open(f"{OUT}/raw_hits.json", "w", encoding="utf-8") as f:
        json.dump(list(seen.values()), f, ensure_ascii=False, indent=2)

    kw_re = re.compile(
        r"toplot|toplin|kalorimet|heat\b|grijan|heating|daljinsk|district\s*heat|"
        r"mjeril.{0,40}(toplot|energ|toplin)|heat\s*meter|cost\s*allocat|"
        r"raspodjel.{0,40}(trošk|trosk|grijan|topl)|"
        r"obra[cč]un.{0,40}(toplot|energ|grijan|toplin)|"
        r"isporuk.{0,40}(toplot|energ|toplin)|"
        r"korisn.{0,20}povr|floor\s*area|"
        r"povr[sš]in.{0,40}(prostor|zgrad|poda|koris|zapremin)|"
        r"energetsk.{0,30}(svojstv|efikas|perform)|energy\s*performance|"
        r"prora[cč]un.{0,30}(topl|grijan|optere)|design\s*heat|toplotn.{0,20}optere|"
        r"centralnog grijanja|heating systems|radiator|podnog grijanja|floor heating|"
        r"termostat.{0,20}radiator|heat cost|evaporativ|alokator",
        re.I,
    )
    code_re = re.compile(
        r"\bEN(\s|/)?(ISO\s*)?(1434|834|835|9836|12831|15316|15378|13941|14419|"
        r"15603|15232|52000|52016|15450|14336|1264|442|215|12170|12828|16911)\b",
        re.I,
    )

    relevant = []
    for it in seen.values():
        text = " ".join(
            [it.get("standard_code") or "", it.get("title") or "", it.get("title_ba") or ""]
        )
        if kw_re.search(text) or code_re.search(text):
            relevant.append(it)

    print("RELEVANT", len(relevant))
    with open(f"{OUT}/relevant_candidates.json", "w", encoding="utf-8") as f:
        json.dump(relevant, f, ensure_ascii=False, indent=2)

    for it in sorted(relevant, key=lambda x: x.get("standard_code") or ""):
        print(f"{it['standard_code']}\t{(it.get('title') or '')[:110]}")


if __name__ == "__main__":
    main()
