# Provjera BAS standarda (ISBIH katalog) — SPM / toplotna energija

Izvor: [ISBIH Katalog](https://isbih.gov.ba/p/katalog) (`catalogue-api/list`, `standard_status=2` = Važeći).

## Status

1. **STANDARDI_SPN.xlsx** (kolona A) — još nije učitan u okruženje. Nakon uploada/povezivanja OneDrive pokrenuti:
   ```bash
   python3 validate_standardi_spn.py /putanja/do/STANDARDI_SPN.xlsx
   ```
   Skripta dodaje kolonu **Nova oznaka (ISBIH važeća)** gdje je oznaka/izdanje promijenjeno i generiše listu standarda koji nedostaju.

2. **Katalog** — pretraženi važeći BAS standardi vezani za mjerenje, obračun/raspodjelu, isporuku (daljinsko grijanje) i proračun toplotne energije/opterećenja.

## Ključni nalazi iz kataloga

| Tema | Važeći BAS primjeri |
|------|---------------------|
| Mjerenje toplotne energije | `BAS EN 1434-1:2023` … `1434-6:2023`, `BAS EN 1434-3:2026`; smjernice `BAS CEN/TR 13582:2022` |
| Raspodjela / obračun troškova | `BAS EN 834:2014` (+ Cor1:2016), `BAS EN 835:2010` |
| Isporučni cjevovodi / DH | `BAS EN 253+A1:2025`, `13941-1/2+A1:2023`, `14419:2020`, `15632-*`, `15698-*`, `448`, `488-*`, `489-1` |
| Proračun opterećenja / sistema | `BAS EN 12831-1:2021`, `15316-*`, `15378-*`, `12828+A1:2015` |
| Korisna površina prostora | **Nije pronađen** BAS ekvivalent `ISO 9836` / naslovi o korisnoj površini u važećem katalogu |

## Fajlovi

- `standardi_toplota_jezgro.csv` — sužena lista (Oznaka; Naziv; Izdanje; Važeće izdanje; Link)
- `standardi_toplota_kandidati.csv` — širi skup kandidata
- `validate_standardi_spn.py` — validacija Excel kolone A vs ISBIH
- `fetch_catalog.py` / `refine_and_enrich.py` — ponovno prikupljanje iz kataloga
