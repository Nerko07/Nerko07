# Provjera STANDARDI_SPN prema ISBIH katalogu

Izvor: https://isbih.gov.ba/p/katalog (provjera 28.10.2026, Europe/Sarajevo)

## Deliverables

| Fajl | Opis |
|------|------|
| `STANDARDI_SPN_validated.xlsx` | Original + kolona **Nova oznaka**, ažurirana Izdanje/Važeće izdanje, Status ISBIH |
| `standardi_nedostajuci_toplota.xlsx` / `.csv` | Važeći BAS standardi (mjerenje / obračun / isporuka / proračun) koji **nisu** u STANDARDI_SPN |

## Rezime validacije kolone A (29 oznaka)

- **21** važeće
- **5** promijenjena oznaka → popunjena kolona *Nova oznaka*
- **1** povučen bez važeće BAS zamjene (`BAS ISO 9836:2018`)
- **2** nisu pronađene (`BAS ISO 31004:2020`, `BAS IWA 31000:2022`)

### Promjene oznaka

| Stara | Nova |
|-------|------|
| BAS EN 1434-3:2017 | BAS EN 1434-3:2026 |
| BAS EN 13757-2:2019 | BAS EN 13757-2+A1:2025 |
| BAS EN 13757-3:2019 | BAS EN 13757-3:2026 |
| BAS EN 13757-4:2020 | BAS EN 13757-4:2026 |
| BAS EN 13757-7:2019 | BAS EN 13757-7:2026 |

### Napomene

- `BAS ISO 9836:2018` (korisna/površinske i zapreminske mjere zgrada): sva BAS izdanja su **povučena**; nema važeće BAS oznake u katalogu.
- Šira lista infrastrukturnih DH cijevi: `standardi_nedostajuci_isporuka_infra.csv`.
