# CBSA Tariff Lookup 🇨🇦

A modern desktop app to look up **Canadian import duties and taxes** by HS code and origin country —
built on the official CBSA Customs Tariff (T2026-2), with US reciprocal surtax mapping.

Developed by **Amir Arshad** — <https://github.com/incredibleamir-dot>

## Features

- 🔍 HS code search (code or keyword) + multi-code basket
- 🌍 Origin country autocomplete (243 CBSA countries) with correct preferential treatment per origin
  (MFN, GPT, LDCT, CCCT, CETA/CEUT, CUSMA/UST/MXT, CPTPP/CPTPT, UKT/CPUKT, …)
- ⚖️ Duty math: ad-valorem (% of value), specific rates (¢/kg, $/tonne, …) via quantity input, GST 5%
- 🇺🇸 US reciprocal surtax mapping with import-date period filtering
- 📊 Native sortable tables + Summary table; export to **Excel / CSV / PDF report**
- 🔄 Built-in **Database updater**: downloads the latest CBSA Access DB, keeps every revision,
  switch revisions from the header, A-vs-B change comparison
- 🧾 Breadcrumb descriptions (no more orphan fragments like *"Within access commitment"*)

## Run it

### Option A — no install (Windows)

1. Download `dist/TariffLookupApp/TariffLookupApp.exe` (fast starter) or `dist/TariffLookup.exe` (single file).
2. Keep `tariff_app.db` next to the EXE.
3. Double-click. (SmartScreen may warn on the unsigned EXE → *More info → Run anyway*.)

### Option B — from source

```bash
pip install ttkbootstrap tabulate pandas openpyxl reportlab pyodbc
python tariff_lookup_app.py
```

### Build the EXE

```bash
pip install pyinstaller
python -m PyInstaller --noconfirm --clean --onedir --windowed --name TariffLookupApp tariff_lookup_app.py
# single-file instead:
python -m PyInstaller --noconfirm --clean --onefile --windowed --name TariffLookup tariff_lookup_app.py
```

## Data & transparency — please read

| Dataset | Source | Refresh |
|---|---|---|
| Customs Tariff rates + country list | Official CBSA Access database | **Automatic** — *Database…* button downloads the latest revision from CBSA and stores it alongside older ones |
| US reciprocal surtax list | Finance Canada *Complete list of US products subject to counter tariffs* | **STATIC snapshot, hard-coded** — compiled 2026-10-01, covering measures of 2025-03-04, 2025-03-13, 2025-04-09 and 2026-09-08. The updater does **not** refresh it |

> Always confirm live surtax status via
> [Finance Canada](https://www.canada.ca/en/department-finance/programs/international-trade-finance-policy/canadas-response-us-tariffs/complete-list-us-products-subject-to-counter-tariffs.html),
> [CBSA Customs Notices](https://www.cbsa-asfc.gc.ca/publications/cn-ad/menu-eng.html) and the
> [Customs Tariff Act](https://laws-lois.justice.gc.ca/eng/acts/C-54.011/) before transacting.
> This tool is for information only.

## Project layout

```
TariffApp/
├── tariff_lookup_app.py   # app (ttkbootstrap UI + lookup/import logic)
├── tariff_app.db          # SQLite: tphs (+rev), country_map, tariff_codes,
│                          #         counter_tariffs, db_meta
├── README.md
└── dist/                  # built EXEs (see above)
```

Headless logic tests live alongside development: import the module and call
`get_duty_data()`, `build_summary_table()`, `diff_revs()`, `export_results()` —
no GUI needed. GUI smoke tests instantiate `App()` without `mainloop()`.

## Disclaimer

For information purposes only — verify classifications, rates and surtax periods
against official CBSA / Government of Canada sources before making import decisions.
