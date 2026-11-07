# Nigeria GDP nowcast — 60-calendar-day GDP release-lag scenario

The active package is organized under `Nigeria_GDP_Nowcast`. The model definitions, estimation framework, release-lag scenario, reporting-range method and newsletter files are unchanged.

## Shared data connection

The model reads the common CAPE master workbook directly:

`C:\Users\eakan\Documents\CAPE-ERC\CAPE Inflation Forecasting Model\Main Data\MultipleFrequencyModelData.xlsx`

The local duplicate is not used. Model outputs remain under `Nigeria_GDP_Nowcast\Results`.

## Default complete run

Run from inside `Nigeria_GDP_Nowcast`:

```powershell
python .\GDP_Nowcast_and_Newsletter.py
```

The launcher uses today as the information date, selects the target quarter through the existing release-calendar framework, runs and validates the unchanged model, builds the existing Excel report, refreshes the charts and continues through the newsletter workflow.

For a historical or controlled run:

```powershell
python .\GDP_Nowcast_and_Newsletter.py --as-of 2026-10-14 --target-quarter 2026Q3
```

## Methodological interpretation

Every historical quarter's GDP is assumed to have been published exactly 60 calendar days after quarter-end. This is a scenario assumption, not actual NBS publication history. Explicit GDP publication dates in the calendar are intentionally ignored. Monthly-series release rules are unchanged.

The package remains a final-vintage pseudo-real-time evaluation rather than a genuine historical-vintage test. The eight models, ensemble procedure and 10%-adjusted CAPE reporting range are preserved as supplied.

## Outputs

- `Results\Nigeria_GDP_8_Model_Results.xlsx`
- `Results\forecast.json`
- supporting CSV audit files in `Results`
- newsletter documents under `..\CAPE ERC Economic Newsletter\YYYY\Month Newsletter`

The `Backup` folder is separate from the active package and must not be modified by routine runs.
