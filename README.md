# CAPE Nigeria GDP Nowcast and Economic Newsletter

This is the operating guide for refreshing the Nigeria quarterly GDP nowcast and
publishing the corresponding monthly CAPE Economic Performance and Prospect Bulletin.
The workflow has one top-level launcher:

```text
GDP_Nowcast_and_Newsletter.py
```

Run that file from the `NEWSLETTER-PROJECT` directory. It reads the central CAPE
master workbook, estimates the GDP model, rebuilds and validates the nowcast workbook,
and then publishes the requested monthly newsletter.

## 1. Locations

Project root:

```text
C:\Users\eakan\Documents\CAPE-ERC\CAPE Inflation Forecasting Model\NEWSLETTER-PROJECT
```

Launcher:

```text
C:\Users\eakan\Documents\CAPE-ERC\CAPE Inflation Forecasting Model\NEWSLETTER-PROJECT\GDP_Nowcast_and_Newsletter.py
```

Master input workbook:

```text
C:\Users\eakan\Documents\CAPE-ERC\CAPE Inflation Forecasting Model\Main Data\MultipleFrequencyModelData.xlsx
```

GDP output folder:

```text
C:\Users\eakan\Documents\CAPE-ERC\CAPE Inflation Forecasting Model\NEWSLETTER-PROJECT\Nigeria GDP Nowcasting
```

Newsletter output folder:

```text
C:\Users\eakan\Documents\CAPE-ERC\CAPE Inflation Forecasting Model\NEWSLETTER-PROJECT\CAPE ERC Economic Newsletter
```

Do not type backslashes before underscores. The filename is
`GDP_Nowcast_and_Newsletter.py`, not `GDP\_Nowcast\_and\_Newsletter.py`.

## 2. Folder structure

```text
CAPE Inflation Forecasting Model/
├── Main Data/
│   └── MultipleFrequencyModelData.xlsx       # master input; read-only during a run
└── NEWSLETTER-PROJECT/
    ├── GDP_Nowcast_and_Newsletter.py         # run this file
    ├── README.md                              # this guide
    ├── CAPE ERC Economic Newsletter/
    │   └── YYYY/
    │       └── Month Newsletter/
    │           ├── CAPE Economic Performance and Prospect Bulletin Month YYYY.docx
    │           └── Plots/
    └── Nigeria GDP Nowcasting/
        ├── Nigeria Quarterly GDP Nowcast.xlsx
        ├── Nigeria GDP Nowcast Operation Manual.docx
        ├── Newsletter Assets/
        │   ├── Nigeria GDP Nowcast Chart.png
        │   └── Nigeria PMI Trend Chart.png
        └── .automation/
            ├── analyze_nowcast.py
            ├── build_workbooks.mjs
            ├── nowcast_results.json
            ├── publish_newsletter.py
            └── verify_refresh.py
```

The `.automation` files are internal components. For a normal monthly refresh, run
only `GDP_Nowcast_and_Newsletter.py`.

## 3. Input data

The workflow reads `Main Data\MultipleFrequencyModelData.xlsx`. It does not update
the master workbook. The launcher calculates its SHA-256 hash before and after the
workflow and stops if the master changes during processing.

Save all intended Excel changes and close the workbook before starting. Unsaved Excel
changes are not read.

### Monthly data

The monthly sheet must be named exactly `monthly data`. Headers are on row 2.

| Column | Header | Description |
|---|---|---|
| A | Date | Monthly period, such as `2026M8` |
| B | HCPI | Headline consumer price index |
| C | Core Inflation CPI | Core consumer price index |
| D | Food Inflation CPI | Food consumer price index |
| E | MPR | Monetary Policy Rate |
| F | BDC | Exchange-rate series |
| G | COP | Crude-oil price |
| H | M2 | Broad money |
| I | CPS | Credit to the private sector |
| J | CPD | Crude-oil production |
| K | PMI | Purchasing Managers' Index |
| L | Inflation Risk | Inflation-risk series used from January 2015 onward |
| M | PCA Inflation Risk | Additional risk field retained in the master workbook |

The active GDP equation currently uses lagged GDP growth, headline inflation,
crude-oil production, PMI, and monthly M2 growth. Other master-workbook fields remain
available to the internal model and workbook reports.

### Quarterly Data

The quarterly sheet must be named exactly `Quarterly Data`.

| Column | Content | Example |
|---|---|---|
| A | Year | `2026` |
| B | Quarter label | `Q2 2026` |
| C | Real GDP growth, year-on-year | Percentage or decimal value |

The model finds the latest official GDP observation and automatically targets the
following quarter. The newsletter month supplied on the command line does not force
the GDP target quarter.

### Data preparation checklist

Before every run:

1. Update all available observations in `monthly data`.
2. Put PMI in column K and Inflation Risk in column L.
3. Add the latest official real GDP observation to `Quarterly Data` when released.
4. Keep the existing date, quarter-label, and worksheet-name formats.
5. Save and close `MultipleFrequencyModelData.xlsx`.
6. Close `Nigeria Quarterly GDP Nowcast.xlsx`.
7. Close the monthly newsletter document that will be updated.

## 4. Requirements

The workflow is designed for Windows PowerShell and requires:

- Python with `numpy`, `openpyxl`, `python-docx`, and `Pillow`;
- Node.js for workbook rendering;
- the existing `.automation\node_modules` dependencies, including
  `@oai/artifact-tool`.

The recommended Python executable is the bundled runtime:

```text
C:\Users\eakan\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe
```

The launcher automatically prefers the bundled Node.js runtime when available. Keep
the `.automation\node_modules` folder in place.

If using another Python installation, install the required packages in that same
environment:

```powershell
python -m pip install numpy openpyxl python-docx pillow
```

The bundled runtime is recommended when `python` points to a new or unconfigured
installation such as Python 3.14.

## 5. Open PowerShell in the project root

```powershell
Set-Location "C:\Users\eakan\Documents\CAPE-ERC\CAPE Inflation Forecasting Model\NEWSLETTER-PROJECT"
```

Confirm that the launcher exists:

```powershell
Get-Item ".\GDP_Nowcast_and_Newsletter.py"
```

The prompt should end with `NEWSLETTER-PROJECT>`.

## 6. Run a dry check first

A dry run verifies paths and prints the commands that would execute. It does not
estimate the model or change output files.

```powershell
& "C:\Users\eakan\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" ".\GDP_Nowcast_and_Newsletter.py" 2026-10 --dry-run
```

If the default Python installation has all dependencies:

```powershell
python ".\GDP_Nowcast_and_Newsletter.py" 2026-10 --dry-run
```

For October 2026, confirm that the dry-run output points to:

- `Main Data\MultipleFrequencyModelData.xlsx` for the input;
- `Nigeria GDP Nowcasting\Nigeria Quarterly GDP Nowcast.xlsx` for the GDP output;
- `CAPE ERC Economic Newsletter\2026\October Newsletter\...docx` for the newsletter.

## 7. Run the complete workflow

Recommended command for October 2026:

```powershell
& "C:\Users\eakan\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" ".\GDP_Nowcast_and_Newsletter.py" 2026-10
```

Short command when the default Python installation is configured:

```powershell
python ".\GDP_Nowcast_and_Newsletter.py" 2026-10
```

The month must use `YYYY-MM` format:

```powershell
python ".\GDP_Nowcast_and_Newsletter.py" 2026-09
python ".\GDP_Nowcast_and_Newsletter.py" 2026-10
python ".\GDP_Nowcast_and_Newsletter.py" 2027-01
```

Omitting the month uses the current calendar month:

```powershell
python ".\GDP_Nowcast_and_Newsletter.py"
```

## 8. What the launcher does

The launcher performs these steps:

1. Resolves the project, master-data, GDP-output, and newsletter-output paths.
2. Verifies that the master workbook and internal automation scripts exist.
3. Hashes the master workbook for read-only protection.
4. Runs `analyze_nowcast.py` to create quarterly features, evaluate the ridge bridge
   model, and write `.automation\nowcast_results.json`.
5. Runs `build_workbooks.mjs` to refresh the GDP workbook and chart assets.
6. Runs `verify_refresh.py` to validate workbook sheets, input positions, forecast
   values, and publication rules.
7. Confirms that the master-workbook hash is unchanged.
8. Runs `publish_newsletter.py` to update the requested bulletin and plots.
9. Checks the master hash again and prints the final output paths.

Model, build, or validation failure prevents newsletter publication.

## 9. Outputs

### GDP workbook

The refreshed workbook remains at:

```text
NEWSLETTER-PROJECT\Nigeria GDP Nowcasting\Nigeria Quarterly GDP Nowcast.xlsx
```

Its main sheets are:

- `Executive Summary`
- `Quarterly Forecast`
- `Model Comparison`
- `Backtest`
- `Model Engine`
- `Monthly Inputs`
- `Quarterly Features`
- `Methodology`
- `Sources`

Charts remain under:

```text
NEWSLETTER-PROJECT\Nigeria GDP Nowcasting\Newsletter Assets
```

### Monthly newsletter

Newsletters are organized by year and month:

```text
NEWSLETTER-PROJECT\CAPE ERC Economic Newsletter\YYYY\Month Newsletter
```

Example:

```text
NEWSLETTER-PROJECT\CAPE ERC Economic Newsletter\2026\October Newsletter\CAPE Economic Performance and Prospect Bulletin October 2026.docx
```

Plots for that edition default to:

```text
NEWSLETTER-PROJECT\CAPE ERC Economic Newsletter\2026\October Newsletter\Plots
```

Missing year, month, and plot folders are created automatically during publication.

## 10. Newsletter template selection

The launcher chooses a template in this order:

1. The file passed with `--newsletter-template`.
2. The existing clean report for the requested month.
3. A monthly file named
   `CAPE Economic Performance and Prospect Bulletin Month YYYY - Template.docx`.
4. The previous month's clean report, including across a year boundary.

For example, if October 2026 does not yet exist, September 2026 can be used as the
template.

Explicit template:

```powershell
python ".\GDP_Nowcast_and_Newsletter.py" 2026-10 --newsletter-template "C:\path\to\October Template.docx"
```

Explicit plot folder:

```powershell
python ".\GDP_Nowcast_and_Newsletter.py" 2026-10 --plots-folder "C:\path\to\Plots"
```

## 11. Forecast and publication rules

- The target is the quarter after the latest official GDP quarter in the master.
- Historical quarters use all three monthly observations.
- The active quarter uses all available observations: one, two, or three months.
- The current ridge bridge uses lagged GDP, headline inflation, crude-oil production,
  PMI, and monthly M2 growth.
- The internal point estimate updates when source data changes.
- A point forecast is published only when all three target-quarter months contain
  headline inflation, M2, crude-oil production, PMI, and Inflation Risk.
- Until those inputs are complete, the current implementation publishes the configured
  indicative range of 4.34%-4.50% and retains the internal point estimate for monitoring.
- The empirical 80% model-error interval is an uncertainty diagnostic.

The `YYYY-MM` argument selects the newsletter edition; it does not override the
data-driven GDP target quarter.

## 12. Review after every run

Before distribution:

1. Confirm that the terminal printed both `Nigeria GDP nowcast refreshed` and
   `Economic newsletter refreshed`.
2. Check that workbook and newsletter modification times reflect the current run.
3. Verify the target quarter and publication mode in `Executive Summary`,
   `Quarterly Forecast`, and `Model Engine`.
4. Check PMI, monthly-input coverage, and all charts.
5. Check the newsletter month and year in the title, text, charts, and filename.
6. Review all other economic sections manually. Reusing a previous bulletin as a
   template does not automatically research or refresh every section.
7. Confirm that the master workbook was not changed.

## 13. Command reference

| Purpose | Command from `NEWSLETTER-PROJECT` |
|---|---|
| Check paths | `python ".\GDP_Nowcast_and_Newsletter.py" 2026-10 --dry-run` |
| Run October 2026 | `python ".\GDP_Nowcast_and_Newsletter.py" 2026-10` |
| Run current month | `python ".\GDP_Nowcast_and_Newsletter.py"` |
| Explicit template | `python ".\GDP_Nowcast_and_Newsletter.py" 2026-10 --newsletter-template "C:\path\template.docx"` |
| Explicit plots | `python ".\GDP_Nowcast_and_Newsletter.py" 2026-10 --plots-folder "C:\path\Plots"` |
| Show help | `python ".\GDP_Nowcast_and_Newsletter.py" --help` |

## 14. Troubleshooting

### Python cannot open the launcher

From `NEWSLETTER-PROJECT`, run exactly:

```powershell
Get-Item ".\GDP_Nowcast_and_Newsletter.py"
python ".\GDP_Nowcast_and_Newsletter.py" 2026-10
```

Do not write `GDP\_Nowcast\_and\_Newsletter.py`; backslashes are directory separators.

From inside `Nigeria GDP Nowcasting`, either move up or use the parent path:

```powershell
Set-Location ..
python ".\GDP_Nowcast_and_Newsletter.py" 2026-10
```

```powershell
python "..\GDP_Nowcast_and_Newsletter.py" 2026-10
```

### A Python module is missing

Use the bundled runtime:

```powershell
& "C:\Users\eakan\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" ".\GDP_Nowcast_and_Newsletter.py" 2026-10
```

### No newsletter template is found

Confirm that the previous month's clean report exists, or provide a file with
`--newsletter-template`.

### Permission denied or file-in-use error

Close the master workbook, output workbook, requested newsletter, and any other running
refresh process before retrying.

### Node.js or renderer error

Keep `.automation\node_modules` intact. The launcher prefers bundled Node.js and falls
back to `node` from `PATH`. Do not move individual automation files.

### Workbook verification fails

Check that:

- the monthly sheet is named `monthly data`;
- PMI is in column K;
- Inflation Risk is in column L;
- official GDP is in `Quarterly Data`.

Correct the master, save and close it, run a dry check, and retry.

## 15. Safeguards

- Run only one refresh process at a time.
- Do not edit the master workbook during a run.
- Do not rename the required sheets or move expected columns.
- Do not separate the launcher from the two output folders without updating the code.
- Keep reviewed copies of distributed newsletters.
- Treat generated forecasts and narrative as draft outputs until economic and editorial
  review is complete.

