# Nigeria GDP Nowcast

This is the active CAPE Nigeria GDP nowcasting package. It uses the shared master workbook at:

`C:\Users\eakan\Documents\CAPE-ERC\CAPE Inflation Forecasting Model\Main Data\MultipleFrequencyModelData.xlsx`

The local `Main Data` copy is no longer an input. All other CAPE models can continue to use the same shared workbook.

## One command: model, validation, charts and newsletter

Run from `NEWSLETTER-PROJECT`:

```powershell
python .\GDP_Nowcast_and_Newsletter.py 2026-10 --accept-model-gdp --inflation-month 2026-09 --inflation-point 15.47 --inflation-lower 15.30 --inflation-upper 15.64 --skip-web-research
```

The target quarter is selected automatically from the information date and GDP release calendar. Use `--as-of YYYY-MM-DD` or `--target-quarter YYYYQn` only when reproducing a specific vintage.

## Model-only command

```powershell
python .\Nigeria_GDP_Nowcast\run_nigeria_nowcast.py
```

This runs the preferred eight-model release (Persistence, AR Ridge, Ridge Bridge, MIDAS-style Ridge, Dynamic Factor Proxy, Ridge, Elastic Net and OLS with intercept), writes `Nigeria Quarterly GDP Nowcast.xlsx`, and saves the audit files under `Results`.

The statistical PI50 remains in the technical model results. In the public newsletter, the displayed PI50 lower bound is increased by 10% and the displayed upper bound is reduced by 10%; the model point forecast is unchanged. This is CAPE’s editorial reporting range, not a statistical 50% interval.
