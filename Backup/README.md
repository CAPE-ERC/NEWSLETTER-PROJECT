# CAPE GDP Nowcast and Newsletter

The active Nigeria GDP nowcasting project is:

`NEWSLETTER-PROJECT\Nigeria_GDP_Nowcast`

It reads the shared master workbook used by the other CAPE projects:

`C:\Users\eakan\Documents\CAPE-ERC\CAPE Inflation Forecasting Model\Main Data\MultipleFrequencyModelData.xlsx`

## Run the complete workflow

From `NEWSLETTER-PROJECT`:

```powershell
python .\GDP_Nowcast_and_Newsletter.py
```

From inside `Nigeria_GDP_Nowcast`, the same command works:

```powershell
python .\GDP_Nowcast_and_Newsletter.py
```

The launcher:

1. checks the dated PMI input;
2. selects the nowcast quarter from the information date and GDP release calendar;
3. runs and validates the preferred eight-model ensemble;
4. rebuilds `Nigeria Quarterly GDP Nowcast.xlsx`;
5. generates the GDP, PMI and inflation charts; and
6. updates the CAPE Economic Performance and Prospect Bulletin.

The model retains its statistical PI50 for audit. The public bulletin reports the unchanged model point forecast and CAPE’s expected range, calculated by increasing the displayed PI50 lower bound by 10% and reducing its displayed upper bound by 10%. The launcher command remains unchanged.

## Model-only run

From `Nigeria_GDP_Nowcast`:

```powershell
python .\run_nigeria_nowcast.py
python .\check_package.py --require-excel
```

Use `--as-of YYYY-MM-DD` or `--target-quarter YYYYQn` only when reproducing a particular information vintage.

## Active outputs

- `Nigeria_GDP_Nowcast\Nigeria Quarterly GDP Nowcast.xlsx`
- `Nigeria_GDP_Nowcast\Results\forecast.json`
- `Nigeria_GDP_Nowcast\Newsletter Assets\`
- `CAPE ERC Economic Newsletter\YYYY\Month Newsletter\CAPE Economic Performance and Prospect Bulletin Month YYYY.docx`

The master workbook is read and integrity-checked during model execution; it is not modified by the nowcast runner.
