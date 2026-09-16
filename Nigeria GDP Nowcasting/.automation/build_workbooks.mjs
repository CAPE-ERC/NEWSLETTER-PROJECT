import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const workDir = path.dirname(fileURLToPath(import.meta.url));
const outDir = process.env.NIGERIA_NOWCAST_OUTPUT_DIR
  ? path.resolve(process.env.NIGERIA_NOWCAST_OUTPUT_DIR)
  : path.resolve("outputs", "nigeria_q3_gdp_nowcast");
const data = JSON.parse(await fs.readFile(path.join(workDir, "nowcast_results.json"), "utf8"));
const sourcePath = data.source;
const stagedSourcePath = null;
const forecastPath = path.join(outDir, "Nigeria Quarterly GDP Nowcast.xlsx");
const previewPath = path.join(outDir, "Nigeria Quarterly GDP Nowcast Preview.png");
const forecastPreviewPath = path.join(outDir, "Nigeria Quarterly GDP Forecast Preview.png");
const enginePreviewPath = path.join(outDir, "Nigeria GDP Model Engine Preview.png");
const methodologyPreviewPath = path.join(outDir, "Nigeria GDP Methodology Preview.png");
const monthlyPreviewPath = path.join(outDir, "Nigeria GDP Monthly Inputs Preview.png");
await fs.mkdir(outDir, { recursive: true });

const C = {
  navy: "#17365D",
  blue: "#1F4E78",
  teal: "#0F6B67",
  gold: "#D9A441",
  lightBlue: "#DCE6F1",
  lightTeal: "#DDEBF7",
  paleGold: "#FFF2CC",
  green: "#E2F0D9",
  red: "#FCE4D6",
  gray: "#F2F2F2",
  grid: "#D9E1F2",
  text: "#1F2937",
  white: "#FFFFFF",
};
const FONT = "Aptos";

function safeNumber(v) {
  return typeof v === "number" && Number.isFinite(v) ? v : null;
}

function excelCol(n) {
  let s = "";
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function applyTitle(sheet, range, title, subtitle = null) {
  sheet.mergeCells(range);
  const cell = sheet.getRange(range.split(":")[0]);
  cell.values = [[title]];
  sheet.getRange(range).format = {
    fill: C.navy,
    font: { name: FONT, size: 18, bold: true, color: C.white },
    verticalAlignment: "center",
  };
  sheet.getRange(range).format.rowHeight = 34;
  if (subtitle) {
    const cols = range.split(":");
    const start = cols[0].replace(/\d+/g, "") + "2";
    const end = cols[1].replace(/\d+/g, "") + "2";
    sheet.mergeCells(`${start}:${end}`);
    sheet.getRange(start).values = [[subtitle]];
    sheet.getRange(`${start}:${end}`).format = {
      fill: C.lightBlue,
      font: { name: FONT, size: 10, italic: true, color: C.text },
      wrapText: true,
      verticalAlignment: "center",
    };
    sheet.getRange(`${start}:${end}`).format.rowHeight = 28;
  }
}

function applyHeader(range) {
  range.format = {
    fill: C.blue,
    font: { name: FONT, size: 10, bold: true, color: C.white },
    wrapText: true,
    verticalAlignment: "center",
    borders: { preset: "all", style: "thin", color: C.grid },
  };
  range.format.rowHeight = 28;
}

function applyBody(range) {
  range.format = {
    font: { name: FONT, size: 10, color: C.text },
    verticalAlignment: "center",
    borders: { preset: "all", style: "thin", color: C.grid },
  };
}

function setWidths(sheet, widths) {
  for (const [col, width] of Object.entries(widths)) {
    sheet.getRange(`${col}:${col}`).format.columnWidth = width;
  }
}

function formatDataSheet(sheet, usedRange, headerRange) {
  sheet.showGridLines = false;
  sheet.freezePanes.freezeRows(Number(headerRange.match(/\d+/)?.[0] || 1));
  applyBody(sheet.getRange(usedRange));
  applyHeader(sheet.getRange(headerRange));
}

const pmiByMonth = new Map(data.pmi_monthly.map((r) => [r.month, r.pmi]));

// ---------------------------------------------------------------------------
// Build the Nigeria-only quarterly nowcast workbook. The CAPE master workbook
// is read only by analyze_nowcast.py and is never opened for writing here.
// ---------------------------------------------------------------------------
const wb = Workbook.create();
const summary = wb.worksheets.add("Executive Summary");
const forecast = wb.worksheets.add("Quarterly Forecast");
const comparison = wb.worksheets.add("Model Comparison");
const backtest = wb.worksheets.add("Backtest");
const engine = wb.worksheets.add("Model Engine");
const monthly = wb.worksheets.add("Monthly Inputs");
const features = wb.worksheets.add("Quarterly Features");
const methods = wb.worksheets.add("Methodology");
const sources = wb.worksheets.add("Sources");

// Monthly Inputs first so downstream formula references exist.
monthly.getRange("A1:M1").values = [["Month", "HCPI", "Core CPI", "Food CPI", "MPR", "BDC FX", "Crude oil", "M2", "CPS", "CPD", "Inflation risk", "PMI", "PMI status"]];
const monthlyRows = [];
const monthRow = new Map();
for (const r of data.monthly_records) {
  const key = r.date.slice(0, 7);
  const pmiValue = pmiByMonth.get(key);
  monthlyRows.push([
    new Date(`${r.date}T00:00:00Z`), safeNumber(r.HCPI), safeNumber(r.Core), safeNumber(r.Food), safeNumber(r.MPR), safeNumber(r.FX), safeNumber(r.Oil), safeNumber(r.M2), safeNumber(r.CPS), safeNumber(r.CPD), safeNumber(r.Risk), safeNumber(pmiValue), pmiValue == null ? "Not available" : (r.PMI == null ? "Seeded from Stanbic/CAPE history" : "Pulled from master workbook"),
  ]);
  monthRow.set(key, monthlyRows.length + 1);
}
monthly.getRange(`A2:M${monthlyRows.length + 1}`).values = monthlyRows;
formatDataSheet(monthly, `A1:M${monthlyRows.length + 1}`, "A1:M1");
monthly.getRange(`A2:A${monthlyRows.length + 1}`).format.numberFormat = "mmm-yyyy";
monthly.getRange(`B2:L${monthlyRows.length + 1}`).format.numberFormat = "0.00";
monthly.getRange(`H2:I${monthlyRows.length + 1}`).format.numberFormat = "#,##0";
setWidths(monthly, { A: 13, B: 11, C: 11, D: 11, E: 10, F: 13, G: 12, H: 15, I: 15, J: 10, K: 14, L: 10, M: 32 });
monthly.tabColor = C.teal;

// Model Engine: auditable joint ridge equation and release gate.
applyTitle(engine, "A1:H1", "Selected Model Engine", `${data.selected.model}, ridge alpha ${data.selected.alpha}; current macro vintage ${data.current_features.macro_month}.`);
engine.getRange("A3:F3").values = [["Term", "Current/raw", "Training mean", "Training std. dev.", "Standardized coefficient", "Contribution"]];
const featureRows = [["Intercept", null, null, null, data.coefficients_standardized.intercept, null]];
for (let i = 0; i < data.selected.features.length; i++) {
  const f = data.selected.features[i];
  featureRows.push([f, safeNumber(data.current_features[f]), data.training_means[i], data.training_stds[i], data.coefficients_standardized[f], null]);
}
engine.getRange(`A4:F${featureRows.length + 3}`).values = featureRows;
engine.getRange("F4").formulas = [["=E4"]];
for (let i = 1; i < featureRows.length; i++) {
  const r = i + 4;
  engine.getRange(`F${r}`).formulas = [[`=E${r}*((B${r}-C${r})/D${r})`]];
}
applyHeader(engine.getRange("A3:F3"));
applyBody(engine.getRange(`A4:F${featureRows.length + 3}`));
engine.getRange(`B4:F${featureRows.length + 3}`).format.numberFormat = "0.000";

engine.getRange("G3:H3").values = [["Forecast component", "Value"]];
applyHeader(engine.getRange("G3:H3"));
const targetStartMonth = (data.current_features.quarter - 1) * 3 + 1;
const pmiKey1 = `${data.current_features.year}-${String(targetStartMonth).padStart(2, "0")}`;
const pmiKey2 = `${data.current_features.year}-${String(targetStartMonth + 1).padStart(2, "0")}`;
const pmiKey3 = `${data.current_features.year}-${String(targetStartMonth + 2).padStart(2, "0")}`;
const pmiCells = [pmiKey1, pmiKey2, pmiKey3]
  .map((key) => monthRow.get(key))
  .filter((row) => row != null)
  .map((row) => `'Monthly Inputs'!L${row}`);
const lowerWidth = data.final_forecast - data.interval_80[0];
const upperWidth = data.interval_80[1] - data.final_forecast;
engine.getRange("G4:G22").values = [["Joint ridge-equation forecast"], ["Current PMI (available-month avg.)"], ["PMI training mean"], ["PMI standardized coefficient"], ["PMI included in equation"], ["Statistical estimate before override"], ["PMI treatment"], ["CAPE monitoring estimate"], ["Separate PMI overlay"], ["Lower model-error width"], ["Upper model-error width"], ["Backtest RMSE"], ["Published range lower"], ["Published range upper"], ["Complete target-quarter months"], ["PMI months available"], ["Release mode"], ["Required monthly inputs"], ["Complete historical PMI quarters"]];
engine.getRange("H4").formulas = [[`=SUM(F4:F${featureRows.length + 3})`]];
if (pmiCells.length > 0) {
  engine.getRange("H5").formulas = [[`=AVERAGE(${pmiCells.join(",")})`]];
} else {
  engine.getRange("H5").values = [[data.pmi_current_average]];
}
engine.getRange("H6:H10").values = [[data.pmi_mean], [data.pmi_beta], ["Yes"], [data.statistical_forecast], ["Joint regression coefficient"]];
engine.getRange("H11").values = [[data.final_forecast]];
engine.getRange("H12").values = [["Not used"]];
engine.getRange("H13:H15").values = [[lowerWidth], [upperWidth], [data.selected.rmse]];
engine.getRange("H16:H21").values = [[data.publication.range[0]], [data.publication.range[1]], [data.publication.complete_months], [data.publication.pmi_months_available], [data.publication.mode], [data.publication.required_monthly_labels.join(", ")]];
engine.getRange("H22").values = [[Object.keys(data.pmi_quarterly).length]];
applyBody(engine.getRange("G4:H22"));
engine.getRange("H4:H7").format.numberFormat = "0.000";
engine.getRange("H9").format.numberFormat = "0.000";
engine.getRange("H11:H17").format.numberFormat = "0.000";
engine.getRange("H18:H19").format.numberFormat = "0";
engine.getRange("G11:H11").format = { fill: C.green, font: { name: FONT, size: 11, bold: true, color: C.navy }, borders: { preset: "all", style: "medium", color: C.teal } };
engine.getRange("G16:H20").format = { fill: C.paleGold, font: { name: FONT, size: 10, bold: true, color: C.navy }, borders: { preset: "all", style: "thin", color: C.gold } };
engine.getRange("H21").format.wrapText = true;
engine.getRange("H10").format.wrapText = true;
engine.getRange("G10:H10").format.rowHeight = 34;
engine.getRange("G21:H21").format.rowHeight = 54;
engine.getRange("H22").format.numberFormat = "0";
engine.getRange("G22:H22").format.rowHeight = 24;
setWidths(engine, { A: 24, B: 15, C: 15, D: 15, E: 22, F: 15, G: 34, H: 22 });
engine.freezePanes.freezeRows(3);
engine.showGridLines = false;
engine.tabColor = C.gold;

// Quarterly Forecast.
applyTitle(forecast, "A1:I1", "Nigeria Quarterly Real GDP Growth Forecast", "Quarterly GDP lags are combined with available-month averages of inflation, CPD, monthly M2 growth and PMI.");
forecast.getRange("A3:I3").values = [["Quarter", "Actual YoY growth (%)", "Joint model (%)", "PMI average", "PMI in equation", "Published point (%)", "Published lower (%)", "Published upper (%)", "Release status"]];
const qRows = data.quarterly_actuals.map((q) => [`${q.year}Q${q.quarter}`, q.growth, null, null, null, null, null, null, "Official actual"]);
const forecastRow = qRows.length + 4;
qRows.push([data.target_period, null, null, null, null, null, null, null, data.publication.mode]);
forecast.getRange(`A4:I${forecastRow}`).values = qRows;
forecast.getRange(`C${forecastRow}`).formulas = [["='Model Engine'!H4"]];
forecast.getRange(`D${forecastRow}`).formulas = [["='Model Engine'!H5"]];
forecast.getRange(`E${forecastRow}`).values = [["Yes"]];
if (data.publication.mode === "Point forecast") {
  forecast.getRange(`F${forecastRow}`).formulas = [["='Model Engine'!H11"]];
} else {
  forecast.getRange(`G${forecastRow}`).formulas = [["='Model Engine'!H16"]];
  forecast.getRange(`H${forecastRow}`).formulas = [["='Model Engine'!H17"]];
}
formatDataSheet(forecast, `A3:I${forecastRow}`, "A3:I3");
forecast.getRange(`B4:D${forecastRow}`).format.numberFormat = "0.00";
forecast.getRange(`E4:E${forecastRow}`).format.numberFormat = "General";
forecast.getRange(`F4:H${forecastRow}`).format.numberFormat = "0.00";
forecast.getRange(`A${forecastRow}:I${forecastRow}`).format = { fill: C.paleGold, font: { name: FONT, size: 10, bold: true, color: C.navy }, borders: { preset: "all", style: "medium", color: C.gold } };
setWidths(forecast, { A: 12, B: 19, C: 15, D: 15, E: 12, F: 18, G: 17, H: 17, I: 24 });
forecast.tabColor = C.gold;

// Chart helper focuses on recent quarters and the nowcast.
const recentStartIndex = Math.max(0, data.quarterly_actuals.length - 18);
const chartStart = recentStartIndex + 4;
const qChart = forecast.charts.add("line", forecast.getRange(`A3:B${forecastRow}`));
qChart.title = "Official real GDP growth history";
qChart.hasLegend = true;
qChart.titleTextStyle.typeface = FONT;
qChart.xAxis = { axisType: "textAxis", textStyle: { typeface: FONT } };
qChart.yAxis = { numberFormatCode: "0.0", numberFormatSourceLinked: false, textStyle: { typeface: FONT } };
qChart.setPosition("K3", "S20");

// Model comparison.
applyTitle(comparison, "A1:G1", "Out-of-Sample Model Comparison", "Expanding-window validation begins in 2018Q1. Lowest RMSE determines the selected model; MAPE is secondary because GDP growth can approach zero.");
comparison.getRange("A3:G3").values = [["Rank", "Model", "Ridge alpha", "RMSE (pp)", "MAE (pp)", "MAPE (%)", "Validation quarters"]];
const compRows = data.top_candidates.map((c, i) => [i + 1, c.model, c.alpha, c.rmse, c.mae, c.mape, c.n]);
comparison.getRange(`A4:G${compRows.length + 3}`).values = compRows;
formatDataSheet(comparison, `A3:G${compRows.length + 3}`, "A3:G3");
comparison.getRange(`C4:F${compRows.length + 3}`).format.numberFormat = "0.00";
comparison.getRange("A4:G4").format = { fill: C.green, font: { name: FONT, size: 10, bold: true, color: C.navy }, borders: { preset: "all", style: "medium", color: C.teal } };
setWidths(comparison, { A: 9, B: 26, C: 13, D: 14, E: 14, F: 13, G: 19 });
comparison.tabColor = C.blue;

// Backtest detail.
applyTitle(backtest, "A1:E1", "Selected-Model Backtest", "Each prediction uses only information available before the predicted quarter.");
backtest.getRange("A3:E3").values = [["Quarter", "Actual (%)", "Predicted (%)", "Error (pp)", "Absolute error (pp)"]];
const btRows = data.backtest.periods.map((p, i) => [p, data.backtest.actual[i], data.backtest.predicted[i], null, null]);
backtest.getRange(`A4:E${btRows.length + 3}`).values = btRows;
for (let i = 0; i < btRows.length; i++) {
  const r = i + 4;
  backtest.getRange(`D${r}:E${r}`).formulas = [[`=C${r}-B${r}`, `=ABS(D${r})`]];
}
formatDataSheet(backtest, `A3:E${btRows.length + 3}`, "A3:E3");
backtest.getRange(`B4:E${btRows.length + 3}`).format.numberFormat = "0.00";
setWidths(backtest, { A: 12, B: 14, C: 15, D: 14, E: 20 });
const btChart = backtest.charts.add("line", backtest.getRange(`A3:C${btRows.length + 3}`));
btChart.title = "Actual vs. expanding-window prediction";
btChart.hasLegend = true;
btChart.titleTextStyle.typeface = FONT;
btChart.xAxis = { axisType: "textAxis", textStyle: { typeface: FONT } };
btChart.yAxis = { numberFormatCode: "0.0", numberFormatSourceLinked: false, textStyle: { typeface: FONT } };
btChart.setPosition("G3", "O20");
backtest.tabColor = C.blue;

// Quarterly model features.
const featureNames = ["period", "growth", "gdp_lag1", "gdp_lag4", "headline_yoy", "core_yoy", "food_yoy", "fx_yoy", "fx_qoq", "oil_yoy", "oil_qoq", "m2", "m2_yoy", "cps", "cps_yoy", "mpr", "cpd", "risk", "pmi", "macro_month"];
features.getRange(`A1:${excelCol(featureNames.length)}1`).values = [featureNames.map((x) => x.replaceAll("_", " "))];
const featureData = data.quarterly_features.map((r) => featureNames.map((f) => safeNumber(r[f]) ?? r[f] ?? null));
features.getRange(`A2:${excelCol(featureNames.length)}${featureData.length + 1}`).values = featureData;
formatDataSheet(features, `A1:${excelCol(featureNames.length)}${featureData.length + 1}`, `A1:${excelCol(featureNames.length)}1`);
features.getRange(`B2:P${featureData.length + 1}`).format.numberFormat = "0.00";
setWidths(features, { A: 12, B: 12, C: 12, D: 12, E: 15, F: 13, G: 13, H: 12, I: 12, J: 12, K: 12, L: 12, M: 12, N: 10, O: 10, P: 12, Q: 14 });
features.tabColor = C.teal;

// Methodology and sources.
applyTitle(methods, "A1:C1", "Methodology and Refresh Logic", "Nigeria-only extraction adapted from the CAPE-ERC-AFDB bridge-model framework.");
const methodRows = [
  ["Target", `${data.target_period} real GDP growth, year-on-year (%)`, "The next quarter after the latest official observation in the master workbook."],
  ["Current vintage", `Monthly indicators: ${data.current_features.macro_month}; PMI available: ${data.publication.pmi_months_available} of 3 months`, "Each monthly indicator uses every observation currently available in the target quarter."],
  ["Model specification", data.selected.features.join(", "), "The displayed specification is generated directly from the active feature list in analyze_nowcast.py."],
  ["Selection rule", `Lowest expanding-window RMSE from 2018Q1: ${data.selected.model}, alpha ${data.selected.alpha}`, `RMSE ${data.selected.rmse.toFixed(2)} pp; MAE ${data.selected.mae.toFixed(2)} pp across ${data.selected.n} quarters.`],
  ["Selected variables", `${data.selected.features.length} jointly estimated predictors`, `Equation variables: ${data.selected.features.join(", ")}.`],
  ["Monthly treatment", `Available-month inputs are used for the active monthly predictors; current PMI average ${data.pmi_current_average.toFixed(1)} from ${data.publication.pmi_months_available} of 3 months`, "The target quarter uses one, two or three monthly observations as available and automatically incorporates the final month. Completed historical quarters use all three."],
  ["Current monitoring estimate", `${data.final_forecast.toFixed(2)}% for ${data.target_period}`, data.forecast_override == null ? "Generated by the statistical model." : `CAPE target-specific override; underlying statistical estimate ${data.statistical_forecast.toFixed(2)}%.`],
  ["Publication rule", `${data.publication.complete_months} of ${data.publication.total_months} target-quarter months complete; current release: ${data.publication.mode}`, `A point forecast is published only when ${data.publication.required_monthly_labels.join(", ")} are present for all three months. Until then, publish ${data.publication.range[0].toFixed(2)}-${data.publication.range[1].toFixed(2)}%.`],
  ["Internal uncertainty", "Empirical 80% model-error interval retained in the engine", "The 80% figure describes interval coverage. It is not the PMI weight and is not the public range."],
  ["Refresh", "Update MultipleFrequencyModelData.xlsx, then run Refresh Nigeria GDP Nowcast.ps1", `The script re-reads ${data.publication.required_monthly_labels.join(", ")} and official quarterly GDP without modifying the master workbook.`],
  ["Master workbook fields", "PMI is read from column K and inflation risk from column L; official GDP is read from Quarterly Data", "The master workbook is read-only during refresh. Generated outputs are written only to the Nigeria GDP Nowcasting folder."],
];
methods.getRange("A3:C3").values = [["Item", "Implementation", "Rationale / audit note"]];
methods.getRange(`A4:C${methodRows.length + 3}`).values = methodRows;
formatDataSheet(methods, `A3:C${methodRows.length + 3}`, "A3:C3");
methods.getRange(`A4:C${methodRows.length + 3}`).format.wrapText = true;
methods.getRange(`A4:C${methodRows.length + 3}`).format.rowHeight = 46;
setWidths(methods, { A: 24, B: 54, C: 64 });
methods.tabColor = C.blue;

applyTitle(sources, "A1:C1", "Sources and Provenance", "Paths and web references used for this nowcast.");
sources.getRange("A3:C3").values = [["Source", "Location", "Use"]];
const sourceRows = [
  ["CAPE master macro workbook", sourcePath, "Monthly headline CPI and M2 levels used to calculate monthly year-on-year rates, monthly crude oil production, monthly PMI and official quarterly GDP."],
  ["CAPE-ERC-AFDB nowcasting system", "C:\\Users\\eakan\\Documents\\CAPE-ERC-AFDB\\quarterly_nowcast_system.py", "Model architecture adapted to a Nigeria-only quarterly growth target."],
  ["Generated", new Date().toISOString(), "Workbook creation timestamp (UTC)."],
];
sources.getRange(`A4:C${sourceRows.length + 3}`).values = sourceRows;
formatDataSheet(sources, `A3:C${sourceRows.length + 3}`, "A3:C3");
sources.getRange(`A4:C${sourceRows.length + 3}`).format.wrapText = true;
sources.getRange(`A4:C${sourceRows.length + 3}`).format.rowHeight = 40;
setWidths(sources, { A: 29, B: 90, C: 60 });
sources.tabColor = C.blue;

// Executive Summary last, with formula links to audited sheets.
applyTitle(summary, "A1:H1", `Nigeria ${data.target_period} GDP Nowcast`, "CAPE Economic Research & Consulting | Nigeria-only quarterly nowcast | Year-on-year real GDP growth");
summary.getRange("A4:H4").values = [["Published range", null, "to", null, "Point forecast", data.publication.point == null ? "Pending" : null, "Release mode", data.publication.mode]];
if (data.publication.mode === "Indicative range") {
  summary.getRange("B4").formulas = [[`='Quarterly Forecast'!G${forecastRow}`]];
  summary.getRange("D4").formulas = [[`='Quarterly Forecast'!H${forecastRow}`]];
} else {
  summary.getRange("F4").formulas = [[`='Quarterly Forecast'!F${forecastRow}`]];
}
summary.getRange("A4:H4").format = { fill: C.green, font: { name: FONT, size: 11, bold: true, color: C.navy }, borders: { preset: "all", style: "medium", color: C.teal }, horizontalAlignment: "center", verticalAlignment: "center" };
summary.getRange("A4:H4").format.rowHeight = 38;
summary.getRange("B4:D4").format.numberFormat = "0.00";
summary.getRange("F4").format.numberFormat = "0.00";
summary.getRange("A6:H6").values = [["Latest official GDP", "Quarter", "Current PMI", "Complete months", "Selected model", "RMSE", "MAE", "PMI treatment"]];
summary.getRange("A7:H7").values = [[data.target_latest_quarter.growth, `${data.target_latest_quarter.year}Q${data.target_latest_quarter.quarter}`, null, data.publication.complete_months, data.selected.model, data.selected.rmse, data.selected.mae, "Inside equation"]];
summary.getRange("C7").formulas = [["='Model Engine'!H5"]];
applyHeader(summary.getRange("A6:H6"));
applyBody(summary.getRange("A7:H7"));
summary.getRange("E7").format.wrapText = true;
summary.getRange("A7:H7").format.rowHeight = 30;
summary.getRange("A7:G7").format.numberFormat = "0.00";
summary.getRange("D7").format.numberFormat = "0";
summary.getRange("H7").format.numberFormat = "General";
summary.mergeCells("A10:H12");
const summaryInterpretation = data.publication.mode === "Indicative range"
  ? `Interpretation: CAPE publishes an indicative ${data.publication.range[0].toFixed(2)}-${data.publication.range[1].toFixed(2)}% range for ${data.target_period} because only ${data.publication.complete_months} of ${data.publication.total_months} months has all required inputs. The current ${data.final_forecast.toFixed(2)}% estimate remains a monitoring output and is not the published point forecast. The available PMI average of ${data.pmi_current_average.toFixed(1)} remains above 50, signalling private-sector expansion.`
  : `Interpretation: all three target-quarter months are complete, so CAPE publishes the ${data.final_forecast.toFixed(2)}% forecast for ${data.target_period}. Active equation variables: ${data.selected.features.join(", ")}.`;
summary.getRange("A10").values = [[summaryInterpretation]];
summary.getRange("A10:H12").format = { fill: C.lightBlue, font: { name: FONT, size: 11, color: C.text }, wrapText: true, verticalAlignment: "center", borders: { preset: "outside", style: "medium", color: C.blue } };
summary.getRange("A10:H12").format.rowHeight = 26;
summary.mergeCells("A14:H15");
summary.getRange("A14").values = [["Refresh note: update the master workbook inputs and run the included refresh script. The workflow reads the master without modifying it, publishes the 4.34-4.50% range while target-quarter inputs are incomplete, and switches to the model point forecast after all three months are complete."]];
summary.getRange("A14:H15").format = { fill: C.paleGold, font: { name: FONT, size: 10, italic: true, color: C.text }, wrapText: true, verticalAlignment: "center", borders: { preset: "outside", style: "thin", color: C.gold } };
setWidths(summary, { A: 20, B: 16, C: 17, D: 17, E: 30, F: 12, G: 12, H: 14 });
summary.showGridLines = false;
summary.tabColor = C.gold;

// Recent-history chart on dashboard uses a dedicated helper block.
summary.getRange("J2:M2").values = [["Quarter", "Actual", "Range lower", "Range upper"]];
const recentActuals = data.quarterly_actuals.slice(recentStartIndex);
const chartRows = recentActuals.map((q) => [`${q.year}Q${q.quarter}`, q.growth, null, null]);
if (chartRows.length) {
  chartRows[chartRows.length - 1][2] = chartRows[chartRows.length - 1][1];
  chartRows[chartRows.length - 1][3] = chartRows[chartRows.length - 1][1];
}
if (data.publication.mode === "Indicative range") {
  chartRows.push([data.target_period, null, data.publication.range[0], data.publication.range[1]]);
} else {
  chartRows.push([data.target_period, null, data.final_forecast, data.final_forecast]);
}
summary.getRange(`J3:M${chartRows.length + 2}`).values = chartRows;
const summaryChart = summary.charts.add("line", summary.getRange(`J2:M${chartRows.length + 2}`));
summaryChart.title = data.publication.mode === "Indicative range" ? "Nigeria real GDP growth: actuals and indicative range" : "Nigeria real GDP growth: actuals and point forecast";
summaryChart.hasLegend = true;
summaryChart.titleTextStyle.typeface = FONT;
summaryChart.xAxis = { axisType: "textAxis", textStyle: { typeface: FONT } };
summaryChart.yAxis = { numberFormatCode: "0.0", numberFormatSourceLinked: false, textStyle: { typeface: FONT } };
summaryChart.setPosition("A17", "H34");
summary.getRange(`J2:M${chartRows.length + 2}`).format.font = { name: FONT, size: 9, color: "#808080" };
summary.getRange("J:M").format.columnWidth = 12;

// Inspect, render, and export.
const summaryInspect = await wb.inspect({ kind: "region", sheetId: "Executive Summary", range: "A1:H15", maxChars: 6000 });
console.log(summaryInspect.ndjson || summaryInspect);
const errors = await wb.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A", options: { useRegex: true, maxResults: 100 }, maxChars: 4000 });
console.log(errors.ndjson || errors);
const preview = await wb.render({ sheetName: "Executive Summary", range: "A1:H34", scale: 1.2, format: "png" });
await fs.writeFile(previewPath, new Uint8Array(await preview.arrayBuffer()));
const forecastPreview = await wb.render({ sheetName: "Quarterly Forecast", range: `A1:I${forecastRow}`, scale: 1.1, format: "png" });
await fs.writeFile(forecastPreviewPath, new Uint8Array(await forecastPreview.arrayBuffer()));
const enginePreview = await wb.render({ sheetName: "Model Engine", range: "A1:H22", scale: 1.2, format: "png" });
await fs.writeFile(enginePreviewPath, new Uint8Array(await enginePreview.arrayBuffer()));
const methodologyPreview = await wb.render({ sheetName: "Methodology", range: "A1:C14", scale: 1.2, format: "png" });
await fs.writeFile(methodologyPreviewPath, new Uint8Array(await methodologyPreview.arrayBuffer()));
const monthlyPreview = await wb.render({ sheetName: "Monthly Inputs", range: "A172:M188", scale: 1.2, format: "png" });
await fs.writeFile(monthlyPreviewPath, new Uint8Array(await monthlyPreview.arrayBuffer()));
const out = await SpreadsheetFile.exportXlsx(wb);
await out.save(forecastPath);

console.log(JSON.stringify({ forecastPath, stagedSourcePath, previewPath, forecastPreviewPath, enginePreviewPath, methodologyPreviewPath, monthlyPreviewPath, forecastRow, internalPoint: data.final_forecast, publication: data.publication }, null, 2));
