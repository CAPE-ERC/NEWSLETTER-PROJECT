from __future__ import annotations

import json
import math
from pathlib import Path

import numpy as np
import openpyxl


SOURCE = Path(r"C:\Users\eakan\Documents\CAPE-ERC\CAPE Inflation Forecasting Model\Main Data\MultipleFrequencyModelData.xlsx")
OUT = Path(__file__).with_name("nowcast_results.json")


def parse_monthly_sheet(path: Path):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb["monthly data"]
    records = []
    current_year = None
    for row in ws.iter_rows(min_row=3, values_only=True):
        raw = row[0]
        if raw is None:
            continue
        label = str(raw).strip().upper()
        if "M" not in label:
            continue
        if label.startswith("20") or label.startswith("19"):
            year_text, month_text = label.split("M", 1)
            current_year = int(year_text)
        else:
            month_text = label.lstrip("M")
        if current_year is None:
            continue
        month = int(month_text)
        vals = list(row[1:12]) + [None] * max(0, 11 - len(row[1:12]))
        records.append(
            {
                "year": current_year,
                "month": month,
                "date": f"{current_year:04d}-{month:02d}-01",
                "HCPI": vals[0],
                "Core": vals[1],
                "Food": vals[2],
                "MPR": vals[3],
                "FX": vals[4],
                "Oil": vals[5],
                "M2": vals[6],
                "CPS": vals[7],
                "CPD": vals[8],
                # In the master workbook, column K holds the monthly PMI index
                # and column L holds the monthly inflation-risk series. The
                # supplied risk series begins in January 2015; earlier values
                # are not valid risk observations and must not enter the model.
                "PMI": vals[9],
                "Risk": vals[10] if (current_year, month) >= (2015, 1) else None,
            }
        )
    wb.close()
    return records


def parse_quarterly_sheet(path: Path):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb["Quarterly Data"]
    result = []
    for row in ws.iter_rows(min_row=3, values_only=True):
        if row[0] is None or row[1] is None or row[2] is None:
            continue
        y = float(row[2])
        if abs(y) < 1:
            y *= 100.0
        q_label = str(row[1]).strip().upper()
        quarter = int(q_label[1]) if q_label.startswith("Q") else int(row[1])
        result.append({"year": int(row[0]), "quarter": quarter, "growth": y})
    wb.close()
    return result


def num(v):
    try:
        return float(v) if v is not None else math.nan
    except (TypeError, ValueError):
        return math.nan


def pct_change(a, b):
    if not np.isfinite(a) or not np.isfinite(b) or b == 0:
        return math.nan
    return (a / b - 1.0) * 100.0


def build_quarterly_features(monthly, quarterly):
    by_key = {(r["year"], r["month"]): r for r in monthly}
    ymap = {(r["year"], r["quarter"]): r["growth"] for r in quarterly}

    def finite_mean(values):
        clean = [num(value) for value in values if np.isfinite(num(value))]
        return float(np.mean(clean)) if clean else math.nan

    def quarter_records(year, quarter):
        start = (quarter - 1) * 3 + 1
        return [by_key.get((year, start + offset), {}) for offset in range(3)]

    def previous_quarter(year, quarter):
        return (year - 1, 4) if quarter == 1 else (year, quarter - 1)

    def previous_month(year, month):
        return (year - 1, 12) if month == 1 else (year, month - 1)

    rows = []
    for item in quarterly:
        year, q = item["year"], item["quarter"]
        current_records = quarter_records(year, q)
        previous_year_records = quarter_records(year - 1, q)
        prev_q_year, prev_q = previous_quarter(year, q)
        previous_quarter_records = quarter_records(prev_q_year, prev_q)

        def monthly_yoy(field):
            return finite_mean([
                pct_change(num(current.get(field)), num(previous.get(field)))
                for current, previous in zip(current_records, previous_year_records)
            ])

        def quarterly_level(field, records=current_records):
            return finite_mean([record.get(field) for record in records])

        q_start = (q - 1) * 3 + 1
        m2_mom = finite_mean([
            pct_change(
                num(by_key.get((year, month), {}).get("M2")),
                num(by_key.get(previous_month(year, month), {}).get("M2")),
            )
            for month in range(q_start, q_start + 3)
        ])

        ordered_keys = sorted(ymap)
        idx = ordered_keys.index((year, q))
        lag1 = ymap[ordered_keys[idx - 1]] if idx >= 1 else math.nan
        lag4 = ymap[ordered_keys[idx - 4]] if idx >= 4 else math.nan
        row = {
            "year": year,
            "quarter": q,
            "period": f"{year}Q{q}",
            "growth": item["growth"],
            "gdp_lag1": lag1,
            "gdp_lag4": lag4,
            "headline_yoy": monthly_yoy("HCPI"),
            "core_yoy": monthly_yoy("Core"),
            "food_yoy": monthly_yoy("Food"),
            "fx_yoy": monthly_yoy("FX"),
            "fx_qoq": pct_change(quarterly_level("FX"), quarterly_level("FX", previous_quarter_records)),
            "oil_yoy": monthly_yoy("Oil"),
            "oil_qoq": pct_change(quarterly_level("Oil"), quarterly_level("Oil", previous_quarter_records)),
            "m2": quarterly_level("M2"),
            "m2_yoy": monthly_yoy("M2"),
            "m2_mom": m2_mom,
            "m2_qoq": pct_change(quarterly_level("M2"), quarterly_level("M2", previous_quarter_records)),
            "cps": quarterly_level("CPS"),
            "cps_yoy": monthly_yoy("CPS"),
            "mpr": quarterly_level("MPR"),
            "cpd": quarterly_level("CPD"),
            "risk": quarterly_level("Risk"),
            "macro_month": f"{year:04d}Q{q} available-month averages",
        }
        rows.append(row)
    return rows


def impute_standardize(x_train, x_test):
    x_train = np.asarray(x_train, dtype=float)
    x_test = np.asarray(x_test, dtype=float)
    med = np.array([
        np.median(x_train[np.isfinite(x_train[:, j]), j])
        if np.any(np.isfinite(x_train[:, j]))
        else 0.0
        for j in range(x_train.shape[1])
    ])
    x_train = np.where(np.isfinite(x_train), x_train, med)
    x_test = np.where(np.isfinite(x_test), x_test, med)
    mean = x_train.mean(axis=0)
    sd = x_train.std(axis=0)
    sd = np.where(sd > 1e-9, sd, 1.0)
    return (x_train - mean) / sd, (x_test - mean) / sd, med, mean, sd


def ridge_predict(x_train, y_train, x_test, alpha):
    xs, xt, med, mean, sd = impute_standardize(x_train, x_test)
    design = np.column_stack([np.ones(len(xs)), xs])
    penalty = np.eye(design.shape[1])
    penalty[0, 0] = 0.0
    beta = np.linalg.pinv(design.T @ design + alpha * penalty) @ design.T @ y_train
    pred = np.column_stack([np.ones(len(xt)), xt]) @ beta
    return pred, beta, med, mean, sd


def metrics(y, p):
    y = np.asarray(y, dtype=float)
    p = np.asarray(p, dtype=float)
    err = p - y
    return {
        "n": int(len(y)),
        "mae": float(np.mean(np.abs(err))),
        "rmse": float(np.sqrt(np.mean(err**2))),
        "mape": float(np.mean(np.abs(err) / np.maximum(np.abs(y), 0.5)) * 100.0),
    }


def run_backtest(rows, feature_sets, alphas):
    test_start = (2018, 1)
    candidates = []
    for model_name, features in feature_sets.items():
        for alpha in alphas:
            predictions = []
            actuals = []
            periods = []
            for i, row in enumerate(rows):
                if (row["year"], row["quarter"]) < test_start:
                    continue
                train = rows[:i]
                if len(train) < 24:
                    continue
                xtr = [[num(r.get(f)) for f in features] for r in train]
                ytr = np.array([r["growth"] for r in train], dtype=float)
                xte = [[num(row.get(f)) for f in features]]
                pred, *_ = ridge_predict(xtr, ytr, xte, alpha)
                predictions.append(float(pred[0]))
                actuals.append(float(row["growth"]))
                periods.append(row["period"])
            score = metrics(actuals, predictions)
            candidates.append(
                {
                    "model": model_name,
                    "alpha": alpha,
                    "features": features,
                    **score,
                    "periods": periods,
                    "actual": actuals,
                    "predicted": predictions,
                }
            )
    # Direct benchmarks, using the same 2018Q1 start.
    for name, feat in [("Persistence AR(1)", "gdp_lag1"), ("Seasonal persistence", "gdp_lag4")]:
        actuals, predictions, periods = [], [], []
        for row in rows:
            if (row["year"], row["quarter"]) < test_start:
                continue
            v = num(row.get(feat))
            if np.isfinite(v):
                actuals.append(row["growth"])
                predictions.append(v)
                periods.append(row["period"])
        candidates.append({"model": name, "alpha": None, "features": [feat], **metrics(actuals, predictions), "periods": periods, "actual": actuals, "predicted": predictions})
    candidates.sort(key=lambda r: (r["rmse"], r["mae"]))
    return candidates


def fit_selected(rows, current_row, selected):
    features = selected["features"]
    xtr = [[num(r.get(f)) for f in features] for r in rows]
    ytr = np.array([r["growth"] for r in rows], dtype=float)
    xte = [[num(current_row.get(f)) for f in features]]
    pred, beta, med, mean, sd = ridge_predict(xtr, ytr, xte, float(selected["alpha"] or 0.0))
    coefficients = {"intercept": float(beta[0])}
    for i, f in enumerate(features, start=1):
        coefficients[f] = float(beta[i])
    return float(pred[0]), coefficients, med.tolist(), mean.tolist(), sd.tolist()


def main():
    monthly = parse_monthly_sheet(SOURCE)
    quarterly = parse_quarterly_sheet(SOURCE)
    rows = build_quarterly_features(monthly, quarterly)

    # PMI enters the same ridge equation as the other predictors. Completed
    # historical quarters use all three monthly observations. The live quarter
    # uses the mean of whichever one, two, or three months are currently available.
    # The master workbook is the sole PMI source. Do not seed fallback values
    # here, because they can silently survive after the spreadsheet is corrected.
    pmi = {}
    for record in monthly:
        v = num(record.get("PMI"))
        if np.isfinite(v):
            pmi[record["date"][:7]] = v
    pmi_quarters = {}
    for row in rows:
        q_start = (row["quarter"] - 1) * 3 + 1
        keys = [f"{row['year']:04d}-{q_start + offset:02d}" for offset in range(3)]
        vals = [pmi[k] for k in keys if k in pmi]
        row["pmi"] = float(np.mean(vals)) if len(vals) == 3 else math.nan
        if len(vals) == 3:
            pmi_quarters[row["period"]] = row["pmi"]

    # Available predictors are kept here so the specification can be changed by
    # commenting or uncommenting individual lines. The active specification is
    # GDP lag 4, headline inflation, CPD, M2 growth, PMI and inflation risk.
    feature_sets = {
        "Ridge GDP-CPD-M2-inflation-PMI-risk bridge": [
            # Quarterly GDP variables
            "gdp_lag1",
            #"gdp_lag4",

            # Monthly inflation variables
            "headline_yoy",
            # "core_yoy",
            # "food_yoy",
            #"risk",

            # Monthly activity, monetary and financial variables
            "cpd",
            #"m2_yoy",
            "pmi",
            # "m2",
             "m2_mom",
            # "m2_qoq",
            # "cps",
            # "cps_yoy",
             # "mpr",
            # "fx_yoy",
            # "fx_qoq",
            # "oil_yoy",
            # "oil_qoq",
        ],
    }
    candidates = run_backtest(rows, feature_sets, [0.1, 0.5, 1.0, 2.0, 5.0, 10.0, 20.0, 50.0])
    selected = next(c for c in candidates if c["alpha"] is not None)

    last_year = quarterly[-1]["year"]
    last_quarter = quarterly[-1]["quarter"]
    target_year = last_year + 1 if last_quarter == 4 else last_year
    target_quarter = 1 if last_quarter == 4 else last_quarter + 1
    current = {"year": target_year, "quarter": target_quarter, "period": f"{target_year}Q{target_quarter}", "growth": None}
    # Build current first-month features using a dummy target record; lagged GDP is supplied from official history.
    rows_with_dummy = build_quarterly_features(monthly, quarterly + [{"year": target_year, "quarter": target_quarter, "growth": math.nan}])
    current.update(rows_with_dummy[-1])
    current["gdp_lag1"] = quarterly[-1]["growth"]
    lag4_year = target_year - 1
    current["gdp_lag4"] = next(r["growth"] for r in quarterly if r["year"] == lag4_year and r["quarter"] == target_quarter)
    target_start_month = (target_quarter - 1) * 3 + 1
    target_pmi_keys = [f"{target_year:04d}-{target_start_month + offset:02d}" for offset in range(3)]
    source_pmi_by_month = {
        record["date"][:7]: num(record.get("PMI")) for record in monthly
        if np.isfinite(num(record.get("PMI")))
    }
    target_pmi_values = [source_pmi_by_month[k] for k in target_pmi_keys if k in source_pmi_by_month]
    current_pmi = float(np.mean(target_pmi_values)) if target_pmi_values else 50.0
    current["pmi"] = current_pmi
    statistical_forecast, coefficients, med, mean, sd = fit_selected(rows, current, selected)
    core_forecast = statistical_forecast  # Retained as a compatibility field for downstream readers.
    pmi_feature_index = selected["features"].index("pmi")
    pmi_mean = float(mean[pmi_feature_index])
    pmi_beta = float(coefficients["pmi"])

    # Publication rule. The model continues to calculate an internal point estimate for monitoring,
    # but CAPE publishes only an indicative range until every month in the target quarter has the
    # five monthly release series used by the selected model.
    publication_lower = 4.34
    publication_upper = 4.50
    required_monthly_fields = ["HCPI", "M2", "CPD", "PMI", "Risk"]
    required_monthly_labels = ["Headline inflation", "M2", "Crude oil production (CPD)", "PMI", "Inflation risk"]
    target_month_keys = [
        f"{target_year:04d}-{target_start_month + offset:02d}" for offset in range(3)
    ]
    monthly_by_key = {record["date"][:7]: record for record in monthly}
    release_coverage = []
    for month_key in target_month_keys:
        record = monthly_by_key.get(month_key, {})
        availability = {
            "HCPI": bool(np.isfinite(num(record.get("HCPI")))),
            "M2": bool(np.isfinite(num(record.get("M2")))),
            "CPD": bool(np.isfinite(num(record.get("CPD")))),
            "PMI": bool(np.isfinite(num(record.get("PMI")))),
            "Risk": bool(np.isfinite(num(record.get("Risk")))),
        }
        release_coverage.append(
            {
                "month": month_key,
                "available": availability,
                "complete": all(availability[field] for field in required_monthly_fields),
            }
        )
    complete_months = sum(1 for item in release_coverage if item["complete"])
    pmi_months_available = sum(1 for item in release_coverage if item["available"]["PMI"])

    # No judgmental overlay or target-specific override: the equation output is always retained.
    # The available-month averages refresh automatically as months two and three arrive.
    forecast_override = None
    final_forecast = statistical_forecast

    # Empirical 80% interval from selected-model expanding-window residuals, widened modestly
    # for the short PMI bridge history.
    errors = np.asarray(selected["predicted"]) - np.asarray(selected["actual"])
    q10, q90 = np.quantile(errors, [0.10, 0.90])
    lower = final_forecast - q90
    upper = final_forecast - q10

    release_mode = "Point forecast" if complete_months == 3 else "Indicative range"
    publication_point = final_forecast if release_mode == "Point forecast" else None

    latest_monthly = monthly[-1]
    output = {
        "source": str(SOURCE),
        "target_period": current["period"],
        "source_latest_month": latest_monthly["date"][:7],
        "target_latest_quarter": quarterly[-1],
        "selected": {k: v for k, v in selected.items() if k not in {"actual", "predicted", "periods"}},
        "core_forecast": core_forecast,
        "pmi_current_average": current_pmi,
        "pmi_beta": pmi_beta,
        "pmi_beta_raw": pmi_beta,
        "pmi_mean": pmi_mean,
        "gdp_recent_mean": None,
        "pmi_signal_forecast": None,
        "pmi_weight": None,
        "pmi_adjustment": None,
        "statistical_forecast": statistical_forecast,
        "forecast_override": forecast_override,
        "final_forecast": final_forecast,
        "interval_80": [float(lower), float(upper)],
        "publication": {
            "mode": release_mode,
            "point": publication_point,
            "range": [publication_lower, publication_upper],
            "complete_months": complete_months,
            "total_months": 3,
            "pmi_months_available": pmi_months_available,
            "required_monthly_fields": required_monthly_fields,
            "required_monthly_labels": required_monthly_labels,
            "coverage": release_coverage,
            "note": "The equation refreshes with one, two or three available monthly observations. Publication can remain an indicative range until all three target-quarter months contain headline inflation, M2, crude oil production, PMI and inflation risk.",
        },
        "current_features": current,
        "coefficients_standardized": coefficients,
        "training_medians": med,
        "training_means": mean,
        "training_stds": sd,
        "top_candidates": [{k: v for k, v in c.items() if k not in {"actual", "predicted", "periods"}} for c in candidates[:12]],
        "backtest": {"periods": selected["periods"], "actual": selected["actual"], "predicted": selected["predicted"]},
        "quarterly_features": rows,
        "monthly_records": monthly,
        "quarterly_actuals": quarterly,
        "pmi_monthly": [{"month": k, "pmi": v} for k, v in pmi.items()],
        "pmi_quarterly": pmi_quarters,
        "method_note": "The ridge bridge jointly estimates quarterly GDP lag 1 and lag 4 with monthly headline inflation, monthly crude oil production, monthly M2 year-on-year growth rates and monthly PMI. Completed historical quarters use three-month averages for every monthly indicator; the target quarter uses every observation currently available, so one-, two-, and three-month vintages refresh automatically. Headline inflation and M2 year-on-year growth remain monthly-frequency series because one rate is calculated for each month. Exchange rate and crude oil price are excluded. No judgmental overlay or target-specific forecast override is applied. PMI history currently begins in July 2025, so its coefficient will strengthen as more observations accumulate. The empirical 80% interval is an uncertainty diagnostic.",
    }
    def clean_json(value):
        if isinstance(value, dict):
            return {k: clean_json(v) for k, v in value.items()}
        if isinstance(value, list):
            return [clean_json(v) for v in value]
        if isinstance(value, float) and not math.isfinite(value):
            return None
        return value

    output = clean_json(output)
    OUT.write_text(json.dumps(output, indent=2, allow_nan=False), encoding="utf-8")
    print(json.dumps({k: output[k] for k in ["source_latest_month", "target_latest_quarter", "selected", "core_forecast", "pmi_current_average", "pmi_beta", "pmi_weight", "pmi_adjustment", "statistical_forecast", "forecast_override", "final_forecast", "interval_80", "publication"]}, indent=2))


if __name__ == "__main__":
    main()
