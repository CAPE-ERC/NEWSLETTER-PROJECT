from __future__ import annotations

import json
import csv
from datetime import date
from release_aware import ReleaseGuard, end_of_month, snapshot_archive, MONTHLY_SERIES, prospective_registry
import math
import os
import re
from pathlib import Path

import numpy as np
import openpyxl


SOURCE = Path(os.environ.get("NIGERIA_MASTER_WORKBOOK", r"C:\Users\eakan\Documents\CAPE-ERC\CAPE Inflation Forecasting Model\Main Data\MultipleFrequencyModelData.xlsx"))
RELEASE_CSV = Path(__file__).with_name('release_calendar.csv')
CPI_QA_CSV = Path(__file__).with_name('verified_cpi_rates.csv')
CPI_QA_TOLERANCE = 0.06
AS_OF = date.fromisoformat(os.environ.get('NIGERIA_AS_OF_DATE', date.today().isoformat()))
GUARD = ReleaseGuard(RELEASE_CSV)
REPLAY = os.environ.get('NIGERIA_REPLAY_JSON', '').strip()


OUT = Path(__file__).with_name("nowcast_results.json")
MIN_TRAIN_QUARTERS = 20
BACKTEST_HORIZON = 24

def num(value):
    try:
        result = float(value)
        return result if math.isfinite(result) else math.nan
    except (TypeError, ValueError):
        return math.nan


def pct_change(current, previous):
    current, previous = num(current), num(previous)
    if not np.isfinite(current) or not np.isfinite(previous) or previous == 0:
        return math.nan
    return (current / previous - 1.0) * 100.0


def previous_month(year: int, month: int) -> tuple[int, int]:
    return (year - 1, 12) if month == 1 else (year, month - 1)


def previous_quarter(year: int, quarter: int) -> tuple[int, int]:
    return (year - 1, 4) if quarter == 1 else (year, quarter - 1)


def quarter_index(year: int, quarter: int) -> int:
    return year * 4 + quarter


def parse_monthly_sheet(path: Path):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb["monthly data"]
    # Position-based parsing is permitted only with an audited schema.
    expected = {1:"HCPI", 4:"MPR", 7:"M2", 8:"CPS", 9:"CPD", 10:"PMI", 11:"Inflation_Risk_Index"}
    mismatched = {col: (name, ws.cell(2, col+1).value) for col,name in expected.items()
                  if str(ws.cell(2,col+1).value).strip() != name}
    if mismatched:
        wb.close()
        raise RuntimeError(f"Unverified master input headers or outdated money-series definition: {mismatched}. Check the column headings in the shared master workbook.")
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
        records.append({
            "year": current_year, "month": month,
            "date": f"{current_year:04d}-{month:02d}-01",
            "HCPI": vals[0], "Core": vals[1], "Food": vals[2],
            "MPR": vals[3], "FX": vals[4], "Oil": vals[5],
            "M2": vals[6], "CPS": vals[7], "CPD": vals[8],
            "PMI": vals[9],
            "Risk": vals[10] if (current_year, month) >= (2015, 1) else None,
        })
    wb.close()

    override_month = os.environ.get("NIGERIA_PMI_OVERRIDE_MONTH", "").strip()
    override_value = os.environ.get("NIGERIA_PMI_OVERRIDE_VALUE", "").strip()
    if override_month or override_value:
        if not override_month or not override_value:
            raise ValueError("PMI override requires both a month and a value.")
        year_text, month_text = override_month.split("-", 1)
        override_year, override_month_number = int(year_text), int(month_text)
        pmi_value = float(override_value)
        if not 1 <= override_month_number <= 12 or not math.isfinite(pmi_value) or not 0 < pmi_value < 100:
            raise ValueError("PMI override month or value is outside its valid range.")
        record = next(
            (item for item in records if (item["year"], item["month"]) == (override_year, override_month_number)),
            None,
        )
        if record is None:
            record = {
                "year": override_year, "month": override_month_number,
                "date": f"{override_year:04d}-{override_month_number:02d}-01",
                "HCPI": None, "Core": None, "Food": None, "MPR": None,
                "FX": None, "Oil": None, "M2": None, "CPS": None,
                "CPD": None, "PMI": pmi_value, "Risk": None,
            }
            records.append(record)
        else:
            record["PMI"] = pmi_value
        records.sort(key=lambda item: (item["year"], item["month"]))
    return records


def parse_quarterly_sheet(path: Path):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb["Quarterly Data"]
    result = []
    for row in ws.iter_rows(min_row=3, values_only=True):
        if row[0] is None or row[1] is None or row[2] is None:
            continue
        growth = float(row[2])
        if abs(growth) < 1:
            growth *= 100.0
        q_label = str(row[1]).strip().upper()
        quarter = int(q_label[1]) if q_label.startswith("Q") else int(row[1])
        result.append({"year": int(row[0]), "quarter": quarter, "growth": growth})
    wb.close()
    return result


def load_official_headline_yoy(csv_path=CPI_QA_CSV):
    reference = {}
    if csv_path.exists():
        with csv_path.open(newline='', encoding='utf-8-sig') as f:
            for item in csv.DictReader(f):
                if item['series'] == 'HCPI':
                    year, month = item['reference_period'].split('-')
                    reference[(int(year), int(month))] = float(item['official_yoy_pct'])
    return reference


def build_monthly_predictors(monthly, cpi_audit=None):
    raw = {(r['year'], r['month']): r for r in monthly}
    reference = load_official_headline_yoy()
    output = {}
    for record in monthly:
        year, month = record['year'], record['month']
        prior_year = raw.get((year - 1, month), {})
        prior_month = raw.get(previous_month(year, month), {})
        calc_headline = pct_change(record.get('HCPI'), prior_year.get('HCPI'))
        # The archived headline CPI 2024 is retrospectively re-referenced;
        # 2025 headline year-on-year levels reproduce NBS revised rates.
        # For 2026, independently published rates are used to QC the archive.
        verified = reference.get((year, month))
        headline = calc_headline
        if verified is not None:
            difference = calc_headline - verified if np.isfinite(calc_headline) else math.nan
            matches = np.isfinite(difference) and abs(difference) <= CPI_QA_TOLERANCE
            if not matches:
                # Use a documented official *rate*, not an invented CPI level.
                # Raw HCPI in the input workbook is left untouched.
                headline = verified
            if cpi_audit is not None:
                cpi_audit.append({
                    'month': f'{year:04d}-{month:02d}',
                    'archived_hcpi_yoy': calc_headline if np.isfinite(calc_headline) else None,
                    'official_headline_yoy': verified,
                    'action': 'source_rate_matches' if matches else 'use_official_rate_flag_index',
                    'absolute_difference_pp': abs(difference) if np.isfinite(difference) else None,
                })
        elif year >= 2026 and np.isfinite(calc_headline):
            # Fail closed: new source levels must be independently verified.
            headline = math.nan
            if cpi_audit is not None:
                cpi_audit.append({
                    'month': f'{year:04d}-{month:02d}',
                    'archived_hcpi_yoy': calc_headline,
                    'official_headline_yoy': None,
                    'action': 'missing_official_rate_validation_excluded',
                    'absolute_difference_pp': None,
                })
        output[(year, month)] = {
            'headline_yoy': headline,
            # 2024 core/food indices in archive are old-base, while 2025
            # core/food indices are rebased. We do NOT manufacture 2024 levels.
            'core_yoy': math.nan if year == 2025 else pct_change(record.get('Core'), prior_year.get('Core')),
            'food_yoy': math.nan if year == 2025 else pct_change(record.get('Food'), prior_year.get('Food')),
            'fx_mom': pct_change(record.get('FX'), prior_month.get('FX')),
            'oil_mom': pct_change(record.get('Oil'), prior_month.get('Oil')),
            'm2_mom': pct_change(record.get('M2'), prior_month.get('M2')),
            'cps_yoy': pct_change(record.get('CPS'), prior_year.get('CPS')),
            'mpr': num(record.get('MPR')), 'cpd': num(record.get('CPD')),
            'risk': num(record.get('Risk')), 'pmi': num(record.get('PMI')),
        }
    return output

def select_predictors(monthly_predictors, min_nonmissing=24):
    names = next(iter(monthly_predictors.values())).keys()
    return [
        name for name in names
        if sum(np.isfinite(num(row.get(name))) for row in monthly_predictors.values()) >= min_nonmissing
    ]


def quarter_months(year: int, quarter: int):
    start = (quarter - 1) * 3 + 1
    return [(year, start + offset) for offset in range(3)]


def as_of_vintage(year, quarter, vintage):
    return end_of_month(year, (quarter - 1) * 3 + vintage)


def quarter_from_index(index):
    return (index - 1)//4, (index - 1)%4+1


def build_feature_row(year, quarter, vintage, predictors, monthly_predictors, growth_map, growth=None, origin=None):
    origin = origin or as_of_vintage(year, quarter, vintage)
    current_months = quarter_months(year, quarter)
    previous_year, previous_q = previous_quarter(year, quarter)
    previous_months = quarter_months(previous_year, previous_q)
    qidx = quarter_index(year, quarter)
    row = {
        "year": year, "quarter": quarter, "period": f"{year}Q{quarter}",
        "qidx": qidx, "vintage": vintage, "growth": growth,
        "origin_date": origin.isoformat(),
    }
    published = []
    for lag in range(1, 6):
        prev_q_idx = qidx - lag
        gy, gq = quarter_from_index(prev_q_idx)
        value = growth_map.get(prev_q_idx, math.nan)
        available = GUARD.gdp_available(gy, gq, origin)
        row[f"gdp_lag{lag}"] = value if available else math.nan
        if available and np.isfinite(num(value)):
            published.append((lag, value))
    row["gdp_latest_published"] = float(published[0][1]) if published else math.nan
    row["gdp_age_quarters"] = float(published[0][0]) if published else math.nan

    for variable in predictors:
        current_values = []
        for position, month_key in enumerate(current_months, start=1):
            value = GUARD.feature_value(monthly_predictors, variable, month_key, origin)
            row[f"{variable}_m{position}"] = value if position <= vintage else math.nan
            if position <= vintage and np.isfinite(value):
                current_values.append(value)
        previous_values = [GUARD.feature_value(monthly_predictors, variable, key, origin)
                           for key in previous_months]
        previous_values = [value for value in previous_values if np.isfinite(value)]
        row[f"{variable}_qtd_mean"] = float(np.mean(current_values)) if current_values else math.nan
        row[f"{variable}_qtd_last"] = current_values[-1] if current_values else math.nan
        row[f"{variable}_qtd_change"] = (
            current_values[-1] - current_values[0] if len(current_values) >= 2
            else math.nan
        )
        row[f"{variable}_prev_q_mean"] = float(np.mean(previous_values)) if previous_values else math.nan
    row['cps_yoy_trend'] = GUARD.credit_momentum(monthly_predictors, origin)
    return row


