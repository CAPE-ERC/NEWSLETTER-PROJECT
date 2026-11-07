"""Publication-timing guards for the Nigeria GDP nowcasting engine.

Release-date fields are *not* historical revision vintages. Explicit source dates
supersede conservative proxy lags. Predictions remain final-vintage proxies until
actual source-vintage values are collected prospectively.
"""
from __future__ import annotations

import calendar
import csv
from datetime import date, timedelta
from pathlib import Path
import math

MONTHLY_SERIES = {
    'headline_yoy': 'HCPI', 'core_yoy': 'Core', 'food_yoy': 'Food',
    'fx_mom': 'FX', 'oil_mom': 'Oil', 'm2_mom': 'M2', 'cps_yoy': 'CPS',
    'mpr': 'MPR', 'cpd': 'CPD', 'risk': 'Risk', 'pmi': 'PMI',
}
# Conservative publication *proxy* expressed as the number of full calendar
# months between reference month and last potentially available data month.
# 0=>month end, 1=>following month end, 2=>two months following month end.
# These are end-of-month conservative proxies, NOT verified day-level
# historical publication dates; explicit observations override them.
MONTH_LAGS = {
    'HCPI': 1, 'Core': 1, 'Food': 1, 'FX': 0, 'Oil': 0,
    'M2': 1, 'CPS': 1, 'MPR': 0, 'CPD': 1, 'Risk': 0, 'PMI': 1,
}
CPI_VARIABLES = {'headline_yoy', 'core_yoy', 'food_yoy'}


def end_of_month(year: int, month: int) -> date:
    return date(year, month, calendar.monthrange(year, month)[1])


def add_months(year: int, month: int, delta: int) -> tuple[int,int]:
    n = year * 12 + (month - 1) + delta
    return n // 12, n % 12 + 1


def parse_quarter(period: str):
    year, quarter = period.strip().upper().split('Q')
    return int(year), int(quarter)


class ReleaseGuard:
    def __init__(self, calendar_csv: Path | None = None):
        self.explicit = {}
        self.calendar_csv = calendar_csv
        if calendar_csv and Path(calendar_csv).exists():
            with open(calendar_csv, newline='', encoding='utf-8-sig') as stream:
                for r in csv.DictReader(stream):
                    series = r.get('series', '').strip()
                    period = r.get('reference_period', '').strip()
                    published = r.get('release_date', '').strip()
                    if series and period and published:
                        self.explicit[(series, period)] = date.fromisoformat(published)

    def gdp_date(self, year: int, quarter: int) -> date:
        period = f'{year}Q{quarter}'
        if ('GDP', period) in self.explicit:
            return self.explicit[('GDP', period)]
        end_q = end_of_month(year, quarter * 3)
        # Conservative fallback, not an asserted NBS historical publication date.
        return end_q + timedelta(days=62)

    def gdp_available(self, year: int, quarter: int, origin: date) -> bool:
        return self.gdp_date(year, quarter) <= origin

    def monthly_release(self, series: str, year: int, month: int) -> date:
        period = f'{year:04d}-{month:02d}'
        if (series, period) in self.explicit:
            return self.explicit[(series, period)]
        yy, mm = add_months(year, month, MONTH_LAGS[series])
        return end_of_month(yy, mm)

    def monthly_available(self, predictor: str, year: int, month: int, origin: date) -> bool:
        series = MONTHLY_SERIES[predictor]
        return self.monthly_release(series, year, month) <= origin

    def feature_value(self, monthly_predictors, predictor, month_key, origin):
        y, m = month_key
        # Series comparability is validated when constructing predictors.
        # Preserve compatible headline CPI (including 2025), while leaving
        # invalid sub-index cross-base comparisons as missing.
        if not self.monthly_available(predictor, y, m, origin):
            return math.nan
        try:
            return float(monthly_predictors.get(month_key, {}).get(predictor))
        except (TypeError, ValueError):
            return math.nan

    def credit_momentum(self, monthly_predictors, origin: date) -> float:
        # Exactly six contiguous months ending with the newest published CPS
        # growth figure. Do not interpolate or skip missing months.
        available = [(y,m) for (y,m), r in monthly_predictors.items()
                     if self.monthly_available('cps_yoy', y, m, origin)]
        if not available: return math.nan
        last = max(available)
        last_i = last[0] * 12 + last[1] - 1
        numbers = []
        for i in range(last_i-5, last_i+1):
            y, m = i//12, i%12 + 1
            val = self.feature_value(monthly_predictors, 'cps_yoy', (y,m), origin)
            if not math.isfinite(val): return math.nan
            numbers.append(val)
        return float(sum(numbers[3:])/3 - sum(numbers[:3])/3)


def snapshot_archive(path: Path, monthly, quarterly, origin: date):
    """Store what the engine *saw* at run time; no historical release backfill claims.

    Values for old reference periods are marked legacy_backfill_on_first_run.
    Future runs add new observed_at_run records if the observation changes.
    """
    path.parent.mkdir(parents=True, exist_ok=True)
    existing = set()
    columns = ['observed_as_of', 'series','reference_period','value','snapshot_status']
    is_new = not path.exists()
    if not is_new:
        with path.open(newline='', encoding='utf-8') as stream:
            for item in csv.DictReader(stream):
                existing.add((item['observed_as_of'],item['series'],item['reference_period'],item['value']))
    rows=[]
    guard = ReleaseGuard(path.with_name('release_calendar.csv'))
    for r in monthly:
        p=f'{int(r["year"]):04d}-{int(r["month"]):02d}'
        for source in sorted(set(MONTHLY_SERIES.values())):
            if guard.monthly_release(source, int(r['year']), int(r['month'])) > origin:
                continue
            val=r.get(source)
            try: v=float(val)
            except (TypeError,ValueError): continue
            if not math.isfinite(v):continue
            status=('legacy_backfill_on_first_run' if is_new and (origin.year*12+origin.month)-(r['year']*12+r['month']) > 3 else 'observed_at_run')
            rows.append([origin.isoformat(),source,p,str(v),status])
    for r in quarterly:
        if not guard.gdp_available(int(r['year']), int(r['quarter']), origin):
            continue
        v=float(r['growth']);p=f'{r["year"]}Q{r["quarter"]}' 
        rows.append([origin.isoformat(),'GDP',p,str(v),'legacy_backfill_on_first_run' if is_new and (origin.year*4+(origin.month-1)//3)-(r['year']*4+r['quarter']-1)>2 else 'observed_at_run'])
    with path.open('a',newline='',encoding='utf-8') as stream:
        writer=csv.writer(stream)
        if is_new:writer.writerow(columns)
        for row in rows:
            if tuple(row[:4]) not in existing:writer.writerow(row)


def prospective_registry(path: Path, scored_path: Path, shadow: dict, target_period: str,
                         monthly: list, quarterly: list, origin: date, guard: ReleaseGuard):
    """Register future forecasts before the target release and score them later.

    Every successive run appends a separate dated forecast. Scores reflect the
    first observed outcome in that run, without claiming original-source vintage.
    """
    path.parent.mkdir(parents=True, exist_ok=True)
    columns=['as_of_date','target_period','vintage','model','point_forecast_pp',
             'cps_momentum','data_mode','registration_status']
    rows=[];is_new=not path.exists()
    already=set()
    if not is_new:
        with path.open(newline='',encoding='utf-8') as f:
            already={(r['as_of_date'],r['target_period'],r['vintage'],r['model']) for r in csv.DictReader(f)}
    target_y,target_q = parse_quarter(target_period)
    if guard.gdp_available(target_y,target_q,origin):
        raise RuntimeError('Cannot register a supposedly prospective forecast after the target GDP release date.')
    for model,pred in shadow['current_model_forecasts'].items():
        key=(origin.isoformat(),target_period,shadow['vintage'],model)
        if key in already:continue
        rows.append([*key,str(pred),str(shadow.get('feature_value')),
                     'release_lagged_final_vintage_proxy','registered_before_target_release'])
    with path.open('a',newline='',encoding='utf-8') as f:
        writer=csv.writer(f)
        if is_new:writer.writerow(columns)
        writer.writerows(rows)

    known={(f"{r['year']}Q{r['quarter']}"):float(r['growth']) for r in quarterly}
    completed=[]
    with path.open(newline='',encoding='utf-8') as f:
        for r in csv.DictReader(f):
            period=r['target_period']
            if period not in known:continue
            y,q=parse_quarter(period)
            release=guard.gdp_date(y,q)
            if release>origin:continue
            pred=float(r['point_forecast_pp']);outcome=known[period]
            completed.append([r['as_of_date'],period,r['vintage'],r['model'],
                              pred,outcome,round(outcome-pred,6),
                              round((outcome-pred)**2,6),origin.isoformat(),
                              release.isoformat(), 'observed_at_scoring_run'])
    scoring_columns=['as_of_date','target_period','vintage','model','point_forecast_pp',
                     'gdp_actual_pp','error_pp','squared_error_pp','scored_as_of',
                     'gdp_release_date_or_assumed','scoring_note']
    with scored_path.open('w',newline='',encoding='utf-8') as f:
        writer=csv.writer(f);writer.writerow(scoring_columns);writer.writerows(completed)
    return len(rows),len(completed)
