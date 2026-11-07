"""Refresh the existing Nigeria model and newsletter without a PowerShell script."""
from __future__ import annotations

import argparse
import calendar
import hashlib
import json
import math
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
import urllib.request
from datetime import date

import openpyxl


def shared_hash(path):
    """Read only; permit Excel's shared access on Windows."""
    if os.name == "nt":
        import ctypes
        from ctypes import wintypes
        import msvcrt
        kernel = ctypes.WinDLL("kernel32", use_last_error=True)
        create = kernel.CreateFileW
        create.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD,
                           wintypes.LPVOID, wintypes.DWORD, wintypes.DWORD, wintypes.HANDLE]
        create.restype = wintypes.HANDLE
        handle = create(str(path), 0x80000000, 3, None, 3, 0x80, None)
        if handle == wintypes.HANDLE(-1).value:
            raise ctypes.WinError(ctypes.get_last_error())
        try:
            descriptor = msvcrt.open_osfhandle(handle, os.O_RDONLY | os.O_BINARY)
        except Exception:
            close = kernel.CloseHandle
            close.argtypes = [wintypes.HANDLE]
            close(handle)
            raise
        stream = os.fdopen(descriptor, "rb")
    else:
        stream = open(path, "rb")
    with stream:
        digest = hashlib.sha256()
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
        return digest.hexdigest()


def month_value(value):
    if not re.fullmatch(r"\d{4}-(0[1-9]|1[0-2])", value):
        raise argparse.ArgumentTypeError("Use YYYY-MM, for example 2026-10.")
    try:
        date.fromisoformat(value + "-01")
    except ValueError as exc:
        raise argparse.ArgumentTypeError(str(exc)) from exc
    return value


def read_hcpi_levels(path: Path) -> dict[date, float]:
    """Read dated headline CPI index observations from the monthly data sheet."""
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb["monthly data"]
    headers = [str(cell.value).strip() if cell.value is not None else "" for cell in ws[2]]
    if "Date" not in headers or "HCPI" not in headers:
        wb.close()
        raise RuntimeError("The monthly data sheet must contain Date and HCPI columns.")
    date_col, hcpi_col = headers.index("Date"), headers.index("HCPI")
    current_year = None
    levels = {}
    for row in ws.iter_rows(min_row=3, values_only=True):
        raw = row[date_col]
        if raw is None:
            continue
        label = str(raw).strip().upper()
        if "M" not in label:
            continue
        if label.startswith(("19", "20")):
            year_text, month_text = label.split("M", 1)
            current_year = int(year_text)
        else:
            month_text = label.lstrip("M")
        if current_year is None:
            continue
        try:
            month = int(month_text)
            hcpi = float(row[hcpi_col])
        except (TypeError, ValueError):
            continue
        if 1 <= month <= 12 and math.isfinite(hcpi):
            levels[date(current_year, month, 1)] = hcpi
    wb.close()
    if not levels:
        raise RuntimeError("No dated headline CPI observation was found in the master workbook.")
    return levels


def latest_hcpi_month(path: Path) -> str:
    """Return the month after the latest non-empty headline CPI observation."""
    latest = max(read_hcpi_levels(path))
    year, month = latest.year, latest.month
    if month == 12:
        year, month = year + 1, 1
    else:
        month += 1
    return f"{year:04d}-{month:02d}"


def previous_month_date(value: date) -> date:
    if value.month == 1:
        return date(value.year - 1, 12, 1)
    return date(value.year, value.month - 1, 1)


def inflation_actual_context(path: Path, target_month: str) -> dict | None:
    """Return the two latest published annual inflation rates before the forecast month."""
    levels = read_hcpi_levels(path)
    target = date.fromisoformat(target_month + "-01")
    eligible = sorted(period for period in levels if period < target)
    for latest in reversed(eligible):
        prior = previous_month_date(latest)
        latest_year_ago = date(latest.year - 1, latest.month, 1)
        prior_year_ago = date(prior.year - 1, prior.month, 1)
        required = (prior, latest_year_ago, prior_year_ago)
        if all(period in levels and levels[period] != 0 for period in required):
            latest_yoy = 100.0 * (levels[latest] / levels[latest_year_ago] - 1.0)
            prior_yoy = 100.0 * (levels[prior] / levels[prior_year_ago] - 1.0)
            return {
                "latest_month": latest,
                "prior_month": prior,
                "latest_yoy": latest_yoy,
                "prior_yoy": prior_yoy,
                "change": latest_yoy - prior_yoy,
            }
    return None


def actual_inflation_statement(context: dict) -> str:
    latest = context["latest_yoy"]
    prior = context["prior_yoy"]
    change = context["change"]
    latest_name = context["latest_month"].strftime("%B %Y")
    prior_name = context["prior_month"].strftime("%B")
    if change > 0.05:
        return (
            f"Latest actual: headline inflation rose to {latest:.2f}% in {latest_name}, "
            f"from {prior:.2f}% in {prior_name}. The annual rate accelerated by "
            f"{change:.2f} percentage points, signalling firmer inflationary pressure."
        )
    if change < -0.05:
        return (
            f"Latest actual: headline inflation eased to {latest:.2f}% in {latest_name}, "
            f"from {prior:.2f}% in {prior_name}. The annual rate moderated by "
            f"{abs(change):.2f} percentage points, indicating continued disinflation."
        )
    return (
        f"Latest actual: headline inflation was broadly unchanged at {latest:.2f}% in "
        f"{latest_name}, compared with {prior:.2f}% in {prior_name}. Inflationary momentum "
        "was therefore broadly stable."
    )


def forecast_inflation_statement(point: float, context: dict, target_name: str) -> str:
    latest = context["latest_yoy"]
    latest_name = context["latest_month"].strftime("%B")
    change = point - latest
    if change > 0.05:
        return (
            f"The {target_name} outlook points to renewed inflationary pressure: headline "
            f"inflation rises from {latest:.2f}% in {latest_name} to {point:.2f}%, an increase "
            f"of {change:.2f} percentage points."
        )
    if change < -0.05:
        return (
            f"The {target_name} outlook extends the disinflationary trajectory: "
            f"headline inflation eases from {latest:.2f}% in {latest_name} to {point:.2f}%, "
            f"a decline of {abs(change):.2f} percentage points."
        )
    return (
        f"The {target_name} outlook points to broadly stable price pressures: "
        f"headline inflation moves from {latest:.2f}% in {latest_name} to {point:.2f}%."
    )


def prompt_number(label: str) -> float:
    while True:
        raw = input(label).strip().replace("%", "").replace(",", "")
        try:
            value = float(raw)
        except ValueError:
            print("Enter a number, for example 15.40.")
            continue
        if not math.isfinite(value) or not -100.0 < value < 200.0:
            print("Enter a finite percentage between -100 and 200.")
            continue
        return value


def prompt_optional_number(label: str, default: float) -> float:
    while True:
        raw = input(f"{label} [{default:.2f}]: ").strip().replace("%", "").replace(",", "")
        if not raw:
            return default
        try:
            value = float(raw)
        except ValueError:
            print("Enter a number or press Enter to retain the displayed value.")
            continue
        if not math.isfinite(value) or not -100.0 < value < 200.0:
            print("Enter a finite percentage between -100 and 200.")
            continue
        return value


def previous_month(month_text: str) -> str:
    year, month = map(int, month_text.split("-"))
    if month == 1:
        return f"{year - 1:04d}-12"
    return f"{year:04d}-{month - 1:02d}"


def inspect_master_pmi(path: Path) -> dict:
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb["monthly data"]
    headers = [str(cell.value).strip() if cell.value is not None else "" for cell in ws[2]]
    if "Date" not in headers or "PMI" not in headers:
        wb.close()
        raise RuntimeError("The monthly data sheet must contain Date and PMI columns.")
    date_col, pmi_col = headers.index("Date"), headers.index("PMI")
    current_year = None
    dated = []
    undated = []
    for row_number, row in enumerate(ws.iter_rows(min_row=3, values_only=True), start=3):
        raw, pmi_raw = row[date_col], row[pmi_col]
        try:
            pmi = float(pmi_raw)
        except (TypeError, ValueError):
            continue
        if not math.isfinite(pmi):
            continue
        if raw is None or not str(raw).strip():
            undated.append({"row": row_number, "value": pmi})
            continue
        label = str(raw).strip().upper()
        if "M" not in label:
            continue
        if label.startswith(("19", "20")):
            year_text, month_text = label.split("M", 1)
            current_year = int(year_text)
        else:
            month_text = label.lstrip("M")
        if current_year is None:
            continue
        try:
            month = int(month_text)
        except ValueError:
            continue
        if 1 <= month <= 12:
            dated.append({"month": f"{current_year:04d}-{month:02d}", "value": pmi, "row": row_number})
    wb.close()
    if not dated:
        raise RuntimeError("No dated PMI observation was found in the master workbook.")
    return {"latest": dated[-1], "undated": undated}


def pmi_override_inputs(args, master: Path) -> tuple[str | None, float | None]:
    expected = args.pmi_month or previous_month(args.report_month)
    if args.pmi_month is not None and args.pmi_value is None:
        raise ValueError("--pmi-month requires --pmi-value.")
    if args.pmi_value is None:
        # No PMI check or prompt: the GDP nowcast uses the PMI in the master workbook as it is.
        # A one-off value can still be given with --pmi-value (and --pmi-month).
        return None, None
    audit = inspect_master_pmi(master)
    latest = audit["latest"]
    stale = latest["month"] < expected
    print(
        f"\nPMI DATA CHECK\nLatest dated master observation: {latest['month']} = {latest['value']:.1f}. "
        f"Expected for this newsletter: {expected}."
    )
    if audit["undated"]:
        details = ", ".join(f"row {item['row']} = {item['value']:.1f}" for item in audit["undated"])
        print(f"Warning: undated PMI value(s) were found and cannot be assigned safely: {details}.")
    if args.pmi_value is not None:
        value = float(args.pmi_value)
        if not math.isfinite(value) or not 0 < value < 100:
            raise ValueError("PMI override must be between 0 and 100.")
        return expected, value
    if args.dry_run or not stale:
        return None, None
    print("The dated PMI series is stale. Confirm the latest official release before continuing.")
    raw = input(f"Enter the official PMI for {expected}, or press Enter to continue with the dated master series: ").strip()
    if not raw:
        print("Continuing without a PMI override; the GDP nowcast will retain the dated master series.")
        return None, None
    try:
        value = float(raw)
    except ValueError as exc:
        raise ValueError("PMI override must be numeric.") from exc
    if not math.isfinite(value) or not 0 < value < 100:
        raise ValueError("PMI override must be between 0 and 100.")
    print(f"Using PMI {value:.1f} for {expected} in this run; the master workbook remains unchanged.")
    return expected, value


def collect_monthly_research(report_month: str, destination: Path) -> Path:
    """Record an auditable set of current official sources for editorial review."""
    sources = [
        ("Central Bank of Nigeria monetary policy decisions", "https://www.cbn.gov.ng/MonetaryPolicy/decisions.html"),
        ("Central Bank of Nigeria data and PMI repository", "https://www.cbn.gov.ng/rates/"),
        ("National Bureau of Statistics", "https://www.nigerianstat.gov.ng/"),
        ("Stanbic IBTC purchasing managers reports", "https://www.stanbicibtcbank.com/nigeriabank/corporate-and-investment/insights"),
        ("IMF World Economic Outlook", "https://www.imf.org/en/Publications/WEO"),
        ("World Bank Global Economic Prospects", "https://www.worldbank.org/en/publication/global-economic-prospects"),
    ]
    checks = []
    headers = {"User-Agent": "Mozilla/5.0 CAPE-ERC-monthly-research/1.0"}
    for label, url in sources:
        status, title = "unavailable", ""
        try:
            request = urllib.request.Request(url, headers=headers)
            with urllib.request.urlopen(request, timeout=12) as response:
                status = f"HTTP {getattr(response, 'status', 200)}"
                body = response.read(400_000).decode("utf-8", errors="ignore")
            match = re.search(r"<title[^>]*>(.*?)</title>", body, flags=re.IGNORECASE | re.DOTALL)
            if match:
                title = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", "", match.group(1))).strip()
        except Exception as exc:
            status = f"unavailable: {type(exc).__name__}"
        checks.append({"label": label, "url": url, "status": status, "page_title": title})
    destination.mkdir(parents=True, exist_ok=True)
    path = destination / f"Monthly Research Sources {report_month}.json"
    path.write_text(
        json.dumps(
            {
                "report_month": report_month,
                "purpose": "Official-source register for monthly economic editorial review",
                "sources": checks,
            },
            indent=2,
        ),
        encoding="utf-8",
    )
    available = sum(item["status"].startswith("HTTP") for item in checks)
    print(f"Monthly research source register: {path} ({available} of {len(checks)} sources reached)")
    return path


def inflation_inputs(args, master: Path) -> tuple[str, float | None, float | None, float | None]:
    target_month = args.inflation_month or latest_hcpi_month(master)
    point, lower, upper = args.inflation_point, args.inflation_lower, args.inflation_upper
    if args.dry_run:
        return target_month, point, lower, upper
    month_name = date.fromisoformat(target_month + "-01").strftime("%B %Y")
    print(f"\nINFLATION FORECAST INPUTS FOR {month_name.upper()}")
    context = inflation_actual_context(master, target_month)
    if context:
        print(actual_inflation_statement(context))
    if point is None:
        point = prompt_number("Expected headline inflation (% year-on-year): ")
    if context:
        print(forecast_inflation_statement(point, context, month_name))
    if lower is None:
        lower = prompt_number("Forecast range lower bound (% year-on-year): ")
    if upper is None:
        upper = prompt_number("Forecast range upper bound (% year-on-year): ")
    if lower > upper:
        raise ValueError("Inflation forecast lower bound cannot exceed the upper bound.")
    if not lower <= point <= upper:
        raise ValueError("The expected inflation rate must fall within the forecast range.")
    print(f"Publishing {month_name} headline inflation at {point:.2f}%, within a "
          f"{lower:.2f}%–{upper:.2f}% forecast range.")
    if context:
        change = point - context["latest_yoy"]
        latest_name = context["latest_month"].strftime("%B")
        if change > 0.05:
            print(f"This would mark renewed inflationary pressure relative to {latest_name}'s "
                  f"{context['latest_yoy']:.2f}% outturn.")
        elif change < -0.05:
            print(f"This would extend the easing in annual inflation relative to "
                  f"{latest_name}'s {context['latest_yoy']:.2f}% outturn.")
        else:
            print(f"This would leave annual inflation broadly stable relative to {latest_name}'s "
                  f"{context['latest_yoy']:.2f}% outturn.")
    return target_month, point, lower, upper


def gdp_inputs(args, results_path: Path) -> tuple[float, float, float]:
    results = json.loads(results_path.read_text(encoding="utf-8"))
    model_point = float(results["published_candidate_forecast_pct"])
    model_lower = float(results.get("reporting_range_lower_pct", results["pi50_lower_pct"]))
    model_upper = float(results.get("reporting_range_upper_pct", results["pi50_upper_pct"]))
    target = str(results["target_quarter"])
    print("\n" + "=" * 64)
    print("NEW GDP NOWCAST RESULTS")
    print("=" * 64)
    print(f"Target quarter: {target}")
    print(f"Model point nowcast: {model_point:.2f}% year-on-year")
    print(f"CAPE reporting range: {model_lower:.2f}% to {model_upper:.2f}%")
    print(f"Information vintage: {results['vintage']}")
    print(f"Models in ensemble: {results['models_eligible']}")
    print(f"Best single model: {results['best_single_model_last8']} ({results['best_single_model_forecast_pct']:.2f}%)")

    if args.gdp_point is not None or args.gdp_lower is not None or args.gdp_upper is not None:
        point = model_point if args.gdp_point is None else float(args.gdp_point)
        lower = model_lower if args.gdp_lower is None else float(args.gdp_lower)
        upper = model_upper if args.gdp_upper is None else float(args.gdp_upper)
    elif args.accept_model_gdp:
        point, lower, upper = model_point, model_lower, model_upper
    else:
        choice = input(
            "Press Enter to publish the new model point and range, or type O to override them: "
        ).strip().lower()
        if choice in ("", "a", "auto", "accept"):
            point, lower, upper = model_point, model_lower, model_upper
        elif choice in ("o", "override"):
            point = prompt_optional_number("Expected GDP growth (% year-on-year)", model_point)
            lower = prompt_optional_number("GDP forecast range lower bound (% year-on-year)", model_lower)
            upper = prompt_optional_number("GDP forecast range upper bound (% year-on-year)", model_upper)
        else:
            raise ValueError("GDP publication choice must be Enter for automatic or O for override.")
    if lower > upper:
        raise ValueError("GDP forecast lower bound cannot exceed the upper bound.")
    if not lower <= point <= upper:
        raise ValueError("The expected GDP growth rate must fall within the forecast range.")
    print(f"Publishing GDP growth at {point:.2f}% with a CAPE reporting range of {lower:.2f}%–{upper:.2f}%.")
    return point, lower, upper


def report_paths(root, month, template=None):
    year, number = map(int, month.split("-"))
    name = calendar.month_name[number]
    folder = root / str(year) / f"{name} Newsletter"
    stem = f"CAPE Economic Performance and Prospect Bulletin {name} {year}"
    output = folder / f"{stem}.docx"
    previous_number = number - 1 or 12
    previous_year = year if number > 1 else year - 1
    previous_name = calendar.month_name[previous_number]
    previous = root / str(previous_year) / f"{previous_name} Newsletter" / (
        f"CAPE Economic Performance and Prospect Bulletin {previous_name} {previous_year}.docx")
    candidates = [Path(template).resolve()] if template else [output, folder / f"{stem} - Template.docx", previous]
    selected = next((p for p in candidates if p.is_file()), None)
    if selected is None:
        raise FileNotFoundError("No newsletter template found. Supply --newsletter-template PATH. Checked: "
                                + "; ".join(map(str, candidates)))
    return output, selected


def stamp(path):
    return path.stat().st_mtime_ns if path.exists() else None


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("report_month", nargs="?", type=month_value, default=date.today().strftime("%Y-%m"),
                        help="Newsletter month YYYY-MM (default: current month); model quarter remains data-driven.")
    parser.add_argument("--as-of", type=date.fromisoformat, default=date.today(),
                        help="Information date YYYY-MM-DD (default: today).")
    parser.add_argument("--target-quarter", help="Optional YYYYQn override; normally selected automatically.")
    parser.add_argument("--newsletter-template", type=Path)
    parser.add_argument("--plots-folder", type=Path)
    parser.add_argument("--gdp-point", type=float, help="Expected GDP growth, year-on-year percent.")
    parser.add_argument("--gdp-lower", type=float, help="Lower bound of the GDP growth forecast range.")
    parser.add_argument("--gdp-upper", type=float, help="Upper bound of the GDP growth forecast range.")
    parser.add_argument("--accept-model-gdp", action="store_true",
                        help="Publish the newly computed GDP point and empirical range without prompting.")
    parser.add_argument("--pmi-month", type=month_value,
                        help="Month YYYY-MM for a temporary PMI override; defaults to the month before the newsletter.")
    parser.add_argument("--pmi-value", type=float,
                        help="Temporary PMI value used by the GDP model without modifying the master workbook.")
    parser.add_argument("--skip-web-research", action="store_true",
                        help="Do not check the standard official monthly research sources.")
    parser.add_argument("--inflation-month", type=month_value,
                        help="Inflation forecast month YYYY-MM; defaults to the month after the latest actual CPI.")
    parser.add_argument("--inflation-point", type=float, help="Expected headline inflation, year-on-year percent.")
    parser.add_argument("--inflation-lower", type=float, help="Lower bound of the headline inflation forecast range.")
    parser.add_argument("--inflation-upper", type=float, help="Upper bound of the headline inflation forecast range.")
    parser.add_argument("--dry-run", action="store_true", help="Check paths and show commands; do not change any files.")
    args = parser.parse_args(argv)
    project_root = Path(__file__).resolve().parent
    nowcast_root = project_root / "Nigeria_GDP_Nowcast"
    model_root = project_root.parent
    automation = nowcast_root / ".automation"
    master = model_root / "Main Data" / "MultipleFrequencyModelData.xlsx"
    newsletter_root = project_root / "CAPE ERC Economic Newsletter"
    output, template = report_paths(newsletter_root, args.report_month, args.newsletter_template)
    plots = args.plots_folder.resolve() if args.plots_folder else output.parent / "Plots"
    for path in [master, nowcast_root / "run_nigeria_nowcast.py", nowcast_root / "check_package.py",
                 automation / "publish_newsletter.py"]:
        if not path.is_file():
            raise FileNotFoundError(path)
    python_candidates = [sys.executable, shutil.which("python")]
    python_executable = None
    for candidate in dict.fromkeys(str(item) for item in python_candidates if item):
        probe = subprocess.run(
            [candidate, "-c", "import pandas, sklearn, openpyxl, PIL, docx"],
            capture_output=True,
            text=True,
        )
        if probe.returncode == 0:
            python_executable = candidate
            break
    if python_executable is None:
        raise RuntimeError(
            "The nowcasting automation requires Python with pandas, scikit-learn, openpyxl, Pillow and python-docx."
        )
    python = [python_executable, "-X", "utf8"]
    workbook = nowcast_root / "Results" / "Nigeria_GDP_8_Model_Results.xlsx"
    pmi_month, pmi_value = pmi_override_inputs(args, master)
    model_command = python + [str(nowcast_root / "run_nigeria_nowcast.py"), "--as-of", args.as_of.isoformat()]
    if args.target_quarter:
        model_command += ["--target-quarter", args.target_quarter.upper()]
    commands = [model_command, python + [str(nowcast_root / "check_package.py"), "--require-excel"]]
    if args.dry_run:
        print("DRY RUN: no model estimation, files changed, or report published.")
        print(f"Read-only master: {master}\nReport: {output}")
        for command in commands:
            print(subprocess.list2cmdline(command))
        print("After validation, the run displays the new GDP nowcast, then asks whether to accept it or override it.")
        print("Inflation point and range inputs are requested only after the GDP results are visible.")
        return 0
    env = dict(
        os.environ,
        NIGERIA_NOWCAST_OUTPUT_DIR=str(nowcast_root),
        NIGERIA_SKIP_SOURCE_UPDATE="1",
        PYTHONDONTWRITEBYTECODE="1",
    )
    if pmi_month is not None and pmi_value is not None:
        env["NIGERIA_PMI_OVERRIDE_MONTH"] = pmi_month
        env["NIGERIA_PMI_OVERRIDE_VALUE"] = str(pmi_value)

    def run(index, check=True):
        return subprocess.run(commands[index], cwd=nowcast_root, env=env, check=check).returncode

    before = shared_hash(master)
    results = nowcast_root / "Results" / "forecast.json"
    results_before = stamp(results)
    workbook_before = stamp(workbook)
    run(0)
    if stamp(results) is None or stamp(results) == results_before:
        raise RuntimeError("Estimation did not refresh forecast.json; stopping.")
    if stamp(workbook) is None:
        raise RuntimeError("Model run did not create the Excel workbook; stopping.")
    builder_status = run(1, check=False)
    if stamp(workbook) is None or stamp(workbook) == workbook_before:
        raise RuntimeError(f"Model run did not save a fresh workbook (validation status {builder_status}); stopping.")
    if builder_status:
        raise RuntimeError(f"Independent model validation failed with status {builder_status}; stopping.")
    if shared_hash(master) != before:
        raise RuntimeError("Master workbook changed during estimation/build; publication stopped. Do not edit during refresh.")

    # Publication decisions are deliberately deferred until the model has run
    # and its refreshed forecast, range and PMI vintage are visible.
    if not args.skip_web_research:
        collect_monthly_research(args.report_month, output.parent)
    gdp_point, gdp_lower, gdp_upper = gdp_inputs(args, results)
    inflation_month, inflation_point, inflation_lower, inflation_upper = inflation_inputs(args, master)
    inflation_chart = nowcast_root / "Newsletter Assets" / "Nigeria Inflation Forecast Chart.png"
    publication_command = python + [
        str(automation / "publish_newsletter.py"),
        "--workbook", str(workbook),
        "--template", str(template),
        "--output", str(output),
        "--chart", str(nowcast_root / "Newsletter Assets/Nigeria GDP Nowcast Chart.png"),
        "--pmi-chart", str(nowcast_root / "Newsletter Assets/Nigeria PMI Trend Chart.png"),
        "--inflation-chart", str(inflation_chart),
        "--report-month", args.report_month,
        "--plots-folder", str(plots),
        "--gdp-point", str(gdp_point),
        "--gdp-lower", str(gdp_lower),
        "--gdp-upper", str(gdp_upper),
        "--inflation-source", str(master),
        "--inflation-month", inflation_month,
        "--inflation-point", str(inflation_point),
        "--inflation-lower", str(inflation_lower),
        "--inflation-upper", str(inflation_upper),
    ]
    output.parent.mkdir(parents=True, exist_ok=True)
    plots.mkdir(parents=True, exist_ok=True)
    subprocess.run(publication_command, cwd=nowcast_root, env=env, check=True)
    if shared_hash(master) != before:
        raise RuntimeError("Master workbook changed during publication. Review the report and rerun after saving inputs.")
    print(f"Nigeria GDP nowcast refreshed: {workbook}")
    print(f"Economic newsletter refreshed: {output}")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except (OSError, RuntimeError, ValueError, subprocess.CalledProcessError) as exc:
        print(f"Refresh failed: {exc}", file=sys.stderr)
        sys.exit(1)
