from __future__ import annotations

import argparse
from decimal import Decimal, ROUND_HALF_UP
import copy
import json
import math
import re
from io import BytesIO
from datetime import datetime
from pathlib import Path

import openpyxl
from docx import Document
from docx.shared import Inches
from PIL import Image, ImageDraw, ImageFont


PROJECT_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_TEMPLATE = (
    PROJECT_ROOT
    / "CAPE ERC Economic Newsletter"
    / "2026"
    / "September Newsletter"
    / "CAPE Economic Performance and Prospect Bulletin September 2026.docx"
)
DEFAULT_OUTPUT = DEFAULT_TEMPLATE


def quarter_label(value: str) -> str:
    value = str(value).strip().upper()
    if "Q" in value:
        year, quarter = value.split("Q", 1)
        return f"Q{quarter} {year}"
    return value


def inflation_momentum(change: float, future: bool = False) -> str:
    if change > 0.05:
        return "is projected to accelerate" if future else "accelerated"
    if change < -0.05:
        return "is projected to ease" if future else "eased"
    return "is projected to remain broadly stable" if future else "was broadly stable"


def percentage_point_change(change: float) -> str:
    direction = "increase" if change > 0 else "decline"
    return f"a {abs(change):.2f} percentage-point {direction}"


def read_inflation_forecast(source: Path, forecast_month: str, point: float, lower: float, upper: float) -> dict:
    if lower > upper or not lower <= point <= upper:
        raise ValueError("The expected inflation rate must fall within the lower and upper range bounds.")
    wb = openpyxl.load_workbook(source, data_only=True, read_only=True)
    ws = wb["monthly data"]
    headers = [str(cell.value).strip() if cell.value is not None else "" for cell in ws[2]]
    if "Date" not in headers or "HCPI" not in headers:
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
            value = float(row[hcpi_col])
        except (TypeError, ValueError):
            continue
        if 1 <= month <= 12 and math.isfinite(value):
            levels[datetime(current_year, month, 1)] = value
    wb.close()

    target = datetime.strptime(forecast_month, "%Y-%m")
    prior_actuals = [d for d in levels if d < target]
    if not prior_actuals:
        raise RuntimeError("No headline CPI actual precedes the inflation forecast month.")
    latest = max(prior_actuals)
    expected = datetime(latest.year + (1 if latest.month == 12 else 0), 1 if latest.month == 12 else latest.month + 1, 1)
    if target != expected:
        raise RuntimeError(
            f"Inflation forecast month {forecast_month} must immediately follow the latest HCPI actual {latest:%Y-%m}."
        )
    target_year_ago = datetime(target.year - 1, target.month, 1)
    latest_year_ago = datetime(latest.year - 1, latest.month, 1)
    prior_month = datetime(latest.year - (1 if latest.month == 1 else 0), 12 if latest.month == 1 else latest.month - 1, 1)
    prior_month_year_ago = datetime(prior_month.year - 1, prior_month.month, 1)
    missing = [
        d.strftime("%Y-%m")
        for d in (target_year_ago, latest_year_ago, prior_month, prior_month_year_ago)
        if d not in levels
    ]
    if missing:
        raise RuntimeError("HCPI history is missing months needed for inflation calculations: " + ", ".join(missing))

    forecast_level = levels[target_year_ago] * (1.0 + point / 100.0)
    lower_level = levels[target_year_ago] * (1.0 + lower / 100.0)
    upper_level = levels[target_year_ago] * (1.0 + upper / 100.0)
    latest_yoy = (levels[latest] / levels[latest_year_ago] - 1.0) * 100.0
    previous_yoy = (levels[prior_month] / levels[prior_month_year_ago] - 1.0) * 100.0
    return {
        "source": str(source),
        "target": target,
        "latest": latest,
        "point": point,
        "lower": lower,
        "upper": upper,
        "latest_yoy": latest_yoy,
        "previous": prior_month,
        "previous_yoy": previous_yoy,
        "latest_yoy_change": latest_yoy - previous_yoy,
        "forecast_yoy_change": point - latest_yoy,
        "latest_mom": (levels[latest] / levels[prior_month] - 1.0) * 100.0,
        "forecast_level": forecast_level,
        "forecast_mom": (forecast_level / levels[latest] - 1.0) * 100.0,
        "forecast_mom_lower": (lower_level / levels[latest] - 1.0) * 100.0,
        "forecast_mom_upper": (upper_level / levels[latest] - 1.0) * 100.0,
        "history": levels,
    }


def read_nowcast(workbook_path: Path) -> dict:
    wb = openpyxl.load_workbook(workbook_path, data_only=True, read_only=True)
    if not {"Quarterly Forecast", "Model Engine", "Monthly Inputs"}.issubset(wb.sheetnames):
        wb.close()
        return read_preferred_nowcast(workbook_path)
    summary = wb["Executive Summary"]
    forecast = wb["Quarterly Forecast"]
    engine = wb["Model Engine"]
    monthly = wb["Monthly Inputs"]
    models = []
    for row in engine.iter_rows(min_row=4, max_col=2, values_only=True):
        if row[1] is None:
            break
        models.append(str(row[1]))
    if len(models) > 1:
        feature_text = ", ".join(models[:-1]) + f" and {models[-1]}"
    else:
        feature_text = models[0] if models else "the ensemble models"
    values = {
        "internal_point": float(engine["H4"].value),
        "release_mode": "Nowcast with PI50",
        "complete_months": int(engine["H13"].value),
        "total_months": 3,
        "pmi_months_available": int(engine["H15"].value),
        "latest_actual": float(summary["A7"].value),
        "latest_period": str(summary["B7"].value),
        "pmi": float(summary["C7"].value),
        "model": str(summary["E7"].value),
        "rmse": float(summary["F7"].value),
        "mae": float(summary["G7"].value),
        "pmi_treatment": "Included where supported by the candidate specification",
        "features": models,
        "feature_text": feature_text,
        "required_inputs": str(engine["H19"].value),
        "quarters": [],
        "pmi_monthly": [],
    }
    for row in monthly.iter_rows(min_row=2, values_only=True):
        observation_date = row[0]
        pmi_value = row[11] if len(row) > 11 else None
        if observation_date is not None and isinstance(pmi_value, (int, float)):
            values["pmi_monthly"].append({"date": observation_date, "value": float(pmi_value)})
    for row in forecast.iter_rows(min_row=4, values_only=True):
        if not row[0]:
            continue
        status = str(row[8] or "")
        actual = float(row[1]) if isinstance(row[1], (int, float)) else None
        point = float(row[5]) if isinstance(row[5], (int, float)) else None
        lower = float(row[6]) if isinstance(row[6], (int, float)) else None
        upper = float(row[7]) if isinstance(row[7], (int, float)) else None
        values["quarters"].append({"period": str(row[0]), "actual": actual, "point": point, "lower": lower, "upper": upper, "status": status})
        if actual is None and status.lower() != "official actual":
            values["target_period"] = str(row[0])
            values["point"] = point
            values["lower"] = lower
            values["upper"] = upper
    wb.close()
    if "target_period" not in values:
        raise RuntimeError("No nowcast row was found in the Quarterly Forecast sheet.")
    return values


def read_preferred_nowcast(workbook_path: Path) -> dict:
    """Read the preferred eight-model output and its shared master data."""
    root = workbook_path.parent
    if root.name == "Results":          # the workbook now lives in Results\, next to forecast.json
        root = root.parent
    results = json.loads((root / "Results" / "forecast.json").read_text(encoding="utf-8"))
    # forecast.json does not record the master path; use the shared CAPE master workbook
    master = Path(results.get("master_workbook") or PROJECT_ROOT.parent / "Main Data" / "MultipleFrequencyModelData.xlsx")
    wb = openpyxl.load_workbook(master, data_only=True, read_only=True)
    quarterly = []
    for row in wb["Quarterly Data"].iter_rows(min_row=3, values_only=True):
        if row[0] is None or row[1] is None or row[2] is None:
            continue
        growth = float(row[2])
        if abs(growth) < 1:
            growth *= 100.0
        qtext = str(row[1]).strip().upper()
        quarter = int(qtext[1]) if qtext.startswith("Q") else int(row[1])
        quarterly.append({"period": f"{int(row[0])}Q{quarter}", "actual": growth})
    monthly = []
    current_year = None
    for row in wb["monthly data"].iter_rows(min_row=3, values_only=True):
        raw = row[0]
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
            month, pmi = int(month_text), float(row[10])
        except (TypeError, ValueError):
            continue
        if math.isfinite(pmi):
            monthly.append({"date": datetime(current_year, month, 1), "value": pmi})
    wb.close()
    models_path = root / "Results" / "model_forecasts.csv"
    import csv
    with models_path.open(encoding="utf-8-sig", newline="") as handle:
        model_rows = list(csv.DictReader(handle))
    target = results["target_quarter"]
    target_year, target_q = int(target[:4]), int(target[-1])
    target_months = {(target_q - 1) * 3 + offset for offset in (1, 2, 3)}
    complete_months = sum(
        1 for month in target_months
        if any(item["date"].year == target_year and item["date"].month == month for item in monthly)
    )
    latest = quarterly[-1]
    best = next((row for row in model_rows if row["model"] == results["best_single_model_last8"]), model_rows[0])
    quarters = [dict(row, point=None, lower=None, upper=None, status="Official actual") for row in quarterly]
    quarters.append({
        "period": target, "actual": None,
        "point": float(results["published_candidate_forecast_pct"]),
        "lower": float(results["pi50_lower_pct"]),
        "upper": float(results["pi50_upper_pct"]), "status": "Nowcast with PI50",
    })
    model_names = [row["model"] for row in model_rows if float(row.get("weight") or 0) > 0]
    feature_text = ", ".join(model_names[:-1]) + (f" and {model_names[-1]}" if model_names else "")
    current_pmi = [item["value"] for item in monthly if item["date"].year == target_year and item["date"].month in target_months]
    return {
        "internal_point": float(results["statistical_ensemble_pct"]),
        "release_mode": "Nowcast with PI50", "complete_months": complete_months, "total_months": 3,
        "pmi_months_available": len(current_pmi), "latest_actual": latest["actual"],
        "latest_period": latest["period"], "pmi": sum(current_pmi) / len(current_pmi) if current_pmi else monthly[-1]["value"],
        "model": results["best_single_model_last8"], "rmse": float(best["rmse_last8_pp"]),
        "mae": float(best["mae_last8_pp"]), "pmi_treatment": "Included where supported by the candidate specification",
        "features": model_names, "feature_text": feature_text,
        "required_inputs": "Latest released GDP, CPS level, CBN Composite PMI and inflation risk",
        "quarters": quarters, "pmi_monthly": monthly, "target_period": target,
        "point": float(results["published_candidate_forecast_pct"]),
        "lower": float(results["pi50_lower_pct"]), "upper": float(results["pi50_upper_pct"]),
    }


def create_gdp_chart(data: dict, chart_path: Path) -> None:
    recent = [r for r in data["quarters"] if r["actual"] is not None][-10:]
    target = next(r for r in data["quarters"] if r["period"] == data["target_period"])
    labels = [quarter_label(r["period"]) for r in recent] + [quarter_label(target["period"])]
    actual = [r["actual"] for r in recent]
    width, height = 1548, 920
    left, right, top, bottom = 150, 55, 145, 190
    plot_left, plot_right = left, width - right
    plot_top, plot_bottom = top, height - bottom
    image = Image.new("RGB", (width, height), "#FBFCFE")
    draw = ImageDraw.Draw(image)
    font_dir = Path(r"C:\Windows\Fonts")

    def font(size: int, bold: bool = False):
        preferred = font_dir / ("calibrib.ttf" if bold else "calibri.ttf")
        fallback = font_dir / ("arialbd.ttf" if bold else "arial.ttf")
        if preferred.exists() or fallback.exists():
            return ImageFont.truetype(str(preferred if preferred.exists() else fallback), size)
        # Portable fallback for report verification outside Windows.
        linux_font = Path("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold
                          else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf")
        return ImageFont.truetype(str(linux_font), size)

    label_font = font(27)
    small_font = font(25)
    legend_font = font(27)
    bold_font = font(28, bold=True)
    title_font = font(38, bold=True)
    release_values = [v for v in (data.get("point"), data.get("lower"), data.get("upper")) if v is not None]
    y_min = max(0.0, min(release_values + actual) - 0.7)
    y_max = max(release_values + actual) + 0.7
    y_min = (int(y_min * 2) / 2.0)
    y_max = ((int(y_max * 2) + 1) / 2.0)

    def y_px(value: float) -> float:
        return plot_bottom - (value - y_min) / (y_max - y_min) * (plot_bottom - plot_top)

    x_positions = [plot_left + i * (plot_right - plot_left) / (len(labels) - 1) for i in range(len(labels))]
    forecast_left = (x_positions[-2] + x_positions[-1]) / 2
    draw.rectangle((forecast_left, plot_top, plot_right, plot_bottom), fill="#FFF1EF")
    draw.text((plot_right - 145, plot_top + 14), "FORECAST", fill="#A61B1B", font=small_font)
    tick = y_min
    while tick <= y_max + 1e-9:
        y = y_px(tick)
        draw.line((plot_left, y, plot_right, y), fill="#D9E2F3", width=2)
        txt = f"{tick:.1f}"
        bbox = draw.textbbox((0, 0), txt, font=small_font)
        draw.text((plot_left - 18 - (bbox[2] - bbox[0]), y - (bbox[3] - bbox[1]) / 2), txt, fill="#595959", font=small_font)
        tick += 0.5

    draw.line((plot_left, plot_top, plot_left, plot_bottom), fill="#A6A6A6", width=3)
    draw.line((plot_left, plot_bottom, plot_right, plot_bottom), fill="#A6A6A6", width=3)

    actual_points = [(x_positions[i], y_px(v)) for i, v in enumerate(actual)]
    draw.polygon(
        actual_points + [(actual_points[-1][0], plot_bottom), (actual_points[0][0], plot_bottom)],
        fill="#EAF2F8",
    )
    draw.line(actual_points, fill="#1F4E78", width=7, joint="curve")
    for x, y in actual_points:
        draw.ellipse((x - 8, y - 8, x + 8, y + 8), fill="white", outline="#1F4E78", width=5)

    xf = x_positions[len(actual)]
    low_y, high_y, yf = y_px(data["lower"]), y_px(data["upper"]), y_px(data["point"])
    draw.line((x_positions[len(actual) - 1], y_px(actual[-1]), xf, yf), fill="#C00000", width=7)
    draw.line((xf, high_y, xf, low_y), fill="#C00000", width=12)
    draw.line((xf - 25, high_y, xf + 25, high_y), fill="#C00000", width=8)
    draw.line((xf - 25, low_y, xf + 25, low_y), fill="#C00000", width=8)
    draw.polygon([(xf, yf - 13), (xf + 13, yf), (xf, yf + 13), (xf - 13, yf)], fill="#C00000")

    latest_text = f"{data['latest_actual']:.2f}%"
    latest_box = draw.textbbox((0, 0), latest_text, font=bold_font)
    draw.text((actual_points[-1][0] - (latest_box[2] - latest_box[0]) - 10, actual_points[-1][1] + 18), latest_text, fill="#1F4E78", font=bold_font)
    forecast_text = f"{data['point']:.2f}% ({data['lower']:.2f}%–{data['upper']:.2f}%)"
    forecast_label_y = yf + 24 if data["point"] <= data["latest_actual"] else yf - 70
    forecast_label_x = xf - 365
    forecast_box = draw.textbbox((forecast_label_x, forecast_label_y), forecast_text, font=bold_font)
    draw.rectangle((forecast_box[0] - 8, forecast_box[1] - 5, forecast_box[2] + 8, forecast_box[3] + 5), fill="white")
    draw.text((forecast_label_x, forecast_label_y), forecast_text, fill="#C00000", font=bold_font)

    for x, label in zip(x_positions, labels):
        bbox = draw.textbbox((0, 0), label, font=label_font)
        draw.text((x - (bbox[2] - bbox[0]) / 2, plot_bottom + 20), label, fill="#404040", font=label_font)

    draw.text((plot_left, 18), "Nigeria real GDP growth", fill="#17365D", font=title_font)
    # Legend.
    legend_y = 76
    draw.line((plot_left, legend_y + 14, plot_left + 52, legend_y + 14), fill="#1F4E78", width=7)
    draw.text((plot_left + 65, legend_y), "NBS actual", fill="#404040", font=legend_font)
    lx2 = plot_left + 300
    draw.line((lx2 + 26, legend_y - 2, lx2 + 26, legend_y + 31), fill="#C00000", width=10)
    draw.line((lx2 + 10, legend_y - 2, lx2 + 42, legend_y - 2), fill="#C00000", width=6)
    draw.line((lx2 + 10, legend_y + 31, lx2 + 42, legend_y + 31), fill="#C00000", width=6)
    draw.polygon([(lx2 + 26, legend_y + 3), (lx2 + 37, legend_y + 14), (lx2 + 26, legend_y + 25), (lx2 + 15, legend_y + 14)], fill="#C00000")
    draw.text((lx2 + 65, legend_y), "CAPE forecast range", fill="#404040", font=legend_font)

    # Rotated y-axis title.
    y_title = "Real GDP growth (%, year-on-year)"
    title_box = Image.new("RGBA", (800, 70), (255, 255, 255, 0))
    title_draw = ImageDraw.Draw(title_box)
    title_draw.text((0, 0), y_title, fill="#404040", font=label_font)
    title_box = title_box.crop(title_box.getbbox()).rotate(90, expand=True)
    image.paste(title_box, (25, int((plot_top + plot_bottom - title_box.height) / 2)), title_box)

    note = (
        f"Actual through {quarter_label(data['latest_period'])}; {quarter_label(data['target_period'])} is the CAPE "
        f"CAPE growth forecast for {quarter_label(data['target_period'])}, with an expected range."
    )
    draw.text((plot_left, height - 80), note, fill="#595959", font=small_font)
    chart_path.parent.mkdir(parents=True, exist_ok=True)
    image.save(chart_path, format="PNG", optimize=True)


def create_pmi_chart(data: dict, chart_path: Path) -> None:
    series = data["pmi_monthly"][-14:]
    if len(series) < 2:
        raise RuntimeError("At least two monthly PMI observations are required for Figure 2.")

    width, height = 1548, 920
    left, right, top, bottom = 145, 55, 105, 190
    plot_left, plot_right = left, width - right
    plot_top, plot_bottom = top, height - bottom
    image = Image.new("RGB", (width, height), "white")
    draw = ImageDraw.Draw(image)
    font_dir = Path(r"C:\Windows\Fonts")

    def font(size: int, bold: bool = False):
        preferred = font_dir / ("calibrib.ttf" if bold else "calibri.ttf")
        fallback = font_dir / ("arialbd.ttf" if bold else "arial.ttf")
        if preferred.exists() or fallback.exists():
            return ImageFont.truetype(str(preferred if preferred.exists() else fallback), size)
        # Portable fallback for report verification outside Windows.
        linux_font = Path("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold
                          else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf")
        return ImageFont.truetype(str(linux_font), size)

    label_font = font(24)
    small_font = font(25)
    legend_font = font(27)
    bold_font = font(29, bold=True)
    values = [item["value"] for item in series]
    labels = [item["date"].strftime("%b-%y") for item in series]
    y_min = min(48.0, math.floor(min(values) - 0.5))
    y_max = max(58.0, math.ceil(max(values) + 0.5))

    def y_px(value: float) -> float:
        return plot_bottom - (value - y_min) / (y_max - y_min) * (plot_bottom - plot_top)

    x_positions = [plot_left + i * (plot_right - plot_left) / (len(series) - 1) for i in range(len(series))]
    tick = math.ceil(y_min / 2) * 2
    while tick <= y_max + 1e-9:
        y = y_px(tick)
        draw.line((plot_left, y, plot_right, y), fill="#D9E2F3", width=2)
        txt = f"{tick:.0f}"
        bbox = draw.textbbox((0, 0), txt, font=small_font)
        draw.text((plot_left - 18 - (bbox[2] - bbox[0]), y - (bbox[3] - bbox[1]) / 2), txt, fill="#595959", font=small_font)
        tick += 2

    threshold_y = y_px(50.0)
    dash = 20
    x = plot_left
    while x < plot_right:
        draw.line((x, threshold_y, min(x + dash, plot_right), threshold_y), fill="#C00000", width=4)
        x += dash * 2

    draw.line((plot_left, plot_top, plot_left, plot_bottom), fill="#A6A6A6", width=3)
    draw.line((plot_left, plot_bottom, plot_right, plot_bottom), fill="#A6A6A6", width=3)

    points = [(x_positions[i], y_px(value)) for i, value in enumerate(values)]
    draw.line(points, fill="#1F4E78", width=7, joint="curve")
    for x, y in points[:-1]:
        draw.ellipse((x - 7, y - 7, x + 7, y + 7), fill="white", outline="#1F4E78", width=4)
    latest_x, latest_y = points[-1]
    draw.ellipse((latest_x - 11, latest_y - 11, latest_x + 11, latest_y + 11), fill="#D9A441", outline="#1F4E78", width=4)
    latest_text = f"{values[-1]:.1f}"
    latest_box = draw.textbbox((0, 0), latest_text, font=bold_font)
    draw.text((latest_x - (latest_box[2] - latest_box[0]) - 12, latest_y - 55), latest_text, fill="#1F4E78", font=bold_font)

    for x, label in zip(x_positions, labels):
        bbox = draw.textbbox((0, 0), label, font=label_font)
        draw.text((x - (bbox[2] - bbox[0]) / 2, plot_bottom + 20), label, fill="#404040", font=label_font)

    legend_y = 28
    draw.line((plot_left, legend_y + 14, plot_left + 52, legend_y + 14), fill="#1F4E78", width=7)
    draw.text((plot_left + 65, legend_y), "CBN Composite PMI", fill="#404040", font=legend_font)
    lx2 = plot_left + 300
    draw.line((lx2, legend_y + 14, lx2 + 52, legend_y + 14), fill="#C00000", width=4)
    draw.text((lx2 + 65, legend_y), "50-point threshold", fill="#404040", font=legend_font)

    y_title = "PMI index"
    title_box = Image.new("RGBA", (300, 70), (255, 255, 255, 0))
    title_draw = ImageDraw.Draw(title_box)
    title_draw.text((0, 0), y_title, fill="#404040", font=label_font)
    title_box = title_box.crop(title_box.getbbox()).rotate(90, expand=True)
    image.paste(title_box, (30, int((plot_top + plot_bottom - title_box.height) / 2)), title_box)

    change = values[-1] - values[-2]
    draw.text((plot_left, 65), f"Latest monthly change: {change:+.1f} points ({labels[-2]} to {labels[-1]})", fill="#0F6B67", font=small_font)
    px, py = points[-2]
    draw.text((px - 45, py + 18), f"{values[-2]:.1f}", fill="#1F4E78", font=small_font)
    note = "Above 50: improving conditions. Below 50: deteriorating conditions. Values are index levels."
    draw.text((plot_left, height - 80), note, fill="#595959", font=small_font)
    chart_path.parent.mkdir(parents=True, exist_ok=True)
    image.save(chart_path, format="PNG", optimize=True)


def create_inflation_chart(data: dict, chart_path: Path) -> None:
    history = data["history"]
    actual = []
    for observation_date in sorted(d for d in history if d < data["target"]):
        year_ago = datetime(observation_date.year - 1, observation_date.month, 1)
        if year_ago in history:
            actual.append((observation_date, (history[observation_date] / history[year_ago] - 1.0) * 100.0))
    actual = actual[-14:]
    if len(actual) < 2:
        raise RuntimeError("At least two annual inflation observations are required for the inflation chart.")

    width, height = 1548, 920
    left, right, top, bottom = 150, 55, 145, 190
    plot_left, plot_right = left, width - right
    plot_top, plot_bottom = top, height - bottom
    image = Image.new("RGB", (width, height), "#FBFCFE")
    draw = ImageDraw.Draw(image)
    font_dir = Path(r"C:\Windows\Fonts")

    def font(size: int, bold: bool = False):
        preferred = font_dir / ("calibrib.ttf" if bold else "calibri.ttf")
        fallback = font_dir / ("arialbd.ttf" if bold else "arial.ttf")
        if preferred.exists() or fallback.exists():
            return ImageFont.truetype(str(preferred if preferred.exists() else fallback), size)
        # Portable fallback for report verification outside Windows.
        linux_font = Path("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold
                          else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf")
        return ImageFont.truetype(str(linux_font), size)

    label_font = font(24)
    small_font = font(25)
    legend_font = font(27)
    bold_font = font(29, bold=True)
    title_font = font(38, bold=True)
    labels = [d.strftime("%b-%y") for d, _ in actual] + [data["target"].strftime("%b-%y")]
    actual_values = [v for _, v in actual]
    all_values = actual_values + [data["lower"], data["upper"], data["point"]]
    y_min = math.floor((min(all_values) - 0.7) * 2.0) / 2.0
    y_max = math.ceil((max(all_values) + 0.7) * 2.0) / 2.0
    if y_max <= y_min:
        y_max = y_min + 1.0

    def y_px(value: float) -> float:
        return plot_bottom - (value - y_min) / (y_max - y_min) * (plot_bottom - plot_top)

    x_positions = [plot_left + i * (plot_right - plot_left) / (len(labels) - 1) for i in range(len(labels))]
    forecast_left = (x_positions[-2] + x_positions[-1]) / 2
    draw.rectangle((forecast_left, plot_top, plot_right, plot_bottom), fill="#FFF1EF")
    draw.text((plot_right - 145, plot_top + 14), "FORECAST", fill="#A61B1B", font=small_font)
    tick = y_min
    while tick <= y_max + 1e-9:
        y = y_px(tick)
        draw.line((plot_left, y, plot_right, y), fill="#D9E2F3", width=2)
        text = f"{tick:.1f}"
        bbox = draw.textbbox((0, 0), text, font=small_font)
        draw.text((plot_left - 18 - (bbox[2] - bbox[0]), y - (bbox[3] - bbox[1]) / 2), text, fill="#595959", font=small_font)
        tick += 0.5
    draw.line((plot_left, plot_top, plot_left, plot_bottom), fill="#A6A6A6", width=3)
    draw.line((plot_left, plot_bottom, plot_right, plot_bottom), fill="#A6A6A6", width=3)

    actual_points = [(x_positions[i], y_px(v)) for i, v in enumerate(actual_values)]
    draw.polygon(
        actual_points + [(actual_points[-1][0], plot_bottom), (actual_points[0][0], plot_bottom)],
        fill="#EAF2F8",
    )
    draw.line(actual_points, fill="#1F4E78", width=7, joint="curve")
    for x, y in actual_points:
        draw.ellipse((x - 8, y - 8, x + 8, y + 8), fill="white", outline="#1F4E78", width=5)
    xf = x_positions[-1]
    lower_y, upper_y, point_y = y_px(data["lower"]), y_px(data["upper"]), y_px(data["point"])
    draw.line((xf, upper_y, xf, lower_y), fill="#C00000", width=12)
    draw.line((xf - 25, upper_y, xf + 25, upper_y), fill="#C00000", width=8)
    draw.line((xf - 25, lower_y, xf + 25, lower_y), fill="#C00000", width=8)
    draw.ellipse((xf - 10, point_y - 10, xf + 10, point_y + 10), fill="#C00000")
    forecast_text = f"{data['point']:.2f}% ({data['lower']:.2f}%–{data['upper']:.2f}%)"
    label_x = xf - 360
    label_y = point_y - 66
    label_box = draw.textbbox((label_x, label_y), forecast_text, font=bold_font)
    draw.rectangle((label_box[0] - 8, label_box[1] - 5, label_box[2] + 8, label_box[3] + 5), fill="white")
    draw.text((label_x, label_y), forecast_text, fill="#C00000", font=bold_font)

    for x, label in zip(x_positions, labels):
        bbox = draw.textbbox((0, 0), label, font=label_font)
        draw.text((x - (bbox[2] - bbox[0]) / 2, plot_bottom + 20), label, fill="#404040", font=label_font)
    draw.text((plot_left, 18), "Nigeria headline inflation", fill="#17365D", font=title_font)
    legend_y = 76
    draw.line((plot_left, legend_y + 14, plot_left + 52, legend_y + 14), fill="#1F4E78", width=7)
    draw.text((plot_left + 65, legend_y), "NBS actual", fill="#404040", font=legend_font)
    lx2 = plot_left + 300
    draw.line((lx2 + 26, legend_y - 2, lx2 + 26, legend_y + 31), fill="#C00000", width=10)
    draw.text((lx2 + 65, legend_y), "CAPE forecast range", fill="#404040", font=legend_font)
    y_title = "Headline inflation (%, year-on-year)"
    title_box = Image.new("RGBA", (800, 70), (255, 255, 255, 0))
    title_draw = ImageDraw.Draw(title_box)
    title_draw.text((0, 0), y_title, fill="#404040", font=label_font)
    title_box = title_box.crop(title_box.getbbox()).rotate(90, expand=True)
    image.paste(title_box, (25, int((plot_top + plot_bottom - title_box.height) / 2)), title_box)
    note = (
        f"Expected monthly inflation is {data['forecast_mom']:.2f}%; "
        f"range {data['forecast_mom_lower']:.2f}%–{data['forecast_mom_upper']:.2f}%."
    )
    draw.text((plot_left, height - 80), note, fill="#595959", font=small_font)
    chart_path.parent.mkdir(parents=True, exist_ok=True)
    image.save(chart_path, format="PNG", optimize=True)


def replace_paragraph_text(paragraph, text: str) -> None:
    template_rpr = None
    for run in paragraph.runs:
        if run._r.rPr is not None:
            template_rpr = copy.deepcopy(run._r.rPr)
            break
    for run in list(paragraph.runs):
        paragraph._p.remove(run._r)
    run = paragraph.add_run(text)
    if template_rpr is not None:
        if run._r.rPr is not None:
            run._r.remove(run._r.rPr)
        run._r.insert(0, template_rpr)


def replace_caption_keep_drawing(paragraph, text: str) -> None:
    template_rpr = None
    for run in paragraph.runs:
        has_drawing = bool(run._r.xpath(".//*[local-name()='drawing']"))
        if not has_drawing and run._r.rPr is not None and template_rpr is None:
            template_rpr = copy.deepcopy(run._r.rPr)
        if not has_drawing:
            paragraph._p.remove(run._r)
    run = paragraph.add_run(text)
    if template_rpr is not None:
        if run._r.rPr is not None:
            run._r.remove(run._r.rPr)
        run._r.insert(0, template_rpr)


def replace_source_paragraph(paragraph) -> None:
    lead_rpr = copy.deepcopy(paragraph.runs[0]._r.rPr) if paragraph.runs and paragraph.runs[0]._r.rPr is not None else None
    body_rpr = copy.deepcopy(paragraph.runs[1]._r.rPr) if len(paragraph.runs) > 1 and paragraph.runs[1]._r.rPr is not None else lead_rpr
    for run in list(paragraph.runs):
        paragraph._p.remove(run._r)
    lead = paragraph.add_run("Source: ")
    body = paragraph.add_run("National Bureau of Statistics (NBS) and CAPE Economic Research and Consulting")
    for run, rpr in ((lead, lead_rpr), (body, body_rpr)):
        if rpr is not None:
            if run._r.rPr is not None:
                run._r.remove(run._r.rPr)
            run._r.insert(0, copy.deepcopy(rpr))


def find_paragraph(doc: Document, startswith: str):
    for p in doc.paragraphs:
        if p.text.strip().startswith(startswith):
            return p
    raise RuntimeError(f"Newsletter paragraph not found: {startswith}")


def starts_with_flexible_spacing(text: str, prefix: str) -> bool:
    """Match newsletter text despite Word introducing, removing, or changing spaces."""
    normalized_text = re.sub(r"\s+", " ", text).strip().casefold()
    normalized_prefix = re.sub(r"\s+", " ", prefix).strip().casefold()
    if normalized_text.startswith(normalized_prefix):
        return True
    compact_text = re.sub(r"\s+", "", normalized_text)
    compact_prefix = re.sub(r"\s+", "", normalized_prefix)
    return compact_text.startswith(compact_prefix)


def find_paragraph_any(doc: Document, prefixes: tuple[str, ...]):
    for p in doc.paragraphs:
        text = p.text.strip()
        if any(starts_with_flexible_spacing(text, prefix) for prefix in prefixes):
            return p
    raise RuntimeError(f"Newsletter paragraph not found for any prefix: {prefixes}")


def find_last_paragraph_any(doc: Document, prefixes: tuple[str, ...]):
    matches = [
        p for p in doc.paragraphs
        if any(starts_with_flexible_spacing(p.text, prefix) for prefix in prefixes)
    ]
    if not matches:
        raise RuntimeError(f"Newsletter paragraph not found for any prefix: {prefixes}")
    return matches[-1]


def find_next_drawing_paragraph(doc: Document, paragraph):
    found = False
    for candidate in doc.paragraphs:
        if candidate._p is paragraph._p:
            found = True
            continue
        if found and candidate._p.xpath(".//*[local-name()='blip']"):
            return candidate
    raise RuntimeError("A drawing paragraph was not found after the figure caption.")


def replace_figure_image(doc: Document, paragraph, image_bytes: bytes) -> None:
    blips = paragraph._p.xpath(".//*[local-name()='blip']")
    if not blips:
        raise RuntimeError("Figure image relationship was not found.")
    rid = blips[0].get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
    image_part = doc.part.rels[rid].target_part
    image_part._blob = image_bytes
    # Replacement plots include their own labels; do not inherit template crops.
    for crop in paragraph._p.xpath(".//*[local-name()='srcRect']"):
        crop.getparent().remove(crop)


def resize_figure(paragraph, width_inches: float, height_inches: float) -> None:
    """Resize an inline or anchored figure while preserving its placement."""
    cx, cy = str(int(width_inches * 914400)), str(int(height_inches * 914400))
    for element in paragraph._p.xpath(".//*[local-name()='extent'] | .//*[local-name()='ext']"):
        if element.get("cx") is not None:
            element.set("cx", cx)
        if element.get("cy") is not None:
            element.set("cy", cy)


def replace_caption_with_inline_figure(paragraph, text: str, image_bytes: bytes, width_inches: float) -> None:
    """Replace a floating template figure with an inline image that participates in page flow."""
    replace_paragraph_text(paragraph, text)
    run = paragraph.add_run()
    run.add_break()
    run.add_picture(BytesIO(image_bytes), width=Inches(width_inches))


def apply_plot_overrides(
    doc: Document,
    plots_folder: Path | None,
    skip_figures: set[int] | None = None,
) -> list[str]:
    if plots_folder is None or not plots_folder.exists():
        return []
    captions_path = plots_folder / "captions.json"
    captions = json.loads(captions_path.read_text(encoding="utf-8")) if captions_path.exists() else {}
    applied = []
    skip_figures = skip_figures or set()
    for figure_number in range(1, 8):
        if figure_number in skip_figures:
            continue
        candidates = [
            plots_folder / f"Figure {figure_number}.png",
            plots_folder / f"Figure {figure_number}.jpg",
            plots_folder / f"Figure {figure_number}.jpeg",
        ]
        image_path = next((candidate for candidate in candidates if candidate.exists()), None)
        caption_text = captions.get(str(figure_number))
        if image_path is None and caption_text is None:
            continue
        caption_p = find_paragraph(doc, f"Figure {figure_number}:")
        if caption_text:
            replacement = str(caption_text).strip()
            if not replacement.startswith(f"Figure {figure_number}:"):
                replacement = f"Figure {figure_number}: {replacement}"
            replace_caption_keep_drawing(caption_p, replacement)
        if image_path is not None:
            drawing_p = caption_p if caption_p._p.xpath(".//*[local-name()='blip']") else find_next_drawing_paragraph(doc, caption_p)
            replace_figure_image(doc, drawing_p, image_path.read_bytes())
            applied.append(image_path.name)
    return applied



def cape_reporting_range(pi50_lower: float, pi50_upper: float) -> tuple[float, float]:
    """For public reporting only: round PI50 to two decimals, then adjust by ±10%."""
    low = Decimal(f"{pi50_lower:.2f}") * Decimal("1.10")
    high = Decimal(f"{pi50_upper:.2f}") * Decimal("0.90")
    return (float(low.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)),
            float(high.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)))


def gdp_comparison(point: float, last_actual: float, last_label: str) -> str:
    """Avoid spurious changes that disappear at normal publication precision."""
    diff = point - last_actual
    if abs(diff) < 0.05:
        return f"The projection is broadly in line with the {last_actual:.2f} per cent growth recorded in {last_label}, suggesting a broadly steady pace of expansion."
    if diff > 0:
        return f"This is {diff:.2f} percentage points above the {last_actual:.2f} per cent growth recorded in {last_label}, pointing to a modest pickup in growth."
    return f"This is {abs(diff):.2f} percentage points below the {last_actual:.2f} per cent growth recorded in {last_label}, indicating a modest easing in growth."


def refresh_verified_october_context(doc, report_month: str) -> None:
    """October 2026 editorial fact checks; do not overwrite later reports with dated facts."""
    if report_month != "2026-10":
        return
    changes = (
        ("Global growth remains moderate", "The IMF’s October 2026 World Economic Outlook projects global growth of 3.1 per cent in 2026 and 3.4 per cent in 2027. Investment in technology and resilient services support economic activity, but elevated energy prices, trade frictions and geopolitical uncertainty continue to constrain the outlook."),
        ("Major advanced-economy central banks", "Monetary policy across advanced economies remains cautious. In September 2026, the US Federal Reserve raised the federal funds target range to 3.75–4.00 per cent, while the Bank of England maintained Bank Rate at 3.75 per cent. The European Central Bank’s deposit rate stood at 2.50 per cent in early October. Policymakers remain attentive to persistent inflation and the economic effects of higher energy prices."),
        ("In Nigeria, the Monetary Policy Committee", "At its September 2026 meeting, the Central Bank of Nigeria’s Monetary Policy Committee set the Monetary Policy Rate at 23.0 per cent and retained the cash reserve requirement for deposit money banks at 45.0 per cent. Although the policy rate was reduced from July’s 26.5 per cent, financing conditions remain restrictive. Greater stability in food supply, energy and transport would help moderate inflation while supporting productive activity."),
        ("Food inflation rose to", "Food inflation eased to 19.57 per cent year-on-year in August 2026 from 20.31 per cent in July, with a month-on-month increase of 1.02 per cent. Core inflation fell to 13.29 per cent year-on-year, while the core price index declined by 0.06 per cent over the month. The improvement was encouraging, although food distribution, insecurity and energy costs remain significant risks."),
    )
    for prefix, text in changes:
        alternatives = {
            "Global growth remains moderate": ("Global growth remains moderate", "The IMF’s October 2026 World Economic Outlook"),
            "Major advanced-economy central banks": ("Major advanced-economy central banks", "Monetary policy across advanced economies"),
            "In Nigeria, the Monetary Policy Committee": ("In Nigeria, the Monetary Policy Committee", "At its September 2026 meeting"),
            "Food inflation rose to": ("Food inflation rose to", "Food inflation eased to"),
        }
        replace_paragraph_text(find_paragraph_any(doc, alternatives[prefix]), text)


def publish(
    template: Path,
    workbook: Path,
    output: Path,
    chart: Path,
    pmi_chart: Path,
    inflation_chart: Path,
    inflation: dict,
    gdp: dict,
    report_month: str,
    plots_folder: Path | None,
) -> dict:
    values = read_nowcast(workbook)
    # The launcher passes the CAPE reporting range itself (already the PI50 adjusted by 10%, as
    # shown and confirmed at the prompt), so it is published as given and not adjusted again.
    # The statistical PI50 comes from the model results. Only the public range is used in chart and bulletin.
    public_lower, public_upper = gdp["lower"], gdp["upper"]
    if not public_lower <= gdp["point"] <= public_upper:
        raise ValueError("CAPE reporting range excludes the GDP point forecast; review the bounds before publication.")
    values.update(
        {
            "release_mode": "CAPE forecast range",
            "point": gdp["point"],
            "pi50_lower": values.get("lower"),
            "pi50_upper": values.get("upper"),
            "lower": public_lower,
            "upper": public_upper,
        }
    )
    create_gdp_chart(values, chart)
    create_pmi_chart(values, pmi_chart)
    create_inflation_chart(inflation, inflation_chart)
    doc = Document(template)
    report_date = datetime.strptime(report_month, "%Y-%m")
    report_month_name = report_date.strftime("%B")
    report_year = report_date.year
    report_title = f"Cape Economic Performance and Prospect Bulletin – {report_month_name} {report_year}"
    replace_paragraph_text(
        find_paragraph_any(
            doc,
            (
                "Cape Economic Performance and Prospect Bulletin",
                "CAPE Economic Performance and Prospect Bulletin",
            ),
        ),
        report_title,
    )

    forecast_momentum = inflation_momentum(inflation["forecast_yoy_change"], future=True)
    opening_inflation = (
        f"CAPE expects headline inflation to reach {inflation['point']:.2f} per cent in "
        f"{inflation['target'].strftime('%B %Y')}, within a {inflation['lower']:.2f}–"
        f"{inflation['upper']:.2f} per cent forecast range. Relative to the "
        f"{inflation['latest_yoy']:.2f} per cent outturn in {inflation['latest'].strftime('%B')}, inflation "
        f"{forecast_momentum}. Consumer prices are expected to rise by {inflation['forecast_mom']:.2f} per cent "
        f"during the month."
    )
    replace_paragraph_text(
        find_paragraph_any(
            doc,
            ("CAPE expects headline inflation", "CAPE forecasts headline inflation", "CAPE projects headline inflation"),
        ),
        opening_inflation,
    )

    pmi_series = values["pmi_monthly"]
    pmi_latest = pmi_series[-1]
    pmi_previous = pmi_series[-2]
    consecutive_expansion = 0
    for observation in reversed(pmi_series):
        if observation["value"] > 50.0:
            consecutive_expansion += 1
        else:
            break
    if pmi_latest["value"] > pmi_previous["value"]:
        pmi_change = (
            f"rose to {pmi_latest['value']:.1f} points in {pmi_latest['date'].strftime('%B %Y')} from "
            f"{pmi_previous['value']:.1f} points in {pmi_previous['date'].strftime('%B')}"
        )
    elif pmi_latest["value"] < pmi_previous["value"]:
        pmi_change = (
            f"eased to {pmi_latest['value']:.1f} points in {pmi_latest['date'].strftime('%B %Y')} from "
            f"{pmi_previous['value']:.1f} points in {pmi_previous['date'].strftime('%B')}"
        )
    else:
        pmi_change = (
            f"was unchanged at {pmi_latest['value']:.1f} points in {pmi_latest['date'].strftime('%B %Y')}"
        )
    pmi_status = (
        f"This marks {consecutive_expansion} consecutive months above the 50.0 threshold and signals continued "
        f"private-sector expansion."
        if pmi_latest["value"] > 50.0
        else "The reading is below the 50.0 threshold and signals a deterioration in private-sector conditions."
    )
    pmi_narrative = (
        f"The Central Bank of Nigeria Composite Purchasing Managers’ Index (PMI) {pmi_change}. {pmi_status} "
        f"The reading indicates a broadening expansion in domestic activity, although elevated financing, energy, "
        f"transport and input costs continue to compress operating margins."
    )
    replace_paragraph_text(
        find_paragraph_any(
            doc,
            (
                "The Stanbic IBTC Purchasing Managers’ Index",
                "The Stanbic IBTC Nigeria Purchasing Managers’ Index",
                "The Central Bank of Nigeria Composite Purchasing Managers’ Index",
                "Nigeria’s Purchasing Managers’ Index",
            ),
        ),
        pmi_narrative,
    )

    pmi_caption_p = find_paragraph(doc, "Figure 2:")
    pmi_caption = (
        f"Figure 2: Trend of Nigeria’s CBN Composite PMI "
        f"({pmi_series[-14]['date'].strftime('%b %Y')}–{pmi_latest['date'].strftime('%b %Y')})"
    )
    replace_caption_keep_drawing(pmi_caption_p, pmi_caption)
    replace_figure_image(doc, find_next_drawing_paragraph(doc, pmi_caption_p), pmi_chart.read_bytes())
    pmi_source = find_paragraph_any(
        doc,
        ("Source: Stanbic-IBTC", "Source: Stanbic IBTC", "Source: Central Bank of Nigeria (CBN)"),
    )
    replace_paragraph_text(
        pmi_source,
        "Source: Central Bank of Nigeria (CBN) and CAPE Economic Research and Consulting",
    )

    latest_momentum = inflation_momentum(inflation["latest_yoy_change"])
    actual_inflation = (
        f"Headline inflation {latest_momentum} to {inflation['latest_yoy']:.2f} per cent year-on-year in "
        f"{inflation['latest'].strftime('%B %Y')}, from {inflation['previous_yoy']:.2f} per cent in "
        f"{inflation['previous'].strftime('%B')}; this represents "
        f"{percentage_point_change(inflation['latest_yoy_change'])}. The consumer price index rose by "
        f"{inflation['latest_mom']:.2f} per cent during the month. The movement in the annual rate therefore "
        f"reflects the interaction of current price pressures and base effects, rather than a fall in the price level."
    )
    if report_month == "2026-10" and inflation["latest"].strftime("%Y-%m") == "2026-08":
        # NBS August CPI release: 15.39% versus 15.43% in July, 0.71% MoM.
        # Archived rebased HCPI can differ in rounded published series.
        actual_inflation = (
            "Headline inflation eased to 15.39 per cent year-on-year in August 2026, "
            "from 15.43 per cent in July, a decline of 0.04 percentage points. "
            "Consumer prices nevertheless rose by 0.71 per cent during August. "
            "The slower annual inflation rate reflects easing price pressures, not a decline in the overall price level."
        )
    replace_paragraph_text(find_paragraph(doc, "Headline inflation"), actual_inflation)

    inflation_caption_p = find_paragraph(doc, "Figure 5:")
    chart_start = sorted(d for d in inflation["history"] if d < inflation["target"])[-14]
    inflation_caption = (
        f"Figure 5: Nigeria Headline Inflation and CAPE {inflation['target'].strftime('%B %Y')} Forecast "
        f"({chart_start:%b %Y}–{inflation['target']:%b %Y})"
    )
    replace_caption_keep_drawing(inflation_caption_p, inflation_caption)
    inflation_drawing = (
        inflation_caption_p
        if inflation_caption_p._p.xpath(".//*[local-name()='blip']")
        else find_next_drawing_paragraph(doc, inflation_caption_p)
    )
    replace_figure_image(doc, inflation_drawing, inflation_chart.read_bytes())

    inflation_outlook = (
        f"CAPE forecasts headline inflation at {inflation['point']:.2f} per cent in "
        f"{inflation['target'].strftime('%B %Y')}, within a {inflation['lower']:.2f}–"
        f"{inflation['upper']:.2f} per cent range. The annual rate {forecast_momentum} relative to "
        f"{inflation['latest'].strftime('%B')}. Consumer prices are expected to rise by "
        f"{inflation['forecast_mom']:.2f} per cent during the month. Across the forecast range, the monthly "
        f"increase would lie between {inflation['forecast_mom_lower']:.2f} and "
        f"{inflation['forecast_mom_upper']:.2f} per cent. "
        f"The balance of risks remains sensitive to food supply conditions, transport and energy costs, exchange-rate "
        f"pass-through, insecurity and the stance of monetary policy."
    )
    replace_paragraph_text(
        find_last_paragraph_any(doc, ("CAPE forecasts headline inflation", "CAPE expects headline inflation")),
        inflation_outlook,
    )

    pmi_outlook = (
        f"The CBN Composite PMI strengthened from {pmi_previous['value']:.1f} in "
        f"{pmi_previous['date'].strftime('%B')} to {pmi_latest['value']:.1f} in "
        f"{pmi_latest['date'].strftime('%B')}, remaining above the 50-point expansion threshold. The improvement "
        f"supports a firmer near-term activity profile, but the durability of the expansion will depend on reliable "
        f"power and transport services, access to working capital and continued foreign-exchange stability."
    )
    replace_paragraph_text(
        find_paragraph_any(
            doc,
            (
                "Forward-looking indicators remain cautiously positive",
                "The latest PMI readings point to a cautious improvement",
                "The improvement in business activity",
                "The CBN Composite PMI strengthened",
            ),
        ),
        pmi_outlook,
    )

    target_label = quarter_label(values["target_period"])
    latest_label = quarter_label(values["latest_period"])
    comparison = gdp_comparison(values["point"], values["latest_actual"], latest_label)
    highlight = (
        f"Nigeria’s real GDP grew by {values['latest_actual']:.2f} per cent year-on-year in {latest_label}. "
        f"CAPE forecasts growth of {values['point']:.2f} per cent for {target_label}, within a "
        f"{values['lower']:.2f}–{values['upper']:.2f} per cent range. {comparison}"
    )
    replace_paragraph_text(find_paragraph_any(doc, ("Nigeria’s real GDP growth strengthened", "Nigeria’s real GDP grew")), highlight)

    caption_p = find_paragraph(doc, "Figure 3:")
    caption_p.paragraph_format.page_break_before = True
    release_label = "Forecast and Expected Range"
    caption = f"Figure 3: Nigeria Real GDP Growth and CAPE {target_label} {release_label} (Q1 2024–{target_label}) in percent"
    replace_caption_with_inline_figure(caption_p, caption, chart.read_bytes(), 3.15)

    source_p = find_paragraph(doc, "Source: National Bureau of Statistics (NBS)")
    replace_source_paragraph(source_p)

    outlook = (
        f"CAPE forecasts real GDP growth of {values['point']:.2f} per cent year-on-year in {target_label}, "
        f"within an expected range of {values['lower']:.2f}–{values['upper']:.2f} per cent. "
        f"{comparison} The latest CBN business-activity readings indicate continued expansion, "
        f"while borrowing costs, energy and transport expenses, insecurity and pressure on household purchasing "
        f"power remain risks to the outlook."
    )
    replace_paragraph_text(
        find_paragraph_any(
            doc,
            (
                "Nigeria’s near-term output outlook",
                f"CAPE places {target_label} real GDP growth",
                "CAPE maintains an indicative range",
                "CAPE’s central estimate places real GDP growth",
                "CAPE expects real GDP growth",
                "CAPE nowcasts real GDP growth",
                "CAPE forecasts real GDP growth of",
            ),
        ),
        outlook,
    )

    conclusion_p = find_paragraph(doc, "Nigeria enters ")
    growth_sentence = (
        f"CAPE forecasts {values['point']:.2f} per cent growth in {target_label}, within an expected "
        f"range of {values['lower']:.2f}–{values['upper']:.2f} per cent."
    )
    conclusion = re.sub(r"^Nigeria enters [A-Za-z]+", f"Nigeria enters {report_month_name}", conclusion_p.text.strip())
    inflation_clause = (
        f"while headline inflation {forecast_momentum} to {inflation['point']:.2f} per cent in "
        f"{inflation['target'].strftime('%B')}, within a {inflation['lower']:.2f}–{inflation['upper']:.2f} per cent range."
    )
    conclusion = re.sub(
        r"while headline inflation is .*?(?= Sustaining| CAPE ERC expects| The outlook will depend|$)",
        inflation_clause,
        conclusion,
        flags=re.IGNORECASE,
    )
    growth_start = conclusion.find("Real GDP growth")
    continuation_markers = [" Stronger federation", " CAPE ERC expects", " The outlook will depend"]
    continuation_positions = [conclusion.find(marker, growth_start) for marker in continuation_markers if conclusion.find(marker, growth_start) >= 0]
    growth_end = min(continuation_positions) if continuation_positions else -1
    refreshed_growth = f"Real GDP growth accelerated to {values['latest_actual']:.2f} per cent in {latest_label}. {growth_sentence}"
    if growth_start >= 0 and growth_end > growth_start:
        conclusion = conclusion[:growth_start] + refreshed_growth + conclusion[growth_end:]
    replace_paragraph_text(conclusion_p, conclusion)

    refresh_verified_october_context(doc, report_month)

    # Figures 3 and 5 are generated from the GDP and inflation prompts and must
    # not be replaced by stale manually supplied plots in the Plots folder.
    plot_overrides = apply_plot_overrides(doc, plots_folder, skip_figures={3, 5})

    output.parent.mkdir(parents=True, exist_ok=True)
    doc.save(output)
    return {
        **values,
        "template": str(template),
        "workbook": str(workbook),
        "output": str(output),
        "chart": str(chart),
        "pmi_chart": str(pmi_chart),
        "inflation_chart": str(inflation_chart),
        "inflation": {key: value for key, value in inflation.items() if key != "history"},
        "report_month": report_month,
        "plot_overrides": plot_overrides,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Publish the Nigeria GDP nowcast into the CAPE economic newsletter.")
    parser.add_argument("--workbook", type=Path, required=True)
    parser.add_argument("--template", type=Path, default=DEFAULT_TEMPLATE)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    parser.add_argument("--chart", type=Path)
    parser.add_argument("--pmi-chart", type=Path)
    parser.add_argument("--inflation-chart", type=Path)
    parser.add_argument("--gdp-point", type=float, required=True)
    parser.add_argument("--gdp-lower", type=float, required=True)
    parser.add_argument("--gdp-upper", type=float, required=True)
    parser.add_argument("--inflation-source", type=Path, required=True)
    parser.add_argument("--inflation-month", required=True)
    parser.add_argument("--inflation-point", type=float, required=True)
    parser.add_argument("--inflation-lower", type=float, required=True)
    parser.add_argument("--inflation-upper", type=float, required=True)
    parser.add_argument("--report-month", default=datetime.now().strftime("%Y-%m"))
    parser.add_argument("--plots-folder", type=Path)
    args = parser.parse_args()
    chart = args.chart or args.workbook.parent / "Newsletter Assets" / "Nigeria GDP Nowcast Chart.png"
    pmi_chart = args.pmi_chart or args.workbook.parent / "Newsletter Assets" / "Nigeria PMI Trend Chart.png"
    inflation_chart = args.inflation_chart or args.workbook.parent / "Newsletter Assets" / "Nigeria Inflation Forecast Chart.png"
    template = args.template.resolve()
    output = args.output.resolve()
    if not template.exists() and output.exists():
        template = output
    plots_folder = args.plots_folder.resolve() if args.plots_folder else None
    inflation = read_inflation_forecast(
        args.inflation_source.resolve(), args.inflation_month,
        args.inflation_point, args.inflation_lower, args.inflation_upper,
    )
    if args.gdp_lower > args.gdp_upper or not args.gdp_lower <= args.gdp_point <= args.gdp_upper:
        parser.error("The expected GDP growth rate must fall within the lower and upper range bounds.")
    gdp = {"point": args.gdp_point, "lower": args.gdp_lower, "upper": args.gdp_upper}
    result = publish(
        template, args.workbook.resolve(), output, chart.resolve(), pmi_chart.resolve(),
        inflation_chart.resolve(), inflation, gdp, args.report_month, plots_folder,
    )
    print({
        "status": "newsletter published",
        "target": result["target_period"],
        "release_mode": result["release_mode"],
        "point": result["point"],
        "statistical_pi50": [result["pi50_lower"], result["pi50_upper"]],
        "cape_reporting_range": [result["lower"], result["upper"]],
        "inflation_point": result["inflation"]["point"],
        "inflation_range": [result["inflation"]["lower"], result["inflation"]["upper"]],
        "inflation_month_on_month": result["inflation"]["forecast_mom"],
        "report_month": result["report_month"],
        "plot_overrides": result["plot_overrides"],
        "output": result["output"],
    })


if __name__ == "__main__":
    main()
