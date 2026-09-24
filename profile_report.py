"""Generate a concise one-page A4 vessel operating-profile report."""
from __future__ import annotations

import io
import math
import re
from datetime import datetime
from pathlib import Path
from typing import Any

import numpy as np
import pandas as pd
import reportlab
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas


NAVY = colors.HexColor("#17365D")
BLUE = colors.HexColor("#2F75B5")
PALE_BLUE = colors.HexColor("#EAF2F8")
PALE_GREY = colors.HexColor("#F4F6F8")
MID_GREY = colors.HexColor("#667085")
DARK = colors.HexColor("#182230")
GREEN = colors.HexColor("#217A5B")
AMBER = colors.HexColor("#9A6700")
LINE = colors.HexColor("#D0D5DD")

FONT = "ProfileSans"
FONT_BOLD = "ProfileSans-Bold"
FONT_ITALIC = "ProfileSans-Italic"
_FONT_DIRECTORY = Path(reportlab.__file__).parent / "fonts"
pdfmetrics.registerFont(TTFont(FONT, str(_FONT_DIRECTORY / "Vera.ttf")))
pdfmetrics.registerFont(TTFont(FONT_BOLD, str(_FONT_DIRECTORY / "VeraBd.ttf")))
pdfmetrics.registerFont(TTFont(FONT_ITALIC, str(_FONT_DIRECTORY / "VeraIt.ttf")))


def safe_report_filename(vessel_name: str) -> str:
    safe_name = re.sub(r"[^A-Za-z0-9_-]+", "_", vessel_name.strip()).strip("_")
    return f"{safe_name or 'vessel'}_operating_profile_report.pdf"


def _number(value: Any, decimals: int = 1, suffix: str = "") -> str:
    try:
        number = float(value)
    except (TypeError, ValueError):
        return "Not available"
    if not math.isfinite(number):
        return "Not available"
    return f"{number:,.{decimals}f}{suffix}"


def _date(value: Any) -> str:
    if value is None or pd.isna(value):
        return "Not available"
    timestamp = pd.Timestamp(value)
    return timestamp.strftime("%d %b %Y")


def _wrap_text(text: str, font: str, size: float, width: float) -> list[str]:
    words = str(text).replace("\n", " \n ").split()
    lines: list[str] = []
    current = ""
    for word in words:
        if word == "\n":
            if current:
                lines.append(current)
                current = ""
            continue
        trial = word if not current else f"{current} {word}"
        if stringWidth(trial, font, size) <= width:
            current = trial
        else:
            if current:
                lines.append(current)
            current = word
    if current:
        lines.append(current)
    return lines


def _draw_wrapped(
    pdf: canvas.Canvas,
    text: str,
    x: float,
    y: float,
    width: float,
    font: str = FONT,
    size: float = 8.2,
    leading: float = 10.2,
    colour= DARK,
    max_lines: int | None = None,
) -> float:
    lines = _wrap_text(text, font, size, width)
    if max_lines is not None and len(lines) > max_lines:
        lines = lines[:max_lines]
        last = lines[-1]
        while last and stringWidth(last + "...", font, size) > width:
            last = last[:-1]
        lines[-1] = last.rstrip() + "..."
    pdf.setFont(font, size)
    pdf.setFillColor(colour)
    for line in lines:
        pdf.drawString(x, y, line)
        y -= leading
    return y


def _draw_section_title(pdf: canvas.Canvas, text: str, x: float, y: float) -> None:
    pdf.setFillColor(NAVY)
    pdf.setFont(FONT_BOLD, 10.5)
    pdf.drawString(x, y, text.upper())
    pdf.setStrokeColor(BLUE)
    pdf.setLineWidth(1.4)
    pdf.line(x, y - 4, A4[0] - x, y - 4)


def _draw_card(
    pdf: canvas.Canvas,
    x: float,
    y: float,
    width: float,
    height: float,
    label: str,
    value: str,
) -> None:
    pdf.setFillColor(PALE_GREY)
    pdf.setStrokeColor(LINE)
    pdf.roundRect(x, y, width, height, 5, fill=1, stroke=1)
    pdf.setFillColor(MID_GREY)
    pdf.setFont(FONT_BOLD, 7.4)
    pdf.drawString(x + 8, y + height - 15, label.upper())
    _draw_wrapped(
        pdf,
        value,
        x + 8,
        y + height - 34,
        width - 16,
        font=FONT_BOLD,
        size=11.2,
        leading=12,
        colour=DARK,
        max_lines=2,
    )


def _draw_two_column_rows(
    pdf: canvas.Canvas,
    rows: list[tuple[str, str]],
    x: float,
    top_y: float,
    width: float,
    row_height: float = 22,
) -> float:
    column_width = width / 2
    for index, (label, value) in enumerate(rows):
        column = index % 2
        row = index // 2
        cell_x = x + column * column_width
        cell_y = top_y - row * row_height
        pdf.setFillColor(PALE_GREY if row % 2 == 0 else colors.white)
        pdf.setStrokeColor(LINE)
        pdf.rect(cell_x, cell_y - row_height + 2, column_width, row_height, fill=1, stroke=1)
        pdf.setFillColor(MID_GREY)
        pdf.setFont(FONT, 7.2)
        pdf.drawString(cell_x + 6, cell_y - 8, label)
        _draw_wrapped(
            pdf,
            value,
            cell_x + 6,
            cell_y - 18,
            column_width - 12,
            font=FONT_BOLD,
            size=8.5,
            leading=9,
            colour=DARK,
            max_lines=1,
        )
    return top_y - math.ceil(len(rows) / 2) * row_height


def _dominant_band(profile, x_width: float, x_unit: str) -> tuple[str, float]:
    matrix = profile.percent
    if matrix.empty or not np.isfinite(matrix.to_numpy(dtype=float)).any() or not (matrix.to_numpy(dtype=float) > 0).any():
        return "Not available", 0.0
    values = matrix.to_numpy(dtype=float)
    flat_index = int(np.nanargmax(values))
    row_index, column_index = np.unravel_index(flat_index, values.shape)
    share = float(values[row_index, column_index])
    draft_start = float(matrix.index[row_index])
    x_start = float(matrix.columns[column_index])
    if x_width >= 100:
        x_text = f"{x_start:,.0f}-<{x_start + x_width:,.0f} {x_unit}"
    else:
        x_text = f"{x_start:g}-<{x_start + x_width:g} {x_unit}"
    return f"{x_text} / {draft_start:g}-<{draft_start + 1:g} m", share


def _draw_profile_heatmap(
    pdf: canvas.Canvas,
    profile,
    x: float,
    top_y: float,
    width: float,
    title: str,
    x_axis_title: str,
    band_width: float,
    common_max_pct: float,
) -> None:
    """Draw the complete fixed-bin speed/draught or power/draught matrix."""
    matrix = profile.percent.copy().apply(pd.to_numeric, errors="coerce")
    pdf.setFillColor(DARK)
    pdf.setFont(FONT_BOLD, 7.8)
    pdf.drawString(x, top_y, title)
    if matrix.empty or not (matrix.to_numpy(dtype=float) > 0).any():
        pdf.setFont(FONT, 7.2)
        pdf.setFillColor(MID_GREY)
        pdf.drawString(x, top_y - 30, "No eligible operating-profile data")
        return
    values = matrix.to_numpy(dtype=float)
    n_rows, n_columns = values.shape
    peak = np.unravel_index(int(np.nanargmax(values)), values.shape)
    plot_x = x + 24
    plot_top = top_y - 15
    plot_width = width - 28
    plot_height = 94
    cell_width = plot_width / n_columns
    cell_height = plot_height / n_rows
    empty = (250, 251, 253)
    pale = (221, 237, 247)
    dark = (31, 93, 157)
    for row in range(n_rows):
        for column in range(n_columns):
            value = values[row, column]
            if not np.isfinite(value) or value <= 0:
                rgb = empty
                ratio = 0.0
            else:
                ratio = min(value / max(common_max_pct, 1e-9), 1.0) ** 0.7
                rgb = tuple(int(a + (b - a) * ratio) for a, b in zip(pale, dark))
            cell_x = plot_x + column * cell_width
            cell_y = plot_top - (row + 1) * cell_height
            pdf.setFillColor(colors.Color(*(channel / 255 for channel in rgb)))
            pdf.setStrokeColor(colors.white)
            pdf.setLineWidth(0.35)
            pdf.rect(cell_x, cell_y, cell_width, cell_height, fill=1, stroke=1)
            if np.isfinite(value) and value > 0:
                # Label EVERY coloured cell; small nonzero shares must never
                # round to 0.0 or disappear from the printed report.
                if value < 0.001:
                    mantissa, exponent = f"{value:.1e}".split("e")
                    label = f"{mantissa.rstrip('0').rstrip('.')}e{int(exponent)}"
                elif value < 0.01:
                    label = f"{value:.3f}".removeprefix("0")
                elif value < 1:
                    label = f"{value:.2f}".removeprefix("0")
                else:
                    label = f"{value:.1f}"
                font = FONT_BOLD if (row, column) == peak else FONT
                label_width_at_one_point = stringWidth(label, font, 1.0)
                font_size = min(5.0, (cell_width - 0.6) / label_width_at_one_point)
                pdf.setFont(font, font_size)
                pdf.setFillColor(colors.white if ratio > 0.53 else NAVY)
                pdf.drawCentredString(
                    cell_x + cell_width / 2,
                    cell_y + cell_height / 2 - font_size * 0.34,
                    label,
                )
    # An amber frame marks the cell with the greatest share of eligible hours.
    pdf.setStrokeColor(colors.HexColor("#D28B00"))
    pdf.setLineWidth(1.8)
    pdf.rect(
        plot_x + peak[1] * cell_width - 0.3,
        plot_top - (peak[0] + 1) * cell_height - 0.3,
        cell_width + 0.6, cell_height + 0.6, fill=0, stroke=1,
    )

    pdf.setFillColor(MID_GREY)
    pdf.setFont(FONT, 5.1)
    for row, label in enumerate(matrix.index):
        pdf.drawRightString(plot_x - 3, plot_top - (row + 0.5) * cell_height - 2, str(label))
    label_step = max(1, math.ceil(n_columns / 6))
    for column in range(0, n_columns, label_step):
        label = float(matrix.columns[column])
        number = f"{label/1000:g}k" if band_width >= 1000 and label >= 1000 else f"{label:g}"
        pdf.drawCentredString(plot_x + (column + 0.5) * cell_width, plot_top - plot_height - 8, number)
    pdf.setFont(FONT_BOLD, 5.7)
    pdf.drawCentredString(plot_x + plot_width / 2, plot_top - plot_height - 17, x_axis_title)
    main_band, main_share = _dominant_band(profile, band_width, "kW" if band_width >= 100 else "kn")
    pdf.setFillColor(AMBER)
    pdf.setFont(FONT_BOLD, 6.0)
    main_label = f"Main: {main_band} ({main_share:.1f}%)"
    if stringWidth(main_label, FONT_BOLD, 6.0) > width:
        _draw_wrapped(pdf, main_label, x, top_y - 139, width, FONT_BOLD, 6.0, 7, AMBER, max_lines=1)
    else:
        pdf.drawString(x, top_y - 139, main_label)


def _draw_monthly_line_chart(
    pdf: canvas.Canvas,
    monthly: pd.DataFrame,
    column: str,
    x: float,
    top_y: float,
    width: float,
    height: float,
    title: str,
    unit: str,
) -> None:
    """Draw one compact monthly operating trend graph."""
    pdf.setFillColor(DARK)
    pdf.setFont(FONT_BOLD, 7.0)
    pdf.drawString(x, top_y, title)
    plot_x = x + 30
    plot_y = top_y - height + 11
    plot_width = width - 34
    plot_height = height - 20

    if monthly.empty or column not in monthly:
        pdf.setFont(FONT, 6.2)
        pdf.setFillColor(MID_GREY)
        pdf.drawString(plot_x, plot_y + plot_height / 2, "No monthly data available")
        return

    values = pd.to_numeric(monthly[column], errors="coerce").to_numpy(dtype=float)
    valid = np.isfinite(values)
    if not valid.any():
        pdf.setFont(FONT, 6.2)
        pdf.setFillColor(MID_GREY)
        pdf.drawString(plot_x, plot_y + plot_height / 2, "No monthly data available")
        return

    valid_values = values[valid]
    lower = float(valid_values.min())
    upper = float(valid_values.max())
    padding = max((upper - lower) * 0.15, 1.0 if unit == "%" else 0.5)
    y_min = max(0.0, lower - padding)
    y_max = upper + padding
    if y_max <= y_min:
        y_max = y_min + 1.0

    pdf.setStrokeColor(colors.HexColor("#D9DEE7"))
    pdf.setFillColor(MID_GREY)
    pdf.setFont(FONT, 5.2)
    for tick in (y_min, (y_min + y_max) / 2, y_max):
        tick_y = plot_y + plot_height * (tick - y_min) / (y_max - y_min)
        pdf.line(plot_x, tick_y, plot_x + plot_width, tick_y)
        pdf.drawRightString(plot_x - 3, tick_y - 2, f"{tick:.0f}{unit}")

    dates = pd.to_datetime(monthly["month"], errors="coerce")
    count = len(values)
    points: list[tuple[float, float] | None] = []
    for index, value in enumerate(values):
        if not math.isfinite(value):
            points.append(None)
            continue
        point_x = plot_x + (plot_width * index / max(count - 1, 1))
        point_y = plot_y + plot_height * (value - y_min) / (y_max - y_min)
        points.append((point_x, point_y))

    pdf.setStrokeColor(BLUE)
    pdf.setLineWidth(1.25)
    for first, second in zip(points, points[1:]):
        if first is not None and second is not None:
            pdf.line(first[0], first[1], second[0], second[1])
    pdf.setFillColor(BLUE)
    for point in points:
        if point is not None:
            pdf.circle(point[0], point[1], 1.4, fill=1, stroke=0)

    label_step = max(1, math.ceil(count / 6))
    pdf.setFillColor(MID_GREY)
    pdf.setFont(FONT, 5.1)
    for index in range(0, count, label_step):
        if pd.isna(dates.iloc[index]):
            continue
        label = dates.iloc[index].strftime("%Y/%m")
        point_x = plot_x + (plot_width * index / max(count - 1, 1))
        pdf.drawCentredString(point_x, plot_y - 7, label)


def build_a4_profile_report(
    *,
    vessel_name: str,
    imo_number: str,
    overall: dict[str, Any],
    noon_records: int,
    speed_profile,
    power_profile,
    monthly: pd.DataFrame,
    fuel: dict[str, Any],
    foc_saving_percent: float,
    fuel_price: float,
    payback: dict[str, Any] | None,
    prepared_by: str = "",
    management_comment: str = "",
) -> bytes:
    """Return a polished one-page A4 PDF containing the decision-level results."""
    buffer = io.BytesIO()
    width, height = A4
    pdf = canvas.Canvas(buffer, pagesize=A4, pageCompression=1)
    pdf.setTitle(f"{vessel_name} Operating Profile and Illustrative Retrofit Payback")
    pdf.setAuthor(prepared_by or "Vessel Operating Profile App")

    margin = 13 * mm
    content_width = width - 2 * margin

    pdf.setFillColor(NAVY)
    pdf.rect(0, height - 28 * mm, width, 28 * mm, fill=1, stroke=0)
    pdf.setFillColor(colors.white)
    pdf.setFont(FONT_BOLD, 16)
    pdf.drawString(margin, height - 13 * mm, "Operating Profile + Payback Calculation Report")
    pdf.setFont(FONT, 8.5)
    pdf.drawString(
        margin,
        height - 20 * mm,
        "Operating profile from noon, departure and arrival reports; commercial inputs are assumptions",
    )
    

    meta_y = height - 35 * mm
    pdf.setFillColor(DARK)
    pdf.setFont(FONT_BOLD, 12)
    pdf.drawString(margin, meta_y, vessel_name or "Vessel name not available")
    pdf.setFont(FONT, 8.3)
    pdf.setFillColor(MID_GREY)
    pdf.drawString(margin, meta_y - 13, f"IMO: {imo_number.strip() or 'Not provided'}")
    pdf.drawString(
        margin + 155,
        meta_y - 13,
        f"Data period: {_date(overall.get('data_start'))} to {_date(overall.get('data_end'))}",
    )
    pdf.drawRightString(
        width - margin,
        meta_y - 13,
        f"Prepared: {datetime.now().strftime('%d %b %Y')}",
    )
    if prepared_by.strip():
        pdf.drawRightString(width - margin, meta_y, f"Prepared by: {prepared_by.strip()[:55]}")

    total_hours = float(overall.get("total_hours", 0.0) or 0.0)
    analysis_days = total_hours / 24 if total_hours > 0 else float("nan")
    card_y = meta_y - 74
    card_gap = 6
    card_width = (content_width - 3 * card_gap) / 4
    card_values = [
        ("Period covered", _number(analysis_days, 1, " days")),
        ("Noon reports loaded", f"{int(noon_records):,}"),
        ("M/E propelling days", _number(overall.get("propelling_hours") / 24, 1, " days")),
        ("Assumed fuel saving", _number(foc_saving_percent, 2, "%")),
    ]
    for index, (label, value) in enumerate(card_values):
        _draw_card(
            pdf,
            margin + index * (card_width + card_gap),
            card_y,
            card_width,
            56,
            label,
            value,
        )

    section_y = card_y - 24
    _draw_section_title(pdf, "Operation Profile", margin, section_y)
    operating_rows = [
        ("Highest interval-average STW", _number(overall.get("max_noon_speed_knots"), 2, " kn")),
        ("Highest reported main-engine power", _number(overall.get("max_noon_me_output_kw"), 0, " kW")),
    ]
    table_bottom = _draw_two_column_rows(pdf, operating_rows, margin, section_y - 13, content_width)
    heatmap_top = table_bottom - 12
    heatmap_gap = 11
    heatmap_width = (content_width - heatmap_gap) / 2
    heatmaps = [
        (speed_profile, margin, heatmap_width, "Speed profile", "STW (kn)", 1.0),
        (power_profile, margin + heatmap_width + heatmap_gap, heatmap_width,
         "M/E Output Profile", "M/E power (kW)", 1000.0),
    ]
    peaks = [
        float(np.nanmax(item[0].percent.to_numpy(dtype=float)))
        for item in heatmaps if not item[0].percent.empty
        and np.isfinite(item[0].percent.to_numpy(dtype=float)).any()
    ]
    common_max_pct = max(peaks, default=1.0)
    for profile, plot_x, plot_width, title, axis_title, band_width in heatmaps:
        _draw_profile_heatmap(
            pdf, profile, plot_x, heatmap_top, plot_width, title,
            axis_title, band_width, common_max_pct,
        )
    caption = "Cell labels = % eligible propelling hours (leading zero omitted); amber = main condition."
    power_columns = power_profile.percent.columns
    max_reported_power = float(overall.get("max_noon_me_output_kw", float("nan")))
    if len(power_columns) and math.isfinite(max_reported_power):
        upper_power_bin = float(power_columns[-1]) + 1000
        if max_reported_power >= upper_power_bin:
            caption += (f" Power map ends at <{upper_power_bin:,.0f} kW; "
                        f"max {max_reported_power:,.0f} kW is outside.")
    pdf.setFillColor(MID_GREY)
    pdf.setFont(FONT, 6.2)
    if stringWidth(caption, FONT, 6.2) > content_width:
        raise ValueError("Heatmap range note is too long for the A4 report")
    pdf.drawString(margin, heatmap_top - 152, caption)
    # The monthly graphs start directly beneath the heatmap caption.
    section_y = heatmap_top - 165
    _draw_section_title(pdf, "Operation Condition", margin, section_y)
    monthly_chart_height = 56
    chart_top = section_y - 15
    _draw_monthly_line_chart(
        pdf, monthly, "working_ratio_pct", margin, chart_top, content_width,
        monthly_chart_height, "Monthly propelling hours / elapsed hours", "%",
    )
    chart_top -= monthly_chart_height + 11
    _draw_monthly_line_chart(
        pdf, monthly, "avg_sea_temp_excel", margin, chart_top, content_width,
        monthly_chart_height, "Monthly mean reported seawater temperature", " °C",
    )
    chart_top -= monthly_chart_height + 11
    _draw_monthly_line_chart(
        pdf, monthly, "avg_speed_knots", margin, chart_top, content_width,
        monthly_chart_height, "Monthly mean reported interval STW", " kn",
    )
    table_bottom = chart_top - monthly_chart_height

    equivalent_fuel = float(fuel.get("total_vlsfo_equivalent_mt", 0.0) or 0.0)
    raw_fuel = float(fuel.get("total_raw_mt", 0.0) or 0.0)
    period_saving = equivalent_fuel * foc_saving_percent / 100
    annual_factor = (365 * 24 / total_hours) if total_hours > 0 else float("nan")
    annual_cost_saving = period_saving * annual_factor * fuel_price
    if payback and payback.get("annual_gross_saving_usd") is not None:
        annual_cost_saving = float(payback["annual_gross_saving_usd"])

    section_y = table_bottom - 18
    _draw_section_title(pdf, "M/E fuel consumption and assumed saving", margin, section_y)
    fuel_rows = [
        ("Fuel Oil Consumption", _number(raw_fuel, 2, " t")),
        ("Fuel Oil Consumption (VLSFO Converted)", _number(equivalent_fuel, 2, " t")),
        ("Assumed FOC Saving", _number(period_saving, 2, " t equiv.")),
        ("FOC Save Cost Per Year", f"US$ {_number(annual_cost_saving, 0)}/year"),
    ]
    table_bottom = _draw_two_column_rows(pdf, fuel_rows, margin, section_y - 13, content_width)

    section_y = table_bottom - 18
    _draw_section_title(pdf, "Commercial view", margin, section_y)
    if payback:
        payback_value = payback.get("payback_years")
        payback_text = (
            _number(payback_value, 2, " years")
            if payback_value is not None and pd.notna(payback_value)
            else "No payback under current assumptions"
        )
        end_value = float(payback.get("net_surplus_usd", 0.0) or 0.0)
        unrecovered = float(payback.get("unrecovered_capex_usd", 0.0) or 0.0)
        duration = _number(payback.get("charter_duration_years"), 0, " years")
        commercial_rows = [
            ("Project CAPEX", f"US$ {_number(payback.get('capex_usd'), 0)}"),
            ("Payback period", payback_text),
            ("Charter period", duration),
            ("Payback within charter", str(payback.get("payback_within_charter", "Not available"))),
            ("Projected net benefit at charter end", f"US$ {_number(end_value, 0)}"),
            ("Unrecovered CAPEX at charter end", f"US$ {_number(unrecovered, 0)}"),
        ]
    else:
        commercial_rows = [
            ("Commercial calculation", "Not available"),
            ("Action required", "Enter valid CAPEX and saving assumptions in the Payback tab"),
        ]
    table_bottom = _draw_two_column_rows(pdf, commercial_rows, margin, section_y - 13, content_width)

    if table_bottom < 42:
        raise ValueError("A4 report content would overlap the footer")
    if management_comment.strip():
        _draw_wrapped(pdf, f"Comment: {management_comment.strip()}", margin,
                      39, content_width, font=FONT_ITALIC, size=6.5,
                      leading=7.4, colour=MID_GREY, max_lines=1)

    pdf.setFillColor(MID_GREY)
    pdf.setFont(FONT, 6.7)
    pdf.drawString(margin, 24, "Generated by Vessel Operating Profile and Commercial Assessment")
    pdf.drawRightString(width - margin, 24, "Page 1 of 1")

    pdf.showPage()
    pdf.save()
    return buffer.getvalue()
