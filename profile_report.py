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
    if matrix.empty or not np.isfinite(matrix.to_numpy(dtype=float)).any():
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
    return f"{x_text} at {draft_start:g}-<{draft_start + 1:g} m draft", share


def _draw_profile_distribution(
    pdf: canvas.Canvas,
    profile,
    x: float,
    top_y: float,
    width: float,
    title: str,
    x_axis_title: str,
    band_width: float,
    colour,
) -> None:
    """Draw a compact bar graph of propelling-hour share by operating band."""
    matrix = profile.percent.copy()
    pdf.setFillColor(DARK)
    pdf.setFont(FONT_BOLD, 7.8)
    pdf.drawString(x, top_y, title)
    if matrix.empty:
        pdf.setFont(FONT, 7.2)
        pdf.setFillColor(MID_GREY)
        pdf.drawString(x, top_y - 18, "No operating-profile data available")
        return

    totals = matrix.sum(axis=0).astype(float)
    active = np.where(totals.to_numpy() > 0)[0]
    if not len(active):
        pdf.setFont(FONT, 7.2)
        pdf.setFillColor(MID_GREY)
        pdf.drawString(x, top_y - 18, "No operating-profile data available")
        return

    totals = totals.iloc[active[0] : active[-1] + 1]
    values = totals.to_numpy(dtype=float)
    maximum = max(float(values.max()), 1.0)
    y_max = max(10.0, math.ceil(maximum / 10.0) * 10.0)
    plot_x = x + 28
    plot_y = top_y - 67
    plot_width = width - 34
    plot_height = 48

    pdf.setStrokeColor(colors.HexColor("#D9DEE7"))
    pdf.setFillColor(MID_GREY)
    pdf.setFont(FONT, 5.4)
    for tick in (0, y_max / 2, y_max):
        tick_y = plot_y + plot_height * tick / y_max
        pdf.line(plot_x, tick_y, plot_x + plot_width, tick_y)
        pdf.drawRightString(plot_x - 3, tick_y - 2, f"{tick:.0f}%")

    slot_width = plot_width / max(len(values), 1)
    bar_width = max(2.0, slot_width * 0.66)
    pdf.setFillColor(colour)
    for index, value in enumerate(values):
        bar_height = plot_height * max(value, 0.0) / y_max
        bar_x = plot_x + index * slot_width + (slot_width - bar_width) / 2
        pdf.rect(bar_x, plot_y, bar_width, bar_height, fill=1, stroke=0)

    labels = list(totals.index)
    label_step = max(1, math.ceil(len(labels) / 6))
    pdf.setFillColor(MID_GREY)
    pdf.setFont(FONT, 5.4)
    for index in range(0, len(labels), label_step):
        start = float(labels[index])
        label = f"{start:,.0f}" if band_width >= 100 else f"{start:g}"
        pdf.drawCentredString(plot_x + (index + 0.5) * slot_width, plot_y - 7, label)

    pdf.setFont(FONT_BOLD, 5.9)
    pdf.drawCentredString(plot_x + plot_width / 2, plot_y - 15, x_axis_title)
    pdf.saveState()
    pdf.translate(x + 7, plot_y + plot_height / 2)
    pdf.rotate(90)
    pdf.drawCentredString(0, 0, "Propelling hours (%)")
    pdf.restoreState()


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
    points: list[tuple[float, float]] = []
    for index, value in enumerate(values):
        if not math.isfinite(value):
            continue
        point_x = plot_x + (plot_width * index / max(count - 1, 1))
        point_y = plot_y + plot_height * (value - y_min) / (y_max - y_min)
        points.append((point_x, point_y))

    pdf.setStrokeColor(BLUE)
    pdf.setLineWidth(1.25)
    for first, second in zip(points, points[1:]):
        pdf.line(first[0], first[1], second[0], second[1])
    pdf.setFillColor(BLUE)
    for point_x, point_y in points:
        pdf.circle(point_x, point_y, 1.4, fill=1, stroke=0)

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
    pdf.setTitle(f"{vessel_name} Operating Profile and FOC Saving Assessment")
    pdf.setAuthor(prepared_by or "Vessel Operating Profile App")

    margin = 13 * mm
    content_width = width - 2 * margin

    pdf.setFillColor(NAVY)
    pdf.rect(0, height - 28 * mm, width, 28 * mm, fill=1, stroke=0)
    pdf.setFillColor(colors.white)
    pdf.setFont(FONT_BOLD, 18)
    pdf.drawString(margin, height - 13 * mm, "Vessel Operating Profile and FOC Saving Assessment")
    pdf.setFont(FONT, 8.5)
    pdf.drawString(
        margin,
        height - 20 * mm,
        "Marine performance summary generated from Noon, Departure and Arrival reports",
    )
    pdf.setFont(FONT_BOLD, 7.5)
    pdf.drawRightString(width - margin, height - 24 * mm, "A4 DECISION-SUPPORT REPORT")

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

    pdf.setFillColor(colors.HexColor("#FFF4D6"))
    pdf.setStrokeColor(colors.HexColor("#E6B94A"))
    callout_y = meta_y - 39
    pdf.roundRect(margin, callout_y, content_width, 22, 4, fill=1, stroke=1)
    pdf.setFillColor(AMBER)
    pdf.setFont(FONT_BOLD, 8)
    pdf.drawString(
        margin + 8,
        callout_y + 8,
        "Commercial scenario: the FOC saving percentage is user-defined and is not a verified post-retrofit performance result.",
    )

    total_hours = float(overall.get("total_hours", 0.0) or 0.0)
    analysis_days = total_hours / 24 if total_hours > 0 else float("nan")
    card_y = callout_y - 70
    card_gap = 6
    card_width = (content_width - 3 * card_gap) / 4
    card_values = [
        ("Report span", _number(analysis_days, 1, " days")),
        ("Noon report intervals", f"{int(noon_records):,}"),
        ("M/E propelling hours", _number(overall.get("propelling_hours"), 1, " h")),
        ("FOC saving assumption", _number(foc_saving_percent, 2, "%")),
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

    speed_band, speed_share = _dominant_band(speed_profile, 1.0, "kn")
    power_band, power_share = _dominant_band(power_profile, 1_000.0, "kW")

    section_y = card_y - 24
    _draw_section_title(pdf, "Operating profile", margin, section_y)
    operating_rows = [
        ("Mean STW while propelling", _number(overall.get("avg_speed_knots"), 2, " kn")),
        ("Maximum reported STW", _number(overall.get("max_noon_speed_knots"), 2, " kn")),
        ("Propelling time ratio", _number(overall.get("working_ratio_pct"), 1, "%")),
        ("Maximum reported M/E output", _number(overall.get("max_noon_me_output_kw"), 0, " kW")),
        ("Principal STW/mean-draft condition", f"{speed_band} ({speed_share:.1f}%)"),
        ("Principal M/E output/mean-draft condition", f"{power_band} ({power_share:.1f}%)"),
    ]
    table_bottom = _draw_two_column_rows(pdf, operating_rows, margin, section_y - 13, content_width)

    equivalent_fuel = float(fuel.get("total_vlsfo_equivalent_mt", 0.0) or 0.0)
    raw_fuel = float(fuel.get("total_raw_mt", 0.0) or 0.0)
    period_saving = equivalent_fuel * foc_saving_percent / 100
    period_cost_saving = period_saving * fuel_price
    annual_factor = (365 * 24 / total_hours) if total_hours > 0 else float("nan")
    annual_fuel_saving = period_saving * annual_factor
    annual_cost_saving = annual_fuel_saving * fuel_price

    section_y = table_bottom - 18
    _draw_section_title(pdf, "M/E fuel consumption and assumed saving", margin, section_y)
    fuel_rows = [
        ("VLSFO-equivalent M/E fuel", _number(equivalent_fuel, 2, " MT")),
        ("Assumed period FOC reduction", _number(period_saving, 2, " MT")),
        ("Projected annual FOC reduction", _number(annual_fuel_saving, 2, " MT/year")),
        ("Projected annual bunker-cost saving", f"US$ {_number(annual_cost_saving, 0)}/year"),
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
        commercial_rows = [
            ("Project CAPEX", f"US$ {_number(payback.get('capex_usd'), 0)}"),
            ("Simple payback period", payback_text),
            ("Charter assessment period", _number(payback.get("charter_duration_years"), 0, " years")),
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

    section_y = table_bottom - 18
    _draw_section_title(pdf, "Monthly operating trends", margin, section_y)
    monthly_chart_height = 52
    chart_top = section_y - 15
    _draw_monthly_line_chart(
        pdf, monthly, "working_ratio_pct", margin, chart_top, content_width,
        monthly_chart_height, "Monthly reported propelling ratio", "%",
    )
    chart_top -= monthly_chart_height + 2
    _draw_monthly_line_chart(
        pdf, monthly, "avg_sea_temp_excel", margin, chart_top, content_width,
        monthly_chart_height, "Monthly mean reported sea-water temperature", " C",
    )
    chart_top -= monthly_chart_height + 2
    _draw_monthly_line_chart(
        pdf, monthly, "avg_speed_knots", margin, chart_top, content_width,
        monthly_chart_height, "Monthly mean reported speed", " kn",
    )

    section_y = chart_top - monthly_chart_height - 7
    _draw_section_title(pdf, "Interpretation and limitations", margin, section_y)
    standard_comment = (
        "This report describes the vessel's observed operating envelope and applies the selected FOC saving "
        "assumption to M/E fuel consumed at sea after LCV normalisation to a VLSFO-equivalent basis. Annualised "
        "and payback values assume the reported operating profile remains representative. D/G, auxiliary boiler, "
        "cylinder oil and non-propelling fuel are excluded."
    )
    next_y = _draw_wrapped(
        pdf,
        standard_comment,
        margin,
        section_y - 17,
        content_width,
        size=7.8,
        leading=9.4,
        colour=DARK,
        max_lines=4,
    )
    if management_comment.strip():
        _draw_wrapped(
            pdf,
            f"Management comment: {management_comment.strip()}",
            margin,
            next_y - 2,
            content_width,
            font=FONT_ITALIC,
            size=7.7,
            leading=9.2,
            colour=MID_GREY,
            max_lines=2,
        )

    methodology_y = 42
    pdf.setFillColor(PALE_BLUE)
    pdf.setStrokeColor(LINE)
    pdf.roundRect(margin, methodology_y, content_width, 47, 4, fill=1, stroke=1)
    pdf.setFillColor(NAVY)
    pdf.setFont(FONT_BOLD, 7.4)
    pdf.drawString(margin + 8, methodology_y + 34, "METHOD SUMMARY")
    _draw_wrapped(
        pdf,
        "Noon-report intervals are screened and M/E propelling hours are allocated by STW/mean-draft and M/E output/mean-draft condition. "
        "M/E fuel at sea is normalised by lower calorific value (LCV) before the saving and commercial calculations are applied.",
        margin + 8,
        methodology_y + 23,
        content_width - 16,
        size=6.9,
        leading=8.0,
        colour=DARK,
        max_lines=3,
    )

    pdf.setFillColor(MID_GREY)
    pdf.setFont(FONT, 6.7)
    pdf.drawString(margin, 24, "Generated by Vessel Operating Profile and Commercial Assessment")
    pdf.drawRightString(width - margin, 24, "Page 1 of 1")

    pdf.showPage()
    pdf.save()
    return buffer.getvalue()
