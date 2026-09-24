from __future__ import annotations

import io
import math

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from profile_processing import (
    EXCEL_POWER_EDGES,
    EXCEL_SPEED_EDGES,
    REPORT_NAMES,
    ParsedReport,
    build_excel_data_sum,
    excel_overall_summary,
    fuel_consumption_summary,
    make_excel_profile,
    mapping_table,
    monthly_summary_excel,
    parse_all_report_files,
    profile_segments_from_data_sum,
    profile_with_totals,
    sea_temperature_audit,
    validate_vessel_consistency,
    VesselValidationError,
)
from profile_report import build_a4_profile_report, safe_report_filename


st.set_page_config(page_title="Vessel Operating Profile and Payback Calculation", layout="wide")


def uploaded_bytes(uploaded_file) -> bytes:
    uploaded_file.seek(0)
    return uploaded_file.read()


@st.cache_data(show_spinner=False)
def parse_cached(content: bytes):
    return parse_all_report_files(content)


def report_card(uploaded_file, report_type: str) -> ParsedReport | None:
    if uploaded_file is None:
        st.info(f"Upload a {REPORT_NAMES[report_type]} report.")
        return None
    reports, errors = parse_cached(uploaded_bytes(uploaded_file))
    if report_type in errors:
        st.error(errors[report_type])
        return None
    report = reports[report_type]

    st.success(
        f"Detected {REPORT_NAMES[report_type]} | sheet: {report.sheet_name} | "
        f"confidence: {report.confidence:.0%}"
    )
    if report.missing:
        st.error("This report cannot be processed until its missing required headers are restored.")
    for warning in report.warnings:
        st.warning(warning)
    with st.expander("Show detected column mapping"):
        st.dataframe(mapping_table(report), hide_index=True, use_container_width=True)
    return report


def heatmap(profile, title: str, x_title: str, chart_key: str):
    if profile.percent.empty:
        st.warning(f"No valid rows are available for the {title} profile.")
        return
    values = profile.percent.values
    text = [[f"{value:.2f}%" if value > 0 else "" for value in row] for row in values]
    figure = go.Figure(
        data=go.Heatmap(
            z=values,
            x=profile.percent.columns,
            y=profile.percent.index,
            colorscale=[[0, "#ffffff"], [0.25, "#fee2e2"], [1, "#b91c1c"]],
            colorbar={"title": "% of hours"},
            hovertemplate=(
                f"{x_title}: %{{x}}<br>Draught band start: %{{y}} m<br>"
                "Share of eligible propelling hours: %{z:.3f}%<extra></extra>"
            ),
            text=text,
            texttemplate="%{text}" if values.size <= 300 else None,
        )
    )
    figure.update_layout(
        title=title,
        xaxis_title=x_title,
        yaxis_title="Draught band start [m]",
        height=max(430, 34 * len(profile.percent.index)),
        margin={"l": 20, "r": 20, "t": 60, "b": 20},
    )
    figure.update_yaxes(autorange="reversed")
    st.plotly_chart(figure, use_container_width=True, key=chart_key)


def dataframe_csv(frame: pd.DataFrame, include_index: bool = False) -> bytes:
    return frame.to_csv(index=include_index).encode("utf-8")


def _profile_percent_text(value) -> str:
    if pd.isna(value):
        return ""
    return "-" if abs(float(value)) < 0.0005 else f"{float(value):.3f}%"


def styled_profile_table(frame: pd.DataFrame):
    """Apply compact, dependency-free colouring while keeping values numeric."""
    table = frame.copy()
    table.index.name = "Draught band start [m]"
    body_rows = [index for index in table.index if index != "Total"]
    body_columns = [column for column in table.columns if column != "Total"]
    body_max = float(table.loc[body_rows, body_columns].max().max()) if body_rows else 0.0

    def cell_colours(series: pd.Series) -> list[str]:
        styles: list[str] = []
        for value in series:
            number = float(value) if pd.notna(value) else 0.0
            if number <= 0 or body_max <= 0:
                styles.append("background-color:#f8fafc;color:#94a3b8;text-align:center")
                continue
            strength = min(number / body_max, 1.0)
            red = int(254 - 69 * strength)
            green = int(242 - 214 * strength)
            blue = int(242 - 214 * strength)
            text_colour = "#ffffff" if strength >= 0.58 else "#7f1d1d"
            styles.append(
                f"background-color:rgb({red},{green},{blue});color:{text_colour};"
                "font-weight:600;text-align:center"
            )
        return styles

    styler = table.style.format(_profile_percent_text)
    if body_rows and body_columns:
        styler = styler.apply(
            cell_colours,
            axis=0,
            subset=pd.IndexSlice[body_rows, body_columns],
        )
    if "Total" in table.columns:
        styler = styler.set_properties(
            subset=pd.IndexSlice[:, ["Total"]],
            **{"background-color": "#e2e8f0", "font-weight": "700", "color": "#0f172a"},
        )
    if "Total" in table.index:
        styler = styler.set_properties(
            subset=pd.IndexSlice[["Total"], :],
            **{"background-color": "#334155", "font-weight": "700", "color": "#ffffff"},
        )
    return styler


def render_readable_profile_table(profile):
    """Show the complete operating matrix without duplicate band views."""
    table = profile_with_totals(profile.percent)
    table.index.name = "Draught band start [m]"
    st.caption(
        "Complete matrix in one view; scroll horizontally for later bands. "
        "Zero cells are shown as -. Darker red means a larger share of total propelling hours."
    )
    st.dataframe(
        styled_profile_table(table),
        use_container_width=True,
        height=460,
    )
    return table


def excel_monthly_display(monthly: pd.DataFrame) -> pd.DataFrame:
    """Create the Profile-sheet table that supplies all three line graphs."""
    return pd.DataFrame(
        {
            "YEAR": monthly["month"].dt.year,
            "MONTH": monthly["month"].dt.month,
            "Period Start": monthly["data_start"].dt.strftime("%d/%m/%Y"),
            "Period End": monthly["data_end"].dt.strftime("%d/%m/%Y"),
            "Elapsed Time [h]": monthly["available_hours"],
            "Reported M/E Propelling Hours [h]": monthly["propelling_hours"],
            "Propelling Share of Elapsed Time [%]": monthly["working_ratio_pct"],
            "Mean Reported Seawater Temperature [Â°C]": monthly["avg_sea_temp_excel"],
            "Mean Reported Interval STW [kn]": monthly["avg_speed_knots"],
        }
    )


def show_excel_monthly_table(monthly: pd.DataFrame) -> pd.DataFrame:
    table = excel_monthly_display(monthly)
    st.dataframe(
        table.style.format(
            {
                "Elapsed Time [h]": "{:.1f}",
                "Reported M/E Propelling Hours [h]": "{:.1f}",
                "Propelling Share of Elapsed Time [%]": "{:.0f}%",
                "Mean Reported Seawater Temperature [Â°C]": "{:.6f}",
                "Mean Reported Interval STW [kn]": "{:.5f}",
            },
            na_rep="Needs review",
        ),
        hide_index=True,
        use_container_width=True,
    )
    return table


def plot_excel_monthly_graphs(table: pd.DataFrame, key_prefix: str):
    """Plot the three final table columns directly, without recalculation."""
    chart_data = table.copy()
    chart_data["Period"] = chart_data.apply(
        lambda row: f"{int(row['YEAR'])}/{int(row['MONTH'])}", axis=1
    )
    chart_specs = [
        (
            "Propelling Share of Elapsed Time [%]",
            "Monthly Propelling Share of Elapsed Time",
            "Propelling share [%]",
        ),
        (
            "Mean Reported Seawater Temperature [Â°C]",
            "Monthly Mean Reported Seawater Temperature",
            "Mean reported seawater temperature [Â°C]",
        ),
        (
            "Mean Reported Interval STW [kn]",
            "Monthly Mean Reported Interval STW",
            "Mean reported STW [kn]",
        ),
    ]
    for column, title, y_title in chart_specs:
        figure = px.line(
            chart_data,
            x="Period",
            y=column,
            markers=False,
            title=title,
        )
        figure.update_layout(
            height=260,
            xaxis_title=None,
            yaxis_title=y_title,
            margin={"l": 20, "r": 20, "t": 50, "b": 20},
        )
        st.plotly_chart(
            figure,
            use_container_width=True,
            key=f"{key_prefix}-{column}-table-chart",
        )


def excel_data_sum_display(data_sum: pd.DataFrame, imo_number: str) -> pd.DataFrame:
    """Expose the internal Data_sum in the same column order as the workbook."""
    return pd.DataFrame(
        {
            "IMO Number": imo_number,
            "Vessel": data_sum["vessel"],
            "Time(Noon/SOP/EOP)": data_sum["timestamp"],
            "Duration [h]": data_sum["duration_hours"],
            "Reported STW [kn]": data_sum["speed_knots"],
            "Active Midship Draught [m]": data_sum["draft_m"],
            "Reported Seawater Temperature [Â°C]": data_sum["data_sum_sea_temp"],
            "YEAR": data_sum["year"],
            "MONTH": data_sum["month"],
            "DAY": data_sum["day"],
            "HOUR": data_sum["hour"],
            "MINUTE": data_sum["minute"],
            "Reported Duration [days]": data_sum["duration_days"],
            "M/E Power [kW]": data_sum["me_output_kw"],
            "Source": data_sum["source"],
        }
    )


def render_excel_profile_details(
    profile,
    profile_name: str,
    vessel_name: str,
    imo_number: str,
    overall: dict,
):
    st.divider()
    st.subheader(f"{profile_name} summary")
    vessel_summary = pd.DataFrame(
        [
            {
                "Vessel Name": vessel_name,
                "IMO No.": imo_number or "Not provided",
                "M/E Propelling Duration [days]": overall["propelling_hours"] / 24,
                "Profile Duration [days]": profile.total_hours / 24,
            }
        ]
    )
    st.dataframe(
        vessel_summary.style.format(
            {"M/E Propelling Duration [days]": "{:.7f}", "Profile Duration [days]": "{:.7f}"}
        ),
        hide_index=True,
        use_container_width=True,
    )
    is_power_profile = "M/E Power" in profile_name
    if is_power_profile:
        st.metric(
            "Highest Reported M/E Power",
            f"{overall['max_noon_me_output_kw']:,.0f} kW",
        )
        st.caption(
            "Highest power is taken only from the uploaded Noon report's M/E output column."
        )
    else:
        st.metric(
            "Highest Reported Interval-Average STW",
            f"{overall['max_noon_speed_knots']:,.2f} kn",
        )
        st.caption(
            "Highest STW is taken only from the uploaded Noon report's average-speed column."
        )

    overall_row = {
        "Start Year": overall["year"],
        "Period Start": overall["data_start"].strftime("%d/%m/%Y"),
        "Period End": overall["data_end"].strftime("%d/%m/%Y"),
        "Elapsed Time [h]": overall["total_hours"],
        "Reported M/E Propelling Hours [h]": overall["propelling_hours"],
        "Propelling Share of Elapsed Time [%]": overall["working_ratio_pct"],
        "Mean Reported Seawater Temperature [Â°C]": overall["avg_sea_temp_excel"],
        "Mean Reported Interval STW [kn]": overall["avg_speed_knots"],
    }
    if is_power_profile:
        overall_row["Highest Reported M/E Power [kW]"] = overall["max_noon_me_output_kw"]
    else:
        overall_row["Highest Reported Interval-Average STW [kn]"] = overall["max_noon_speed_knots"]
    overall_table = pd.DataFrame([overall_row])
    overall_formats = {
        "Elapsed Time [h]": "{:.0f}",
        "Reported M/E Propelling Hours [h]": "{:.1f}",
        "Propelling Share of Elapsed Time [%]": "{:.2f}%",
        "Mean Reported Seawater Temperature [Â°C]": "{:.6f}",
        "Mean Reported Interval STW [kn]": "{:.5f}",
    }
    if is_power_profile:
        overall_formats["Highest Reported M/E Power [kW]"] = "{:,.0f}"
    else:
        overall_formats["Highest Reported Interval-Average STW [kn]"] = "{:.2f}"
    st.dataframe(
        overall_table.style.format(overall_formats),
        hide_index=True,
        use_container_width=True,
    )
    st.caption(
        "The mean seawater temperature is calculated from the internally created Data_sum "
        "'Sea Water Temp.' column. "
        "The source is located from the uploaded two-row title 'Sea Water temperature at noon', "
        "regardless of its Excel column position."
    )


def render_fuel_summary(fuel: dict, foc_saving_percent: float, fuel_price: float):
    """Render M/E fuel consumption and the assumed FOC saving."""
    st.subheader("M/E Fuel Consumption and Assumed FOC Saving")
    grades = list(fuel["total_by_grade"])
    fuel_rows = []
    for report_name, grade_values, equivalent in (
        ("Noon", fuel["noon_by_grade"], fuel["noon_vlsfo_equivalent_mt"]),
        ("Arrival", fuel["arrival_by_grade"], fuel["arrival_vlsfo_equivalent_mt"]),
    ):
        row = {"Report": report_name}
        row.update({f"{grade} [t]": grade_values[grade] for grade in grades})
        row["Reported M/E fuel, all grades [t]"] = sum(grade_values.values())
        row["VLSFO-energy-equivalent M/E fuel [t]"] = equivalent
        fuel_rows.append(row)
    fuel_table = pd.DataFrame(fuel_rows)
    fuel_formats = {
        column: "{:,.3f}" for column in fuel_table.columns if column != "Report"
    }
    st.dataframe(
        fuel_table.style.format(fuel_formats),
        hide_index=True,
        use_container_width=True,
    )
    with st.expander("Show fuel-grade conversion to VLSFO equivalent"):
        conversion_table = pd.DataFrame(
            [
                {
                    "Fuel grade": grade,
                    "Reported fuel [t]": fuel["total_by_grade"][grade],
                    "LCV [MJ/kg]": fuel["lcv_mj_per_kg"][grade],
                    "Conversion factor": fuel["conversion_factor"][grade],
                    "VLSFO-energy-equivalent fuel [t]": fuel["equivalent_by_grade"][grade],
                }
                for grade in grades
            ]
        )
        st.dataframe(
            conversion_table.style.format(
                {
                    "Reported fuel [t]": "{:,.3f}",
                    "LCV [MJ/kg]": "{:.1f}",
                    "Conversion factor": "{:.6f}",
                    "VLSFO-energy-equivalent fuel [t]": "{:,.3f}",
                }
            ),
            hide_index=True,
            use_container_width=True,
        )
        st.code("VLSFO-equivalent fuel [t] = reported fuel [t] Ã— fuel LCV / 40.5")

    equivalent_consumption = fuel["total_vlsfo_equivalent_mt"]
    saving_rate = foc_saving_percent / 100
    foc_saving_table = pd.DataFrame(
        [
            {
                "Item": "Total M/E fuel",
                "Common basis": "VLSFO equivalent",
                "Reported M/E fuel, all grades [t]": fuel["total_raw_mt"],
                "M/E fuel, VLSFO-energy equivalent [t]": equivalent_consumption,
                "Assumed FOC Saving [%]": foc_saving_percent,
                "Assumed FOC saving [t VLSFO-eq.]": equivalent_consumption * saving_rate,
                "VLSFO reference price [US$/t]": fuel_price,
                "Estimated fuel-cost saving for period [US$]": equivalent_consumption * saving_rate * fuel_price,
            },
        ]
    )
    st.dataframe(
        foc_saving_table.style.format(
            {
                "Reported M/E fuel, all grades [t]": "{:,.3f}",
                "M/E fuel, VLSFO-energy equivalent [t]": "{:,.3f}",
                "Assumed FOC Saving [%]": "{:.1f}%",
                "Assumed FOC saving [t VLSFO-eq.]": "{:,.3f}",
                "VLSFO reference price [US$/t]": "{:,.2f}",
                "Estimated fuel-cost saving for period [US$]": "{:,.2f}",
            }
        ),
        hide_index=True,
        use_container_width=True,
    )
    st.caption(
        "Includes M/E steaming consumption from Noon and Arrival reports. "
        "D/G, boiler, cylinder oil and stopping-condition fuel remain excluded, matching the Profile scope."
    )


def build_payback_analysis(
    capex_usd: float,
    annual_gross_fuel_saving_usd: float,
    annual_additional_opex_usd: float,
    annual_avoided_co2_cost_usd: float,
    charter_duration_years: int,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Build payback and cumulative cash-flow tables over the charter period."""
    scenarios: list[tuple[str, int | None]] = [
        ("Fuel-saving only (no avoided levy benefit)", None)
    ]
    if annual_avoided_co2_cost_usd > 0:
        scenarios.extend(
            [
                ("Avoided CO2 levy benefit after 1-year delay", 2),
                ("Avoided CO2 levy benefit after 2-year delay", 3),
            ]
        )

    cashflow_rows: list[dict] = []
    summary_rows: list[dict] = []

    for scenario_name, levy_start_year in scenarios:
        cumulative = -capex_usd
        cashflow_rows.append(
            {
                "Scenario": scenario_name,
                "Year": 0,
                "Gross fuel saving [US$]": 0.0,
                "Avoided CO2 levy benefit [US$]": 0.0,
                "Additional OPEX [US$]": 0.0,
                "Net cash flow [US$]": 0.0,
                "Cumulative cash flow [US$]": cumulative,
            }
        )

        payback_years: float | None = None
        charter_net_benefit = 0.0

        # Use a long horizon so payback can still be reported when it occurs
        # after the selected charter period.
        calculation_horizon = max(charter_duration_years, 200)
        for year in range(1, calculation_horizon + 1):
            avoided_co2 = (
                annual_avoided_co2_cost_usd
                if levy_start_year is not None and year >= levy_start_year
                else 0.0
            )
            net_cash_flow = (
                annual_gross_fuel_saving_usd
                + avoided_co2
                - annual_additional_opex_usd
            )
            previous_cumulative = cumulative
            cumulative += net_cash_flow

            if (
                payback_years is None
                and net_cash_flow > 0
                and previous_cumulative < 0 <= cumulative
            ):
                payback_years = (year - 1) + (-previous_cumulative / net_cash_flow)

            if year <= charter_duration_years:
                charter_net_benefit += net_cash_flow
                cashflow_rows.append(
                    {
                        "Scenario": scenario_name,
                        "Year": year,
                        "Gross fuel saving [US$]": annual_gross_fuel_saving_usd,
                        "Avoided CO2 levy benefit [US$]": avoided_co2,
                        "Additional OPEX [US$]": annual_additional_opex_usd,
                        "Net cash flow [US$]": net_cash_flow,
                        "Cumulative cash flow [US$]": cumulative,
                    }
                )

        year_one_net_saving = annual_gross_fuel_saving_usd - annual_additional_opex_usd
        charter_end_cash_flow = charter_net_benefit - capex_usd
        payback_within_charter = (
            payback_years is not None and payback_years <= charter_duration_years
        )
        summary_rows.append(
            {
                "Scenario": scenario_name,
                "First-year net benefit [US$]": year_one_net_saving,
                "Payback period [years]": payback_years,
                "Payback within charter": "Yes" if payback_within_charter else "No",
                "Net surplus at charter end [US$]": max(charter_end_cash_flow, 0.0),
                "Unrecovered CAPEX at charter end [US$]": max(-charter_end_cash_flow, 0.0),
            }
        )

    return pd.DataFrame(cashflow_rows), pd.DataFrame(summary_rows)


def render_payback_analysis(
    overall: dict,
    fuel: dict,
    foc_saving_percent: float,
    fuel_price: float,
):
    """Render the payback tab from the existing Profile fuel result."""
    st.subheader("Payback Calculation and Charter Outcome")
    st.caption(
        "The app annualises the assumed FOC saving calculated over the uploaded report period. "
        "CAPEX and other commercial assumptions must be entered manually."
    )

    analysis_hours = float(overall.get("total_hours", float("nan")))
    equivalent_consumption_mt = float(fuel.get("total_vlsfo_equivalent_mt", 0.0))

    if not math.isfinite(analysis_hours) or analysis_hours <= 0:
        st.error("Payback cannot be calculated because the analysis duration is invalid.")
        return
    if not 0 <= foc_saving_percent <= 100:
        st.error("FOC saving assumption must be between 0% and 100%.")
        return
    if fuel_price < 0:
        st.error("Fuel price cannot be negative.")
        return

    analysis_days = analysis_hours / 24
    period_fuel_saving_mt = equivalent_consumption_mt * foc_saving_percent / 100
    annualisation_factor = (365 * 24) / analysis_hours
    calculated_annual_fuel_saving_mt = period_fuel_saving_mt * annualisation_factor

    basis_columns = st.columns(3)
    basis_columns[0].metric("Period covered", f"{analysis_days:,.1f} days")
    basis_columns[1].metric("Assumed FOC saving for period", f"{period_fuel_saving_mt:,.3f} t")
    basis_columns[2].metric(
        "Illustrative annual FOC saving",
        f"{calculated_annual_fuel_saving_mt:,.3f} t/year",
    )

    if analysis_days < 180:
        st.warning(
            "The uploaded period is shorter than 180 days. Annualising a short period can produce "
            "an unstable payback estimate, especially if vessel operations are seasonal."
        )

    st.markdown("**Financial assumptions**")
    capex_usd = st.number_input(
        "Project CAPEX [US$]",
        min_value=0.0,
        value=331_800.0,
        step=1_000.0,
        key="payback-capex-usd",
    )

    option_columns = st.columns(3)
    with option_columns[0]:
        charter_duration_years = st.number_input(
            "Charter duration [years]",
            min_value=1,
            max_value=50,
            value=10,
            step=1,
            key="payback-charter-duration",
            help=(
                "Enter the period during which the investor receives the fuel-saving benefit. "
                "Savings after the charter ends are not counted."
            ),
        )
    with option_columns[1]:
        annual_additional_opex_usd = st.number_input(
            "Additional annual OPEX [US$]",
            min_value=0.0,
            value=0.0,
            step=1_000.0,
            key="payback-annual-opex",
            help="Extra yearly maintenance, servicing or operating cost caused by the project.",
        )
    with option_columns[2]:
        use_manual_saving = st.checkbox(
            "Override annual fuel saving",
            value=False,
            key="payback-manual-saving-toggle",
            help=(
                "Use this only when an approved annual fuel-saving estimate should replace "
                "the annualised FOC-saving result."
            ),
        )

    if use_manual_saving:
        annual_fuel_saving_mt = st.number_input(
            "Approved annual FOC saving [t VLSFO-eq./year]",
            min_value=0.0,
            value=float(calculated_annual_fuel_saving_mt),
            step=1.0,
            key="payback-manual-annual-saving",
        )
    else:
        annual_fuel_saving_mt = calculated_annual_fuel_saving_mt

    include_co2_scenarios = st.checkbox(
        "Include avoided CO2 levy benefit scenarios",
        value=False,
        key="payback-include-co2",
    )
    if include_co2_scenarios:
        annual_avoided_co2_cost_usd = st.number_input(
            "Avoided CO2 levy benefit [US$/year]",
            min_value=0.0,
            value=54_197.0,
            step=1_000.0,
            key="payback-avoided-co2",
            help=(
                "Enter only the portion of the levy avoided because the project reduces emissions, "
                "not the company's total CO2 levy. The example workbook uses $54,197 because "
                "$122,909 - $68,712 = $54,197."
            ),
        )
    else:
        annual_avoided_co2_cost_usd = 0.0

    annual_gross_saving_usd = annual_fuel_saving_mt * fuel_price
    annual_baseline_net_saving_usd = annual_gross_saving_usd - annual_additional_opex_usd

    if capex_usd <= 0:
        st.error("Project CAPEX must be greater than zero.")
        return
    if annual_baseline_net_saving_usd <= 0:
        st.error(
            "Annual net saving is zero or negative. The baseline project cannot achieve payback "
            "with the current assumptions."
        )

    cashflow, scenario_summary = build_payback_analysis(
        capex_usd=capex_usd,
        annual_gross_fuel_saving_usd=annual_gross_saving_usd,
        annual_additional_opex_usd=annual_additional_opex_usd,
        annual_avoided_co2_cost_usd=annual_avoided_co2_cost_usd,
        charter_duration_years=int(charter_duration_years),
    )

    baseline = scenario_summary.iloc[0]
    baseline_payback = baseline["Payback period [years]"]
    result_columns = st.columns(4)
    result_columns[0].metric("CAPEX", f"US$ {capex_usd:,.0f}")
    result_columns[1].metric(
        "Annual fuel-cost saving before OPEX", f"US$ {annual_gross_saving_usd:,.0f}"
    )
    result_columns[2].metric(
        "Baseline payback",
        f"{baseline_payback:.2f} years" if pd.notna(baseline_payback) else "No payback",
    )
    result_columns[3].metric(
        "Net surplus at charter end",
        f"US$ {baseline['Net surplus at charter end [US$]']:,.0f}",
        help=f"Net saving remaining after CAPEX recovery by the end of the {int(charter_duration_years)}-year charter.",
    )

    if baseline["Payback within charter"] == "No":
        st.warning(
            f"The baseline project does not recover its CAPEX within the "
            f"{int(charter_duration_years)}-year charter. Unrecovered CAPEX at charter end: "
            f"US$ {baseline['Unrecovered CAPEX at charter end [US$]']:,.0f}."
        )

    st.markdown("**Payback and charter-end outcome by scenario**")
    st.dataframe(
        scenario_summary.style.format(
            {
                "First-year net benefit [US$]": "US$ {:,.0f}",
                "Payback period [years]": "{:.2f}",
                "Net surplus at charter end [US$]": "US$ {:,.0f}",
                "Unrecovered CAPEX at charter end [US$]": "US$ {:,.0f}",
            },
            na_rep="No payback",
        ),
        hide_index=True,
        use_container_width=True,
    )

    st.markdown("**Cumulative cash flow over the charter**")
    figure = px.line(
        cashflow,
        x="Year",
        y="Cumulative cash flow [US$]",
        color="Scenario",
        markers=True,
        title="Cumulative Cash Flow Across the Charter and Break-even",
    )
    figure.add_hline(
        y=0,
        line_dash="dash",
        line_color="#ef4444",
        annotation_text="Break-even",
        annotation_position="top left",
    )
    figure.update_layout(
        xaxis_title="Charter year",
        yaxis_title="Cumulative cash flow [US$]",
        hovermode="x unified",
        height=480,
        margin={"l": 20, "r": 20, "t": 60, "b": 20},
    )
    st.plotly_chart(
        figure,
        use_container_width=True,
        key="payback-cumulative-cashflow-chart",
    )

    with st.expander("Show yearly cash-flow calculation"):
        st.dataframe(
            cashflow.style.format(
                {
                    "Gross fuel saving [US$]": "{:,.2f}",
                    "Avoided CO2 levy benefit [US$]": "{:,.2f}",
                    "Additional OPEX [US$]": "{:,.2f}",
                    "Net cash flow [US$]": "{:,.2f}",
                    "Cumulative cash flow [US$]": "{:,.2f}",
                }
            ),
            hide_index=True,
            use_container_width=True,
        )
        st.download_button(
            "Download payback cash flow CSV",
            dataframe_csv(cashflow),
            "payback_cash_flow.csv",
            "text/csv",
            key="payback-cashflow-download",
        )

    with st.expander("Show formulas and assumptions"):
        st.code(
            "Period fuel saving = VLSFO-equivalent consumption x FOC saving assumption %\n"
            "Annual fuel saving = period fuel saving x 8,760 / analysis hours\n"
            "Annual gross saving = annual fuel saving x VLSFO reference price\n"
            "Annual net saving = gross saving + avoided CO2 levy benefit - additional OPEX\n"
            "Payback = time until cumulative cash flow reaches US$0\n"
            "Net surplus at charter end = max(total charter net savings - CAPEX, 0)\n"
            "Unrecovered CAPEX = max(CAPEX - total charter net savings, 0)"
        )
        st.caption(
            "The result inherits the app's VLSFO-equivalent fuel conversion and FOC saving assumption. "
            "It is an estimate, not a measured retrofit saving."
        )

    return {
        "analysis_days": analysis_days,
        "period_fuel_saving_mt": period_fuel_saving_mt,
        "annual_fuel_saving_mt": annual_fuel_saving_mt,
        "annual_saving_overridden": bool(use_manual_saving),
        "capex_usd": capex_usd,
        "annual_gross_saving_usd": annual_gross_saving_usd,
        "annual_additional_opex_usd": annual_additional_opex_usd,
        "payback_years": (
            float(baseline_payback) if pd.notna(baseline_payback) else None
        ),
        "charter_duration_years": int(charter_duration_years),
        "payback_within_charter": baseline["Payback within charter"],
        "net_surplus_usd": float(baseline["Net surplus at charter end [US$]"]),
        "unrecovered_capex_usd": float(
            baseline["Unrecovered CAPEX at charter end [US$]"]
        ),
    }


st.title("Vessel Operating Profile and Payback Calculation")
st.caption(
    "Upload Noon, Departure and Arrival reports. Files are identified from their two-row column "
    "titles - not fixed Excel column positions. The app builds operating-hour profiles, a monthly "
    "operating summary, an assumed M/E FOC saving and a charter-period payback calculation."
)

with st.expander("How file validation works"):
    st.markdown(
        """
        - Each upload slot checks the report's section and column titles before reading data.
        - A Noon report placed in the Departure slot is rejected as the wrong report type.
        - Reordered columns are accepted. Deleted required columns are named explicitly and processing stops.
        - You may upload three separate reports, or upload the same combined workbook in all three slots.
        """
    )

upload_columns = st.columns(3)
with upload_columns[0]:
    noon_file = st.file_uploader("1. Noon report", type=["xlsx", "xlsm"], key="noon")
    noon = report_card(noon_file, "noon")
with upload_columns[1]:
    departure_file = st.file_uploader("2. Departure report", type=["xlsx", "xlsm"], key="departure")
    departure = report_card(departure_file, "departure")
with upload_columns[2]:
    arrival_file = st.file_uploader("3. Arrival report", type=["xlsx", "xlsm"], key="arrival")
    arrival = report_card(arrival_file, "arrival")

reports = [noon, departure, arrival]
if not all(reports):
    st.stop()
if any(report.missing for report in reports if report):
    st.error("Processing stopped because one or more required headers are missing.")
    st.stop()

try:
    detected_vessel = validate_vessel_consistency(
        {"noon": noon, "departure": departure, "arrival": arrival}
    )
except VesselValidationError as exc:
    st.error(str(exc))
    st.stop()
st.success(f"Vessel validation passed: {detected_vessel}")

with st.sidebar:
    st.header("Operating-profile methodology")
    st.info(
        "Excel-compatible fixed-bin settings\n\n"
        "Draught: 7-16 m\n\n"
        "Speed: 9-24 kn\n\n"
        "M/E power: 0-22,000 kW\n\n"
        "Arrival duration: excluded"
    )
    known_imo = {"NYK FUTAGO": "9487524"}
    imo_number = st.text_input(
        "IMO number",
        value=known_imo.get(detected_vessel, ""),
        help="The three downloaded report formats do not contain an IMO-number field, so confirm this once per run.",
    )
    foc_saving_percent = st.number_input(
        "Assumed FOC saving [%]", 0.0, 100.0, 1.0, 0.1
    )
    fuel_price = st.number_input(
        "VLSFO reference price [US$/t]", 0.0, 10_000.0, 539.0, 1.0
    )

data_sum = build_excel_data_sum(noon, departure, arrival)
segments = profile_segments_from_data_sum(data_sum)
speed_profile = make_excel_profile(
    segments, "speed_knots", EXCEL_SPEED_EDGES, "speed_included"
)
power_profile = make_excel_profile(
    segments, "me_output_kw", EXCEL_POWER_EDGES, "power_included"
)
monthly = monthly_summary_excel(data_sum)
if not monthly.empty and "propelling_share_valid" in monthly:
    invalid_months = monthly.loc[~monthly["propelling_share_valid"], "month"]
    if not invalid_months.empty:
        months = ", ".join(invalid_months.dt.strftime("%b %Y"))
        st.warning(
            f"Monthly propelling share is unavailable for {months}. Review overlapping "
            "noon-report intervals or missing time coverage; the app does not cap the value at 100%."
        )
overall = excel_overall_summary(data_sum)
fuel = fuel_consumption_summary(noon, arrival)
temperature_audit = sea_temperature_audit(data_sum, noon)
if not temperature_audit["valid"]:
    st.error(temperature_audit["message"])
    st.stop()

st.subheader("Operating Data Summary")
metrics = st.columns(5)
metrics[0].metric("Noon records loaded", f"{len(segments):,}")
metrics[1].metric("M/E propelling days", f"{overall['propelling_hours'] / 24:,.1f}")
metrics[2].metric("Eligible speed-profile hours", f"{speed_profile.total_hours:,.1f}")
metrics[3].metric("Eligible M/E-power-profile hours", f"{power_profile.total_hours:,.1f}")
metrics[4].metric(
    "Share within displayed M/E bands",
    f"{power_profile.percent.to_numpy().sum():.2f}%",
)

tabs = st.tabs(
    [
        "Operating Profile and FOC Saving",
        "Payback Calculation and Charter Outcome",
        "A4 Operating Profile Report",
        "Internal Data_sum",
    ]
)

with tabs[0]:
    st.subheader("Speedâ€“Draught Profile")
    st.caption(
        "Locked Excel method: draught rows start at 7-16 m and speed columns start at 9-24 kn. "
        "A label such as 9 means the 9-<10 kn band. Each cell uses the Excel SUMIFS denominator logic."
    )
    heatmap(
        speed_profile,
        "Speed Profile: Share of Propelling Hours by STW and Draught",
        "STW band start [kn]",
        "speed-draft-heatmap",
    )
    speed_table = render_readable_profile_table(speed_profile)
    st.download_button(
        "Download speed profile CSV",
        dataframe_csv(speed_table, include_index=True),
        "speed_draft_profile.csv",
        "text/csv",
    )
    with st.expander("Show speed-profile summary"):
        render_excel_profile_details(
            speed_profile,
            "Speedâ€“Draught Profile",
            detected_vessel,
            imo_number,
            overall,
        )

    st.divider()
    st.subheader("M/E Powerâ€“Draught Profile")
    st.caption(
        "Locked Excel method: draught rows start at 7-16 m and M/E power columns start at "
        "0-22,000 kW. A label such as 1000 means the 1,000-<2,000 kW band. "
        "Values above the displayed range remain in the denominator exactly as in Excel."
    )
    heatmap(
        power_profile,
        "M/E Power Profile: Share of Propelling Hours by Power and Draught",
        "M/E power band start [kW]",
        "me-output-draft-heatmap",
    )
    power_table = render_readable_profile_table(power_profile)
    st.download_button(
        "Download M/E output profile CSV",
        dataframe_csv(power_table, include_index=True),
        "me_output_draft_profile.csv",
        "text/csv",
    )
    with st.expander("Show M/E power-profile summary"):
        render_excel_profile_details(
            power_profile,
            "M/E Powerâ€“Draught Profile",
            detected_vessel,
            imo_number,
            overall,
        )

    st.divider()
    st.subheader("Monthly Operating Conditions")
    st.caption(
        "This operating summary is calculated from the internal Data_sum and supplies the three charts. "
        "It runs from the earliest to latest Noon/Departure/Arrival month."
    )
    if monthly.empty:
        st.warning("No valid dated operating periods are available for the monthly summary.")
    else:
        monthly_table = show_excel_monthly_table(monthly)
        plot_excel_monthly_graphs(monthly_table, "monthly-analysis")
        st.download_button(
            "Download monthly operating summary CSV",
            dataframe_csv(monthly_table),
            "monthly_operating_summary.csv",
            "text/csv",
        )
    st.divider()
    render_fuel_summary(
        fuel=fuel,
        foc_saving_percent=foc_saving_percent,
        fuel_price=fuel_price,
    )

with tabs[1]:
    payback_result = render_payback_analysis(
        overall=overall,
        fuel=fuel,
        foc_saving_percent=foc_saving_percent,
        fuel_price=fuel_price,
    )

with tabs[2]:
    st.subheader("A4 Operating Profile Report")
    st.caption(
        "Create a one-page PDF containing the vessel scope, operating profile, fuel basis, "
        "assumed saving, commercial outcome and key limitations."
    )
    report_columns = st.columns(2)
    with report_columns[0]:
        prepared_by = st.text_input(
            "Prepared by (optional)",
            value="",
            max_chars=60,
            key="report-prepared-by",
        )
    with report_columns[1]:
        management_comment = st.text_input(
            "Management comment (optional)",
            value="",
            max_chars=180,
            help="Keep this concise so the report remains on one A4 page.",
            key="report-management-comment",
        )

    try:
        report_pdf = build_a4_profile_report(
            vessel_name=detected_vessel,
            imo_number=imo_number,
            overall=overall,
            noon_records=len(segments),
            speed_profile=speed_profile,
            power_profile=power_profile,
            monthly=monthly,
            fuel=fuel,
            foc_saving_percent=foc_saving_percent,
            fuel_price=fuel_price,
            payback=payback_result,
            prepared_by=prepared_by,
            management_comment=management_comment,
        )
    except Exception as exc:
        st.error(f"The A4 report could not be generated: {exc}")
    else:
        st.info(
            "The report is limited to one A4 page. It identifies the FOC saving as an "
            "assumption and does not present it as measured retrofit performance."
        )
        st.download_button(
            "Download one-page A4 PDF report",
            data=report_pdf,
            file_name=safe_report_filename(detected_vessel),
            mime="application/pdf",
            key="download-a4-profile-report",
            use_container_width=True,
        )

with tabs[3]:
    st.caption(
        "This is the internally created Data_sum calculation-input table. The operating profiles, "
        "monthly summary and charts are calculated from these records."
    )
    displayed_data_sum = excel_data_sum_display(data_sum, imo_number)
    st.dataframe(displayed_data_sum, hide_index=True, use_container_width=True)
    st.download_button(
        "Download internal Data_sum CSV",
        dataframe_csv(displayed_data_sum),
        "internal_data_sum.csv",
        "text/csv",
    )
