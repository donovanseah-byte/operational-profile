from __future__ import annotations

import io
import math

import pandas as pd
import plotly.express as px
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


st.set_page_config(page_title="Vessel Operating Profile & Payback Calculation", layout="wide")


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


_rendered_heatmap_keys: set[str] = set()


def dataframe_csv(frame: pd.DataFrame, include_index: bool = False) -> bytes:
    return frame.to_csv(index=include_index).encode("utf-8")


def _profile_percent_text(value) -> str:
    if pd.isna(value):
        return ""
    return "-" if abs(float(value)) < 0.0005 else f"{float(value):.3f}%"


def styled_profile_table(frame: pd.DataFrame):
    """Apply compact, dependency-free colouring while keeping values numeric."""
    table = frame.copy()
    table.index.name = "Draft band start [m]"
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


def render_full_profile_heatmap(profile, key: str):
    """Render one complete operating-profile matrix as the sole heatmap."""
    table = profile_with_totals(profile.percent)
    table.index.name = "Draft band start [m]"
    st.dataframe(
        styled_profile_table(table),
        use_container_width=True,
        height=500,
        key=key,
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
            "TTL[h]": monthly["available_hours"],
            "Reported Propelling Hours [h]": monthly["propelling_hours"],
            "Reported Propelling Ratio [%]": monthly["working_ratio_pct"],
            "Mean Reported Sea-Water Temperature [deg C]": monthly["avg_sea_temp_excel"],
            "Mean Reported Speed [kn]": monthly["avg_speed_knots"],
        }
    )


def show_excel_monthly_table(monthly: pd.DataFrame) -> pd.DataFrame:
    table = excel_monthly_display(monthly)
    st.dataframe(
        table.style.format(
            {
                "TTL[h]": "{:.1f}",
                "Reported Propelling Hours [h]": "{:.1f}",
                "Reported Propelling Ratio [%]": "{:.0f}%",
                "Mean Reported Sea-Water Temperature [deg C]": "{:.6f}",
                "Mean Reported Speed [kn]": "{:.5f}",
            },
            na_rep="#DIV/0!",
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
            "Reported Propelling Ratio [%]",
            "Monthly Reported Propelling Ratio",
            "Reported propelling ratio [%]",
        ),
        (
            "Mean Reported Sea-Water Temperature [deg C]",
            "Monthly Mean Reported Sea-Water Temperature",
            "Mean reported temperature [deg C]",
        ),
        (
            "Mean Reported Speed [kn]",
            "Monthly Mean Reported Speed",
            "Mean reported speed [kn]",
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
            "Reported Speed [kn]": data_sum["speed_knots"],
            "Active Midship Draft [m]": data_sum["draft_m"],
            "Reported Sea-Water Temperature [deg C]": data_sum["data_sum_sea_temp"],
            "YEAR": data_sum["year"],
            "MONTH": data_sum["month"],
            "DAY": data_sum["day"],
            "HOUR": data_sum["hour"],
            "MINUTE": data_sum["minute"],
            "Reported Duration [days]": data_sum["duration_days"],
            "M/E Output [kW]": data_sum["me_output_kw"],
            "Source": data_sum["source"],
        }
    )


def render_combined_operating_profile_summary(
    speed_profile,
    power_profile,
    vessel_name: str,
    imo_number: str,
    overall: dict,
):
    """Show one combined summary for both operating profiles."""
    st.subheader("Operating Profile Summary")

    summary = pd.DataFrame(
        [
            {
                "Vessel Name": vessel_name,
                "IMO No.": imo_number or "Not provided",
                "Period Start": overall["data_start"].strftime("%d/%m/%Y"),
                "Period End": overall["data_end"].strftime("%d/%m/%Y"),
                "Total Propelling Hours [h]": overall["propelling_hours"],
                "Speed–Draft Profile Hours [h]": speed_profile.total_hours,
                "M/E Output–Draft Profile Hours [h]": power_profile.total_hours,
                "Maximum Reported Noon Speed [kn]": overall["max_noon_speed_knots"],
                "Maximum Reported Noon M/E Output [kW]": overall["max_noon_me_output_kw"],
                "Mean Reported Speed [kn]": overall["avg_speed_knots"],
                "Mean Reported Sea-Water Temperature [deg C]": overall["avg_sea_temp_excel"],
                "Reported Propelling Ratio [%]": overall["working_ratio_pct"],
            }
        ]
    )
    st.dataframe(
        summary.style.format(
            {
                "Total Propelling Hours [h]": "{:,.1f}",
                "Speed–Draft Profile Hours [h]": "{:,.1f}",
                "M/E Output–Draft Profile Hours [h]": "{:,.1f}",
                "Maximum Reported Noon Speed [kn]": "{:.2f}",
                "Maximum Reported Noon M/E Output [kW]": "{:,.0f}",
                "Mean Reported Speed [kn]": "{:.5f}",
                "Mean Reported Sea-Water Temperature [deg C]": "{:.6f}",
                "Reported Propelling Ratio [%]": "{:.2f}%",
            }
        ),
        hide_index=True,
        use_container_width=True,
    )



def render_fuel_summary(fuel: dict, ps3_percent: float, fuel_price: float):
    """Render one concise fuel-saving summary with detailed fuel data collapsed below."""
    st.subheader("M/E Fuel Consumption & Estimated FOC Saving")

    equivalent_consumption = float(fuel.get("total_vlsfo_equivalent_mt", 0.0) or 0.0)
    raw_consumption = float(fuel.get("total_raw_mt", 0.0) or 0.0)
    saving_rate = ps3_percent / 100
    period_saving = equivalent_consumption * saving_rate
    period_cost_saving = period_saving * fuel_price

    summary_table = pd.DataFrame(
        [
            {
                "VLSFO-Equivalent M/E Fuel [MT]": equivalent_consumption,
                "Assumed Period FOC Reduction [%]": ps3_percent,
                "Estimated Fuel Saving - Analysis Period [MT]": period_saving,
                "Estimated Bunker Cost Saving - Analysis Period [US$]": period_cost_saving,
            }
        ]
    )
    st.dataframe(
        summary_table.style.format(
            {
                "VLSFO-Equivalent M/E Fuel [MT]": "{:,.3f}",
                "Assumed Period FOC Reduction [%]": "{:.1f}%",
                "Estimated Fuel Saving - Analysis Period [MT]": "{:,.3f}",
                "Estimated Bunker Cost Saving - Analysis Period [US$]": "US$ {:,.2f}",
            }
        ),
        hide_index=True,
        use_container_width=True,
    )

    with st.expander("Show M/E fuel consumption breakdown"):
        grades = list(fuel["total_by_grade"])
        fuel_rows = []
        for report_name, grade_values, equivalent in (
            ("Noon", fuel["noon_by_grade"], fuel["noon_vlsfo_equivalent_mt"]),
            ("Arrival", fuel["arrival_by_grade"], fuel["arrival_vlsfo_equivalent_mt"]),
        ):
            row = {"Report": report_name}
            row.update({f"{grade} [MT]": grade_values[grade] for grade in grades})
            row["Raw total [MT]"] = sum(grade_values.values())
            row["VLSFO-equivalent [MT]"] = equivalent
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
        st.caption(
            f"Reported M/E fuel across Noon and Arrival reports: {raw_consumption:,.3f} MT. "
            "D/G, boiler, cylinder oil and stopping-condition fuel remain excluded."
        )

    with st.expander("Show fuel-grade conversion to VLSFO equivalent"):
        grades = list(fuel["total_by_grade"])
        conversion_table = pd.DataFrame(
            [
                {
                    "Fuel grade": grade,
                    "Actual total [MT]": fuel["total_by_grade"][grade],
                    "LCV [MJ/kg]": fuel["lcv_mj_per_kg"][grade],
                    "Conversion factor": fuel["conversion_factor"][grade],
                    "VLSFO-equivalent [MT]": fuel["equivalent_by_grade"][grade],
                }
                for grade in grades
            ]
        )
        st.dataframe(
            conversion_table.style.format(
                {
                    "Actual total [MT]": "{:,.3f}",
                    "LCV [MJ/kg]": "{:.1f}",
                    "Conversion factor": "{:.6f}",
                    "VLSFO-equivalent [MT]": "{:,.3f}",
                }
            ),
            hide_index=True,
            use_container_width=True,
        )
        st.code("VLSFO-equivalent MT = actual MT x fuel LCV / 40.5")


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
    ps3_percent: float,
    fuel_price: float,
):
    """Render the payback tab from the existing Profile fuel result."""
    st.subheader("Retrofit Payback & Charter Analysis")
    st.caption(
        "The app annualises the assumed FOC reduction calculated over the analysis period. "
        "CAPEX and other commercial assumptions must be entered manually."
    )

    analysis_hours = float(overall.get("total_hours", float("nan")))
    equivalent_consumption_mt = float(fuel.get("total_vlsfo_equivalent_mt", 0.0))

    if not math.isfinite(analysis_hours) or analysis_hours <= 0:
        st.error("Payback cannot be calculated because the analysis duration is invalid.")
        return
    if not 0 <= ps3_percent <= 100:
        st.error("Assumed FOC reduction must be between 0% and 100%.")
        return
    if fuel_price < 0:
        st.error("Fuel price cannot be negative.")
        return

    analysis_days = analysis_hours / 24
    period_fuel_saving_mt = equivalent_consumption_mt * ps3_percent / 100
    annualisation_factor = (365 * 24) / analysis_hours
    calculated_annual_fuel_saving_mt = period_fuel_saving_mt * annualisation_factor

    basis_columns = st.columns(3)
    basis_columns[0].metric("Analysis Period", f"{analysis_days:,.1f} days")
    basis_columns[1].metric("Fuel Saving — Analysis Period", f"{period_fuel_saving_mt:,.3f} MT")
    basis_columns[2].metric(
        "Projected Annual Fuel Saving",
        f"{calculated_annual_fuel_saving_mt:,.3f} MT/year",
    )

    if analysis_days < 180:
        st.warning(
            "The uploaded period is shorter than 180 days. Annualising a short period can produce "
            "an unstable payback estimate, especially if vessel operations are seasonal."
        )

    st.markdown("**Financial assumptions**")
    input_columns = st.columns(3)
    with input_columns[0]:
        capex_amount = st.number_input(
            "Project CAPEX",
            min_value=0.0,
            value=331_800.0,
            step=1_000.0,
            key="payback-capex-amount",
        )
    with input_columns[1]:
        capex_currency = st.selectbox(
            "CAPEX currency",
            ["EUR", "USD"],
            key="payback-capex-currency",
        )
    with input_columns[2]:
        if capex_currency == "EUR":
            exchange_rate = st.number_input(
                "EUR to USD exchange rate",
                min_value=0.0001,
                value=1.18,
                step=0.01,
                format="%.4f",
                key="payback-eur-usd-rate",
            )
        else:
            exchange_rate = 1.0
            st.text_input(
                "Exchange rate",
                value="Not required for USD CAPEX",
                disabled=True,
                key="payback-usd-rate-note",
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
            "Additional Annual OPEX [US$]",
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
                "the projected annual fuel-saving result."
            ),
        )

    if use_manual_saving:
        annual_fuel_saving_mt = st.number_input(
            "Approved annual fuel saving [VLSFO-equivalent MT/year]",
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

    capex_usd = capex_amount * exchange_rate
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

    st.markdown("**Payback & Charter Outcome by Scenario**")
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
            "Fuel saving — analysis period = VLSFO-equivalent consumption x assumed FOC reduction %\n"
            "Annual fuel saving = period fuel saving x 8,760 / analysis hours\n"
            "Annual gross saving = annual fuel saving x VLSFO reference price\n"
            "Annual net saving = gross saving + avoided CO2 levy benefit - additional OPEX\n"
            "Payback = time until cumulative cash flow reaches US$0\n"
            "Net surplus at charter end = max(total charter net savings - CAPEX, 0)\n"
            "Unrecovered CAPEX = max(CAPEX - total charter net savings, 0)"
        )
        st.caption(
            "The result inherits the app's VLSFO-equivalent fuel conversion and assumed FOC reduction. "
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
        "payback_years": float(baseline_payback) if pd.notna(baseline_payback) else None,
        "charter_duration_years": int(charter_duration_years),
        "payback_within_charter": baseline["Payback within charter"],
        "net_surplus_usd": float(baseline["Net surplus at charter end [US$]"]),
        "unrecovered_capex_usd": float(baseline["Unrecovered CAPEX at charter end [US$]"]),
    }


st.title("Vessel Operating Profile & Payback Calculation")
st.caption(
    "Upload Noon, Departure and Arrival reports. Files are identified from their two-row column "
    "titles - not fixed Excel column positions. The app builds operating-hour profiles, a monthly "
    "operating summary, an M/E fuel-saving estimate and a charter-period payback analysis."
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
    st.header("Operating Profile Settings")
    st.info(
        "Excel-compatible fixed-bin settings\n\n"
        "Draft: 7-16 m\n\n"
        "Speed: 9-24 kn\n\n"
        "M/E output: 0-22,000 kW\n\n"
        "Arrival duration: excluded"
    )
    known_imo = {"NYK FUTAGO": "9487524"}
    imo_number = st.text_input(
        "IMO number",
        value=known_imo.get(detected_vessel, ""),
        help="The three downloaded report formats do not contain an IMO-number field, so confirm this once per run.",
    )
    ps3_percent = st.number_input("Assumed FOC Reduction [%]", 0.0, 100.0, 1.0, 0.1)
    fuel_price = st.number_input(
        "VLSFO reference price (US$/MT)", 0.0, 10_000.0, 539.0, 1.0
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
overall = excel_overall_summary(data_sum)
fuel = fuel_consumption_summary(noon, arrival)
temperature_audit = sea_temperature_audit(data_sum, noon)
if not temperature_audit["valid"]:
    st.error(temperature_audit["message"])
    st.stop()

tabs = st.tabs(
    [
        "Operating Profile",
        "Retrofit Payback & Charter Analysis",
        "A4 Professional Report",
        "Processed Data",
    ]
)

with tabs[0]:
    st.caption(
        "The complete Speed–Draft and M/E Output–Draft operating profiles are shown below. "
        "Each profile uses one full matrix heatmap only."
    )

    st.markdown("### Speed–Draft Operating Profile")
    st.caption("Draft: 7–16 m | Speed: 9–24 kn")
    speed_table = render_full_profile_heatmap(
        speed_profile,
        "speed-draft-full-profile-heatmap",
    )
    st.download_button(
        "Download Speed–Draft profile CSV",
        dataframe_csv(speed_table, include_index=True),
        "speed_draft_profile.csv",
        "text/csv",
    )

    st.divider()
    st.markdown("### M/E Output–Draft Operating Profile")
    st.caption("Draft: 7–16 m | M/E Output: 0–22,000 kW")
    power_table = render_full_profile_heatmap(
        power_profile,
        "me-output-draft-full-profile-heatmap",
    )
    st.download_button(
        "Download M/E Output–Draft profile CSV",
        dataframe_csv(power_table, include_index=True),
        "me_output_draft_profile.csv",
        "text/csv",
    )
    st.caption(
        "Heatmap values represent the percentage of total eligible propelling hours "
        "within each operating band."
    )

    st.divider()
    render_combined_operating_profile_summary(
        speed_profile=speed_profile,
        power_profile=power_profile,
        vessel_name=detected_vessel,
        imo_number=imo_number,
        overall=overall,
    )

    st.markdown("### Monthly Operating Trends")
    st.caption("The three graphs are calculated from the same processed monthly operating data.")
    if monthly.empty:
        st.warning("No valid dated operating periods are available for the monthly summary.")
    else:
        monthly_table = excel_monthly_display(monthly)
        plot_excel_monthly_graphs(monthly_table, "monthly-analysis")
        with st.expander("Show monthly operating data table"):
            show_excel_monthly_table(monthly)
            st.download_button(
                "Download monthly operating summary CSV",
                dataframe_csv(monthly_table),
                "monthly_operating_summary.csv",
                "text/csv",
            )

    st.divider()
    render_fuel_summary(
        fuel=fuel,
        ps3_percent=ps3_percent,
        fuel_price=fuel_price,
    )

with tabs[1]:
    payback_result = render_payback_analysis(
        overall=overall,
        fuel=fuel,
        ps3_percent=ps3_percent,
        fuel_price=fuel_price,
    )

with tabs[2]:
    st.subheader("A4 Professional Report")
    st.caption(
        "Generate the concise one-page A4 report created for this app. It uses the three monthly "
        "operating graphs, the operating-profile summary, fuel-saving basis and payback outcome."
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
            foc_saving_percent=ps3_percent,
            fuel_price=fuel_price,
            payback=payback_result,
            prepared_by=prepared_by,
            management_comment=management_comment,
        )
    except Exception as exc:
        st.error(f"The A4 report could not be generated: {exc}")
    else:
        st.info(
            "The report is limited to one A4 page. The FOC reduction is identified as an "
            "assumption and is not presented as measured retrofit performance."
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
        "This is the processed calculation-input table. The operating profiles, "
        "monthly summary and charts are calculated from these records."
    )
    displayed_data_sum = excel_data_sum_display(data_sum, imo_number)
    st.dataframe(displayed_data_sum, hide_index=True, use_container_width=True)
    st.download_button(
        "Download Processed Data CSV",
        dataframe_csv(displayed_data_sum),
        "internal_data_sum.csv",
        "text/csv",
    )
