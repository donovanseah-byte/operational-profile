from __future__ import annotations

import io

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


st.set_page_config(page_title="Vessel Performance Profile", page_icon="🚢", layout="wide")


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
        f"Detected {REPORT_NAMES[report_type]} • sheet: {report.sheet_name} • "
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
                f"{x_title}: %{{x}}<br>Draft: %{{y}} m<br>Share: %{{z:.3f}}%<extra></extra>"
            ),
            text=text,
            texttemplate="%{text}" if values.size <= 300 else None,
        )
    )
    figure.update_layout(
        title=title,
        xaxis_title=x_title,
        yaxis_title="Mean draft (m)",
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
    return "—" if abs(float(value)) < 0.0005 else f"{float(value):.3f}%"


def styled_profile_table(frame: pd.DataFrame):
    """Apply compact, dependency-free colouring while keeping values numeric."""
    table = frame.copy()
    table.index.name = "Draft start [m]"
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


def render_readable_profile_table(profile, x_label: str, blocks: list[tuple[str, list[str]]]):
    """Offer readable blocks first, with full-matrix and non-zero audit views."""
    table = profile_with_totals(profile.percent)
    table.index.name = "Draft start [m]"
    block_tab, full_tab, detail_tab = st.tabs(
        ["Readable blocks", "Full matrix", "Non-zero cells"]
    )

    with block_tab:
        st.caption(
            "Zero cells are shown as —. Darker red means a larger share of total propelling hours. "
            "The Total column is the row total across all bands."
        )
        for title, columns in blocks:
            available = [column for column in columns if column in table.columns]
            if not available:
                continue
            st.markdown(f"**{title}**")
            block = table[available + ["Total"]]
            st.dataframe(
                styled_profile_table(block),
                use_container_width=True,
                height=430,
            )

    with full_tab:
        st.caption("Complete matrix in one view; scroll horizontally for later bands.")
        st.dataframe(
            styled_profile_table(table),
            use_container_width=True,
            height=460,
        )

    with detail_tab:
        detail = pd.DataFrame(
            {
                "Share [%]": profile.percent.stack(),
                "Propelling hours": profile.hours.stack(),
            }
        ).reset_index()
        detail.columns = ["Draft start [m]", x_label, "Share [%]", "Propelling hours"]
        detail = detail[detail["Share [%]"].gt(0)].reset_index(drop=True)
        detail_max = float(detail["Share [%]"].max()) if not detail.empty else 0.001
        st.dataframe(
            detail,
            hide_index=True,
            use_container_width=True,
            height=460,
            column_config={
                "Share [%]": st.column_config.ProgressColumn(
                    "Share [%]",
                    format="%.3f%%",
                    min_value=0.0,
                    max_value=max(detail_max, 0.001),
                ),
                "Propelling hours": st.column_config.NumberColumn(
                    "Propelling hours", format="%.1f h"
                ),
            },
        )
    return table


def excel_monthly_display(monthly: pd.DataFrame) -> pd.DataFrame:
    """Create the Profile-sheet table that supplies all three line graphs."""
    return pd.DataFrame(
        {
            "YEAR": monthly["month"].dt.year,
            "MONTH": monthly["month"].dt.month,
            "Data_S": monthly["data_start"].dt.strftime("%d/%m/%Y"),
            "Date_E": monthly["data_end"].dt.strftime("%d/%m/%Y"),
            "TTL[h]": monthly["available_hours"],
            "Prop. [h]": monthly["propelling_hours"],
            "Working Ratio[%]": monthly["working_ratio_pct"],
            "Sea Water Temp[°C]": monthly["avg_sea_temp_excel"],
            "Vs[kts]": monthly["avg_speed_knots"],
        }
    )


def show_excel_monthly_table(monthly: pd.DataFrame) -> pd.DataFrame:
    table = excel_monthly_display(monthly)
    st.dataframe(
        table.style.format(
            {
                "TTL[h]": "{:.1f}",
                "Prop. [h]": "{:.1f}",
                "Working Ratio[%]": "{:.0f}%",
                "Sea Water Temp[°C]": "{:.6f}",
                "Vs[kts]": "{:.5f}",
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
        ("Working Ratio[%]", "Working Ratio [%]", "Working ratio (%)"),
        ("Sea Water Temp[°C]", "Sea Water Temp. [°C]", "Temperature (°C)"),
        ("Vs[kts]", "Vs [kts]", "Speed (kn)"),
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
            "Vs": data_sum["speed_knots"],
            "Mid Draft": data_sum["draft_m"],
            "Sea Water Temp. [°C]": data_sum["data_sum_sea_temp"],
            "YEAR": data_sum["year"],
            "MONTH": data_sum["month"],
            "DAY": data_sum["day"],
            "HOUR": data_sum["hour"],
            "MINUTE": data_sum["minute"],
            "Time_R": data_sum["timestamp"],
            "Duration_R [day]": data_sum["duration_days"],
            "M/E output": data_sum["me_output_kw"],
            "Source": data_sum["source"],
        }
    )


def render_excel_profile_details(
    profile,
    profile_name: str,
    vessel_name: str,
    imo_number: str,
    overall: dict,
    monthly: pd.DataFrame,
    fuel: dict,
    ps3_percent: float,
    fuel_price: float,
):
    st.divider()
    st.subheader(f"{profile_name} summary")
    vessel_summary = pd.DataFrame(
        [
            {
                "Vessel Name": vessel_name,
                "IMO No.": imo_number or "Not provided",
                "TTL Duration [day]": overall["propelling_hours"] / 24,
                "Profile TTL [day]": profile.total_hours / 24,
            }
        ]
    )
    st.dataframe(
        vessel_summary.style.format(
            {"TTL Duration [day]": "{:.7f}", "Profile TTL [day]": "{:.7f}"}
        ),
        hide_index=True,
        use_container_width=True,
    )
    is_power_profile = "M/E Output" in profile_name
    if is_power_profile:
        st.metric(
            "Maximum Noon M/E output",
            f"{overall['max_noon_me_output_kw']:,.0f} kW",
        )
        st.caption(
            "Maximum power is taken only from the uploaded Noon report's M/E output column."
        )
    else:
        st.metric(
            "Maximum Noon speed",
            f"{overall['max_noon_speed_knots']:,.2f} kn",
        )
        st.caption(
            "Maximum speed is taken only from the uploaded Noon report's average-speed column."
        )

    overall_row = {
        "YEAR": overall["year"],
        "Data_S": overall["data_start"].strftime("%d/%m/%Y"),
        "Date_E": overall["data_end"].strftime("%d/%m/%Y"),
        "TTL [h]": overall["total_hours"],
        "Prop. [h]": overall["propelling_hours"],
        "Working Ratio [%]": overall["working_ratio_pct"],
        "Sea Water Temp [°C]": overall["avg_sea_temp_excel"],
        "Vs [kts]": overall["avg_speed_knots"],
    }
    if is_power_profile:
        overall_row["Max M/E output [kW]"] = overall["max_noon_me_output_kw"]
    else:
        overall_row["Max Vs [kts]"] = overall["max_noon_speed_knots"]
    overall_table = pd.DataFrame([overall_row])
    overall_formats = {
        "TTL [h]": "{:.0f}",
        "Prop. [h]": "{:.1f}",
        "Working Ratio [%]": "{:.2f}%",
        "Sea Water Temp [°C]": "{:.6f}",
        "Vs [kts]": "{:.5f}",
    }
    if is_power_profile:
        overall_formats["Max M/E output [kW]"] = "{:,.0f}"
    else:
        overall_formats["Max Vs [kts]"] = "{:.2f}"
    st.dataframe(
        overall_table.style.format(overall_formats),
        hide_index=True,
        use_container_width=True,
    )
    st.caption(
        "This value is calculated from the internally created Data_sum 'Sea Water Temp.' column. "
        "The source is located from the uploaded two-row title 'Sea Water temperature at noon', "
        "regardless of its Excel column position."
    )

    if not monthly.empty:
        st.markdown("**Monthly graph-source table**")
        graph_table = show_excel_monthly_table(monthly)
        plot_excel_monthly_graphs(graph_table, profile_name)

    st.subheader("Fuel consumption and PS3 estimate")
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
    with st.expander("Show fuel-grade conversion to VLSFO equivalent"):
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
        st.code("VLSFO-equivalent MT = actual MT × fuel LCV ÷ 40.5")

    equivalent_consumption = fuel["total_vlsfo_equivalent_mt"]
    saving_rate = ps3_percent / 100
    ps3_table = pd.DataFrame(
        [
            {
                "Item": "Total M/E fuel",
                "Common basis": "VLSFO equivalent",
                "Raw consumption [MT]": fuel["total_raw_mt"],
                "Equivalent consumption [MT]": equivalent_consumption,
                "PS3 [%]": ps3_percent,
                "FOC saving [VLSFO-eq. MT]": equivalent_consumption * saving_rate,
                "VLSFO price [US$/MT]": fuel_price,
                "Estimated saving [US$]": equivalent_consumption * saving_rate * fuel_price,
            },
        ]
    )
    st.dataframe(
        ps3_table.style.format(
            {
                "Raw consumption [MT]": "{:,.3f}",
                "Equivalent consumption [MT]": "{:,.3f}",
                "PS3 [%]": "{:.1f}%",
                "FOC saving [VLSFO-eq. MT]": "{:,.3f}",
                "VLSFO price [US$/MT]": "{:,.2f}",
                "Estimated saving [US$]": "{:,.2f}",
            }
        ),
        hide_index=True,
        use_container_width=True,
    )
    st.caption(
        "Includes M/E steaming consumption from Noon and Arrival reports. "
        "D/G, boiler, cylinder oil and stopping-condition fuel remain excluded, matching the Profile scope."
    )


st.title("Vessel Performance Profile")
st.caption(
    "Upload Noon, Departure and Arrival reports. Files are identified from their two-row column "
    "titles—not fixed Excel column positions—and the approved Excel profile method is reproduced automatically."
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
    st.header("Locked methodology")
    st.info(
        "Excel Replication Mode\n\n"
        "Draft: 7–16 m\n\n"
        "Speed: 9–24 kn\n\n"
        "M/E output: 0–22,000 kW\n\n"
        "Arrival duration: excluded"
    )
    known_imo = {"NYK FUTAGO": "9487524"}
    imo_number = st.text_input(
        "IMO number",
        value=known_imo.get(detected_vessel, ""),
        help="The three downloaded report formats do not contain an IMO-number field, so confirm this once per run.",
    )
    ps3_percent = st.number_input("PS3 saving assumption (%)", 0.0, 100.0, 1.0, 0.1)
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

st.subheader("Processing summary")
metrics = st.columns(5)
valid_duration = segments.loc[segments["duration_hours"].gt(0), "duration_hours"].sum()
metrics[0].metric("Loaded Noon rows", f"{len(segments):,}")
metrics[1].metric("Noon propelling hours", f"{valid_duration:,.1f}")
metrics[2].metric("Speed denominator hours", f"{speed_profile.total_hours:,.1f}")
metrics[3].metric("M/E denominator hours", f"{power_profile.total_hours:,.1f}")
metrics[4].metric("M/E table total", f"{power_profile.percent.to_numpy().sum():.2f}%")

tabs = st.tabs(
    ["Data checks", "Profile: Speed vs Draft", "Profile: M/E Output vs Draft", "Monthly analysis", "Processed data"]
)

with tabs[0]:
    st.subheader("Quality-control results")
    if temperature_audit["quality_warnings"]:
        st.warning(
            temperature_audit["message"] + " " + " ".join(temperature_audit["quality_warnings"])
        )
    else:
        st.success(temperature_audit["message"])
    with st.expander("Data_sum Sea Water Temp calculation audit", expanded=True):
        temperature_check = pd.DataFrame(
            [
                {
                    "Calculation source": temperature_audit["source_column"],
                    "Data_sum header": temperature_audit["source_header"],
                    "Uploaded source column": temperature_audit["upstream_column"],
                    "Uploaded source header": temperature_audit["upstream_header"],
                    "Numeric readings": temperature_audit["numeric_count"],
                    "Zero readings included": temperature_audit["zero_count"],
                    "Sum": temperature_audit["sum_value"],
                    "Average": temperature_audit["average_value"],
                    "Minimum": temperature_audit["minimum_value"],
                    "Maximum": temperature_audit["maximum_value"],
                    "Outside -2 to 40 °C": temperature_audit["out_of_range_count"],
                }
            ]
        )
        st.dataframe(
            temperature_check.style.format(
                {
                    "Sum": "{:,.1f}",
                    "Average": "{:.6f}",
                    "Minimum": "{:.1f}",
                    "Maximum": "{:.1f}",
                }
            ),
            hide_index=True,
            use_container_width=True,
        )
        st.code(
            f"{temperature_audit['sum_value']:,.1f} ÷ "
            f"{temperature_audit['numeric_count']:,} = "
            f"{temperature_audit['average_value']:.6f}"
        )
    check_columns = st.columns(2)
    with check_columns[0]:
        st.metric("Speed table coverage", f"{speed_profile.coverage:.2%}")
        speed_excluded = segments.loc[~segments["speed_included"], "speed_exclusion"].value_counts()
        st.write("Speed exclusions")
        st.dataframe(
            speed_excluded.rename_axis("Reason").reset_index(name="Rows"),
            hide_index=True,
            use_container_width=True,
        )
    with check_columns[1]:
        st.metric("M/E table coverage", f"{power_profile.coverage:.2%}")
        power_excluded = segments.loc[~segments["power_included"], "power_exclusion"].value_counts()
        st.write("M/E output exclusions")
        st.dataframe(
            power_excluded.rename_axis("Reason").reset_index(name="Rows"),
            hide_index=True,
            use_container_width=True,
        )
    missing_draft = segments[segments["draft_m"].isna()]
    if not missing_draft.empty:
        st.warning(
            f"{len(missing_draft)} Noon rows occur before the first usable Departure/Arrival draft state. "
            "They are excluded in the same way as the Excel workbook."
        )
    outside_power = power_profile.total_hours - power_profile.hours.to_numpy().sum()
    if outside_power > 0.001:
        st.warning(
            f"{outside_power:,.1f} denominator hours fall outside the fixed 0–<23,000 kW table. "
            "They remain in the denominator, matching the Excel formula."
        )
    high_working_ratio = monthly[monthly["working_ratio_pct"] > 100.5] if not monthly.empty else pd.DataFrame()
    if not high_working_ratio.empty:
        st.warning("Some monthly working ratios exceed 100%; inspect overlapping or duplicate report periods.")

with tabs[1]:
    st.caption(
        "Locked Excel method: draft rows start at 7–16 m and speed columns start at 9–24 kn. "
        "A label such as 9 means the 9–<10 kn band. Each cell uses the Excel SUMIFS denominator logic."
    )
    heatmap(
        speed_profile,
        "Speed–Draft Operating Profile",
        "Average speed (kn)",
        "speed-draft-heatmap",
    )
    speed_table = render_readable_profile_table(
        speed_profile,
        "Speed start [kn]",
        [
            ("Speed columns 9–16 kn", [str(value) for value in range(9, 17)]),
            ("Speed columns 17–24 kn", [str(value) for value in range(17, 25)]),
        ],
    )
    st.download_button(
        "Download speed profile CSV",
        dataframe_csv(speed_table, include_index=True),
        "speed_draft_profile.csv",
        "text/csv",
    )
    render_excel_profile_details(
        speed_profile,
        "Speed–Draft Profile",
        detected_vessel,
        imo_number,
        overall,
        monthly,
        fuel,
        ps3_percent,
        fuel_price,
    )

with tabs[2]:
    st.caption(
        "Locked Excel method: draft rows start at 7–16 m and M/E output columns start at "
        "0–22,000 kW. A label such as 1000 means the 1,000–<2,000 kW band. "
        "Values above the displayed range remain in the denominator exactly as in Excel."
    )
    heatmap(
        power_profile,
        "M/E Output–Draft Operating Profile",
        "M/E output (kW)",
        "me-output-draft-heatmap",
    )
    power_table = render_readable_profile_table(
        power_profile,
        "M/E output start [kW]",
        [
            ("M/E output columns 0–7,000 kW", [str(value) for value in range(0, 8_000, 1_000)]),
            ("M/E output columns 8,000–15,000 kW", [str(value) for value in range(8_000, 16_000, 1_000)]),
            ("M/E output columns 16,000–22,000 kW", [str(value) for value in range(16_000, 23_000, 1_000)]),
        ],
    )
    st.download_button(
        "Download M/E output profile CSV",
        dataframe_csv(power_table, include_index=True),
        "me_output_draft_profile.csv",
        "text/csv",
    )
    render_excel_profile_details(
        power_profile,
        "M/E Output–Draft Profile",
        detected_vessel,
        imo_number,
        overall,
        monthly,
        fuel,
        ps3_percent,
        fuel_price,
    )

with tabs[3]:
    st.caption(
        "This table is calculated from the internal Data_sum and supplies the three graphs. "
        "It runs from the earliest to latest Noon/Departure/Arrival month."
    )
    if monthly.empty:
        st.warning("No valid dated operating periods are available for monthly analysis.")
    else:
        monthly_table = show_excel_monthly_table(monthly)
        plot_excel_monthly_graphs(monthly_table, "monthly-analysis")
        st.download_button(
            "Download monthly analysis CSV",
            dataframe_csv(monthly_table),
            "monthly_analysis.csv",
            "text/csv",
        )

with tabs[4]:
    st.caption(
        "This is the internally created Data_sum. The profile summaries, monthly table and "
        "graphs are calculated from this table."
    )
    displayed_data_sum = excel_data_sum_display(data_sum, imo_number)
    st.dataframe(displayed_data_sum, hide_index=True, use_container_width=True)
    st.download_button(
        "Download internal Data_sum CSV",
        dataframe_csv(displayed_data_sum),
        "internal_data_sum.csv",
        "text/csv",
    )
