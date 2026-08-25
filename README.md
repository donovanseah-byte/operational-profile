# Vessel Performance Profile Streamlit app

This app reads Noon, Departure and Arrival Excel reports, validates the report type from the workbook's two-row headers, and reproduces the approved Excel profile method:

- a speed-versus-draft operating profile;
- an M/E-output-versus-draft operating profile;
- the Excel-style monthly performance table;
- vessel/IMO, total-duration and profile-duration summaries;
- separate monthly Working Ratio, Sea Water Temperature and Speed charts;
- an auditable internal `Data_sum`, including the graph-source calculation;
- Noon/Arrival consumption for all eight supported M/E fuel grades, normalised to VLSFO equivalent before the PS3 estimate;
- data-quality and exclusion diagnostics.

Both profile matrices provide three presentation views without changing any calculation: colour-coded readable column blocks, the complete matrix, and a non-zero-cell list showing percentage share and propelling hours. Zero cells are displayed as dashes and total rows/columns are visually separated.

The Speed–Draft profile summary displays the maximum Noon-report speed, while the M/E Output–Draft summary displays the maximum Noon-report M/E output. Departure and Arrival values are not used for these maxima.

The parser maps fields by header title and section title. It does not depend on fixed Excel column numbers. Reordered columns are accepted; deleted required columns are reported by name.

The app first creates an internal `Data_sum` containing the combined, chronological Noon, Departure and Arrival rows. All profile summaries, monthly tables and graphs are then calculated from this internal table. The internal `Sea Water Temp.` column is populated from the physical Noon field **Sea Water temperature at noon**, located by its two-row title rather than by Excel column letter. Numeric zeroes are included and blanks are ignored. If the required title is deleted, processing stops instead of silently using another column.

The similarly positioned **Rolling > Avg. Period(sec)** field is not used as temperature; it is a period measured in seconds. The quality check also flags a high proportion of zero readings or values outside -2 to 40 °C.

## Fuel-grade conversion

The fuel table reads the M/E steaming-consumption columns for VLSFO, ULSFO, HSFO, LSDO/LSGO, MDO/MGO, BIO-VLSFO, BIO-HSFO and BIO-LSDO/LSGO from both Noon and Arrival reports. It excludes D/G, boiler, cylinder-oil and stopping-condition consumption, matching the Profile-sheet scope.

Before fuel saving is calculated, every grade is converted to a VLSFO-equivalent mass:

```text
VLSFO-equivalent MT = actual MT × fuel LCV ÷ 40.5
```

The LCV assumptions are VLSFO 40.5, ULSFO 40.5, HSFO 40.2, LSDO/LSGO 42.7, MDO/MGO 42.7, BIO-VLSFO 39.5, BIO-HSFO 39.5 and BIO-LSDO/LSGO 41.5 MJ/kg. The selected PS3 percentage and VLSFO reference price are applied only after this conversion. The app displays the raw grade totals, conversion factors and equivalent totals so the calculation can be audited.

Vessel validation is performed before any calculation. Every uploaded report must contain exactly one non-blank vessel name, and the detected Noon, Departure and Arrival vessel names must match. Mixed-vessel files, missing vessel names and cross-file vessel mismatches stop processing with the detected names shown in the error.

Some reporting-system downloads contain incorrect worksheet metadata declaring that the sheet is only `A1`, even though hundreds of rows and columns exist. The app detects this export defect and automatically reloads the complete worksheet; users do not need to open and resave those reports first.

## Run it

Open a terminal in this folder, then run:

```bash
py -m pip install -r requirements.txt
py -m streamlit run app.py
```

On macOS or Linux, use `python3` instead of `py` if necessary.

Upload separate files in the three labelled slots. If your source is one combined workbook containing all three report sheets, upload the same workbook in all three slots; the app selects the correct sheet for each slot.

## Calculation basis

The method is locked. Departure sets the current draft at the report's **Time of Departure**, Arrival/EOP resets the current draft to zero, and following Noon rows inherit that state. Only Noon propelling duration contributes to the official profile, matching the source workbook.

The monthly graph-source table is created from internal `Data_sum` as the 12 Excel rows beginning with the earliest report month. For each month it calculates the first and last timestamp, elapsed hours, summed propelling hours, working ratio, ordinary average of internal `Sea Water Temp.`, and ordinary average speed. The graph layer performs no further calculation: it plots `Working Ratio[%]`, `Sea Water Temp[°C]`, and `Vs[kts]` directly against the table's Year/Month.

The fixed table bands are:

- draft: 7–<8 m through 16–<17 m;
- speed: 9–<10 kn through 24–<25 kn;
- M/E output: 0–<1,000 kW through 22,000–<23,000 kW.

The displayed row and column headings use only the starting value, like the Excel sheet. For example, speed heading `9` means 9–<10 kn, and M/E output heading `1000` means 1,000–<2,000 kW.

Each heatmap cell is:

```text
100 × propelling hours in the cell / total valid profile propelling hours
```

Rows outside a fixed table range remain in the relevant denominator when the Excel formula does so. Therefore, a table may total less than 100%—for example, NYK FUTAGO's M/E output profile totals 99.61% because a 29,880 kW record is above the fixed table range.

This is a calculation prototype. Confirm business limits, report-header aliases and draft policy against your fleet's controlled reporting specification before production use.

All Plotly charts use explicit unique Streamlit keys, including repeated summary charts in the two profile tabs. This prevents duplicate-element errors in newer Streamlit releases.
