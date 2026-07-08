import io

import altair as alt
import pandas as pd
import streamlit as st
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.series import DataPoint
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

STATUS_COLORS = {"Profit": "#0ca30c", "Loss": "#d03b3b"}
STATUS_COLORS_HEX = {"Profit": "0CA30C", "Loss": "D03B3B"}  # no '#', for openpyxl

st.set_page_config(page_title="Tooling Data Cleaning & Budget Tool", layout="wide")

REQUIRED_COLUMNS = [
    "Tooling Line Item Description",
    "Okay PN",
    "Tooling Job No.",
    "Total Revenue",
    "Invoiced Revenue",
    "Vendor POs Cost",
    "Labor Cost",
    "Total Cost",
    "Profit or Loss",
]

SUM_COLS = ["Invoiced Revenue", "Vendor POs Cost", "Labor Cost", "Total Cost"]

NUMERIC_TOTAL_COLS = [
    "Total Revenue",
    "Invoiced Revenue",
    "Vendor POs Cost",
    "Labor Cost",
    "Total Cost",
    "Profit or Loss",
    "Total PO $",
    "Budget Left",
]


def consolidate_duplicates(df: pd.DataFrame) -> pd.DataFrame:
    """Merge rows that share the same Okay PN, following the same rules as the
    original script: sum Total Revenue only when it differs across the group,
    zero out the other cost/revenue columns on the extra rows (or sum them
    onto the first row when they differ), then drop the extra rows."""
    df_cleaned = df[REQUIRED_COLUMNS].copy()
    df_work = df_cleaned.copy()
    rows_to_drop = []

    for okay_pn in df_work["Okay PN"].unique():
        group_indices = df_work[df_work["Okay PN"] == okay_pn].index.tolist()
        if len(group_indices) == 1:
            continue

        group = df_work.loc[group_indices]
        first_idx, other_indices = group_indices[0], group_indices[1:]

        if len(group["Total Revenue"].unique()) != 1:
            df_work.loc[first_idx, "Total Revenue"] = group["Total Revenue"].sum()

        for col in SUM_COLS:
            if len(group[col].unique()) == 1:
                df_work.loc[other_indices, col] = 0
            else:
                df_work.loc[first_idx, col] = group[col].sum()

        rows_to_drop.extend(other_indices)

    df_cleaned = df_work.drop(rows_to_drop).reset_index(drop=True)
    df_cleaned["Profit or Loss"] = df_cleaned["Total Revenue"] - df_cleaned["Total Cost"]
    return df_cleaned


def load_all_sheets(uploaded_files) -> dict[str, pd.DataFrame]:
    sheets = {}
    for uploaded_file in uploaded_files:
        name = uploaded_file.name
        if name.lower().endswith(".csv"):
            sheets[name.rsplit(".", 1)[0]] = pd.read_csv(uploaded_file)
        elif name.lower().endswith((".xlsx", ".xls")):
            excel_file = pd.ExcelFile(uploaded_file)
            for sheet in excel_file.sheet_names:
                sheets[f"{name}_{sheet}"] = pd.read_excel(excel_file, sheet_name=sheet)
    return sheets


def add_totals_row(df: pd.DataFrame) -> pd.DataFrame:
    totals = {col: (df[col].sum() if col in NUMERIC_TOTAL_COLS else "") for col in df.columns}
    totals[df.columns[0]] = "TOTAL"
    return pd.concat([df, pd.DataFrame([totals])], ignore_index=True)


def build_project_summary(final_sheets: dict[str, pd.DataFrame]) -> pd.DataFrame:
    """One row per uploaded file, reusing each file's existing TOTAL row from
    Step 3. If a file contains multiple unique Tooling Job No. values, they are
    joined with '/'."""
    rows = []
    for sheet_name, df in final_sheets.items():
        data_rows = df[df[df.columns[0]] != "TOTAL"]
        totals_row = df.iloc[-1]
        job_numbers = data_rows["Tooling Job No."].dropna().astype(str).unique()

        rows.append(
            {
                "File": sheet_name,
                "Tooling Job No.": "/".join(job_numbers),
                "Total PO $": totals_row["Total PO $"],
                "Total Cost": totals_row["Total Cost"],
            }
        )

    summary = pd.DataFrame(rows)
    summary["Profit or Loss ($)"] = summary["Total PO $"] - summary["Total Cost"]
    summary["Status"] = summary["Profit or Loss ($)"].apply(lambda v: "Profit" if v >= 0 else "Loss")
    return summary


def build_excel(final_sheets: dict[str, pd.DataFrame]) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for sheet_name, df_final in final_sheets.items():
            df_final.to_excel(writer, sheet_name=sheet_name[:31], index=False)
    return buffer.getvalue()


def build_project_summary_excel(project_summary: pd.DataFrame) -> bytes:
    """Project summary table plus a native, editable Excel bar chart colored
    green/red by Profit/Loss status."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        project_summary.to_excel(writer, sheet_name="Project Summary", index=False)
        worksheet = writer.sheets["Project Summary"]
        n_rows = len(project_summary)

        header_fill = PatternFill("solid", fgColor="2A78D6")
        for cell in worksheet[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center")

        currency_cols = ["Total PO $", "Total Cost", "Profit or Loss ($)"]
        for col_name in currency_cols:
            col_letter = get_column_letter(project_summary.columns.get_loc(col_name) + 1)
            for row in range(2, n_rows + 2):
                worksheet[f"{col_letter}{row}"].number_format = '"$"#,##0.00'

        for i, col_name in enumerate(project_summary.columns, start=1):
            max_len = max(project_summary[col_name].astype(str).map(len).max(), len(col_name)) + 2
            worksheet.column_dimensions[get_column_letter(i)].width = max_len

        chart = BarChart()
        chart.type = "col"
        chart.title = "Profit or Loss by Project"
        chart.y_axis.title = "Profit or Loss ($)"
        chart.x_axis.title = "File"
        chart.legend = None
        chart.height, chart.width = 10, 20

        profit_col = project_summary.columns.get_loc("Profit or Loss ($)") + 1
        file_col = project_summary.columns.get_loc("File") + 1
        data = Reference(worksheet, min_col=profit_col, min_row=1, max_row=n_rows + 1)
        cats = Reference(worksheet, min_col=file_col, min_row=2, max_row=n_rows + 1)
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)

        series = chart.series[0]
        series.data_points = [
            DataPoint(idx=i, spPr=None) for i in range(n_rows)
        ]
        for i, status in enumerate(project_summary["Status"]):
            point = series.data_points[i]
            point.graphicalProperties.solidFill = STATUS_COLORS_HEX[status]
            point.graphicalProperties.line.noFill = True

        worksheet.add_chart(chart, f"{get_column_letter(len(project_summary.columns) + 2)}2")

    return buffer.getvalue()


st.title("Tooling Data Cleaning & Budget Tool")
st.markdown(
    """
1. Upload one or more CSV/Excel files.
2. Enter **Total PO $** for each row directly in the table — every other column stays visible so you always know which row you're entering.
3. Click **Generate Final Results** to calculate Budget Left, add column totals, and download the finished workbook.
"""
)

uploaded_files = st.file_uploader(
    "Upload CSV or Excel file(s)", type=["csv", "xlsx", "xls"], accept_multiple_files=True
)

if not uploaded_files:
    st.info("Upload a file to get started.")
    st.stop()

uploaded_names = [f.name for f in uploaded_files]
if st.session_state.get("uploaded_names") != uploaded_names:
    raw_sheets = load_all_sheets(uploaded_files)
    processed, missing_report = {}, {}

    for sheet_name, df in raw_sheets.items():
        missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
        if missing:
            missing_report[sheet_name] = missing
            continue
        processed[sheet_name] = consolidate_duplicates(df)

    st.session_state.uploaded_names = uploaded_names
    st.session_state.missing_report = missing_report
    st.session_state.editor_data = {
        name: df.assign(**{"Total PO $": 0.0}) for name, df in processed.items()
    }
    st.session_state.pop("final_sheets", None)

for sheet_name, missing in st.session_state.missing_report.items():
    st.error(f"Sheet '{sheet_name}' is missing required columns {missing} — skipped.")

if not st.session_state.editor_data:
    st.stop()

st.success(f"Loaded {len(st.session_state.editor_data)} sheet(s) after cleaning duplicates.")
st.header("Step 2: Enter Total PO $ for each row")
st.caption("Only the 'Total PO $' column is editable — all other columns are shown read-only for context.")

edited_data = {}
for sheet_name, df in st.session_state.editor_data.items():
    st.subheader(sheet_name)
    edited_data[sheet_name] = st.data_editor(
        df,
        key=f"editor_{sheet_name}",
        disabled=[c for c in df.columns if c != "Total PO $"],
        use_container_width=True,
        num_rows="fixed",
    )

if st.button("Generate Final Results", type="primary"):
    final_sheets = {}
    for sheet_name, df in edited_data.items():
        df_final = df.copy()
        df_final["Budget Left"] = df_final["Total PO $"] - df_final["Total Cost"]
        final_sheets[sheet_name] = add_totals_row(df_final)
    st.session_state.final_sheets = final_sheets

if "final_sheets" in st.session_state:
    st.header("Step 3: Final Results")
    for sheet_name, df_final in st.session_state.final_sheets.items():
        st.subheader(sheet_name)
        st.dataframe(df_final, use_container_width=True)

    st.download_button(
        label="Download all sheets as Excel",
        data=build_excel(st.session_state.final_sheets),
        file_name="tooling_cleaned_all_sheets.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.header("Step 4: Profit or Loss by Project")
    st.caption("One row per uploaded file, using its TOTAL row from Step 3. Profit or Loss = Total PO $ - Total Cost.")

    project_summary = build_project_summary(st.session_state.final_sheets).sort_values(
        "Profit or Loss ($)", ascending=False, ignore_index=True
    )

    total_open_projects = len(project_summary)
    projects_on_budget = int((project_summary["Status"] == "Profit").sum())
    pct_on_budget = (projects_on_budget / total_open_projects * 100) if total_open_projects else 0

    kpi1, kpi2, kpi3 = st.columns(3)
    kpi1.metric("Current % on Budget", f"{pct_on_budget:.0f}%")
    kpi2.metric("Projects On Budget", projects_on_budget)
    kpi3.metric("Total Open Projects", total_open_projects)

    st.dataframe(
        project_summary,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Total PO $": st.column_config.NumberColumn(format="$%.2f"),
            "Total Cost": st.column_config.NumberColumn(format="$%.2f"),
            "Profit or Loss ($)": st.column_config.NumberColumn(format="$%.2f"),
        },
    )

    st.download_button(
        label="Download project summary as Excel",
        data=build_project_summary_excel(project_summary),
        file_name="project_profit_or_loss_summary.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    file_order = project_summary["File"].tolist()
    base_font = "system-ui, -apple-system, Segoe UI, sans-serif"

    bars = (
        alt.Chart(project_summary)
        .mark_bar(size=36, cornerRadiusTopLeft=4, cornerRadiusTopRight=4)
        .encode(
            x=alt.X("File:N", sort=file_order, title=None, axis=alt.Axis(labelAngle=0, labelColor="#52514e")),
            y=alt.Y(
                "Profit or Loss ($):Q",
                title="Profit or Loss ($)",
                axis=alt.Axis(format="$,.0f", labelColor="#898781", gridColor="#e1e0d9", titleColor="#52514e"),
            ),
            color=alt.Color(
                "Status:N",
                scale=alt.Scale(domain=list(STATUS_COLORS.keys()), range=list(STATUS_COLORS.values())),
                legend=alt.Legend(title=None, orient="top", symbolType="circle"),
            ),
            tooltip=[
                "File",
                "Tooling Job No.",
                alt.Tooltip("Total PO $:Q", format="$,.2f"),
                alt.Tooltip("Total Cost:Q", format="$,.2f"),
                alt.Tooltip("Profit or Loss ($):Q", format="$,.2f"),
                "Status",
            ],
        )
    )
    labels = bars.mark_text(
        dy=alt.expr("datum['Profit or Loss ($)'] >= 0 ? -8 : 14"),
        color="#0b0b0b",
        fontSize=12,
        font=base_font,
    ).encode(text=alt.Text("Profit or Loss ($):Q", format="$,.0f"))
    zero_line = alt.Chart(pd.DataFrame({"y": [0]})).mark_rule(color="#c3c2b7", strokeWidth=1).encode(y="y:Q")

    chart = (
        (bars + zero_line + labels)
        .properties(
            height=380,
            title=alt.TitleParams(
                "Profit or Loss by Project",
                subtitle="Total PO $ minus Total Cost, per uploaded file",
                fontSize=16,
                subtitleFontSize=12,
                subtitleColor="#898781",
                anchor="start",
                font=base_font,
                subtitleFont=base_font,
            ),
        )
        .configure_view(strokeWidth=0)
        .configure_axis(labelFont=base_font, titleFont=base_font, grid=True, domain=False, tickSize=0)
        .configure_legend(labelFont=base_font, labelFontSize=12)
    )

    st.altair_chart(chart, use_container_width=True)
