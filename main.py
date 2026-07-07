import io

import pandas as pd
import streamlit as st

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


def build_excel(final_sheets: dict[str, pd.DataFrame]) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for sheet_name, df_final in final_sheets.items():
            df_final.to_excel(writer, sheet_name=sheet_name[:31], index=False)
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
