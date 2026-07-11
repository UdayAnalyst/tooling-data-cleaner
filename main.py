import io
import re

import altair as alt
import pandas as pd
import streamlit as st
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.series import DataPoint
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

STATUS_COLORS = {"Profit": "#0ca30c", "Loss": "#d03b3b"}
STATUS_COLORS_HEX = {"Profit": "0CA30C", "Loss": "D03B3B"}  # no '#', for openpyxl

PO_REGISTRY_HEADERS = ["Okay PN", "Total PO $"]
GSHEET_SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]


@st.cache_resource
def get_gsheet_client():
    import gspread
    from google.oauth2.service_account import Credentials

    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"], scopes=GSHEET_SCOPES
    )
    return gspread.authorize(creds)


def get_po_registry_worksheet():
    import gspread

    client = get_gsheet_client()
    sheet = client.open_by_key(st.secrets["po_registry_sheet_id"])
    try:
        return sheet.worksheet("PO Registry")
    except gspread.WorksheetNotFound:
        worksheet = sheet.add_worksheet(title="PO Registry", rows=1000, cols=2)
        worksheet.append_row(PO_REGISTRY_HEADERS)
        return worksheet


def load_po_registry() -> dict[str, float]:
    """Returns {Okay PN: Total PO $} remembered from previous days. Returns an
    empty registry (rather than raising) if Google Sheets isn't configured, so
    the app still works without it — see PO_REGISTRY_SETUP.md. Fetches
    UNFORMATTED_VALUE so currency-formatted cells (e.g. '$174,827.00', or
    '$ -' for zero — both just display formatting on an underlying number)
    come back as plain numbers instead of strings that fail float(). Rows
    with a genuinely blank/non-numeric Total PO $ (e.g. a new Okay PN added
    by hand and not yet filled in) are skipped individually rather than
    blanking the whole registry, since this feeds save_po_registry's
    full-sheet overwrite."""
    try:
        worksheet = get_po_registry_worksheet()
        records = worksheet.get_all_records(value_render_option="UNFORMATTED_VALUE")
    except Exception:
        return {}

    registry = {}
    for r in records:
        pn = str(r.get("Okay PN", "")).strip()
        if not pn:
            continue
        try:
            registry[normalize_pn(pn)] = float(r["Total PO $"])
        except (KeyError, TypeError, ValueError):
            continue
    return registry


def is_registry_configured() -> bool:
    try:
        return "gcp_service_account" in st.secrets
    except Exception:
        return False


def save_po_registry(registry: dict[str, float]) -> bool:
    """Overwrites the registry sheet with the given {Okay PN: Total PO $} map.
    Returns False (without raising) if Google Sheets isn't configured."""
    try:
        worksheet = get_po_registry_worksheet()
        worksheet.clear()
        rows = [PO_REGISTRY_HEADERS] + [[k, v] for k, v in registry.items()]
        worksheet.update(rows)
        return True
    except Exception:
        return False


def is_odbc_configured() -> bool:
    try:
        return "odbc_dsn" in st.secrets
    except Exception:
        return False


def get_odbc_connection():
    """A fresh pyodbc connection to Plex's cloud ODBC reporting service, via a
    named Windows DSN (Data Source Administrator) that has the host, port, and
    data source saved in it. The DataDirect OpenAccess SDK driver behind this
    DSN does *not* persist the UID/PWD from the setup dialog, so both are
    passed explicitly here — PWD carries the IAM access token
    (authmethod=iam;accesstoken=...), UID the Plex employee/user ID. Opened
    fresh per query rather than cached, since it's only used a handful of
    times per run and a long-lived cached connection can go stale between
    reruns."""
    import pyodbc

    conn_str = (
        f"DSN={st.secrets['odbc_dsn']};"
        f"UID={st.secrets['odbc_uid']};"
        f"PWD={st.secrets['odbc_pwd']};"
    )
    return pyodbc.connect(conn_str)


# Plex tooling job-cost query, adapted from the version provided for the
# DataDirect OpenAccess SQL engine behind Plex's cloud ODBC service, which
# doesn't support T-SQL scripting (DECLARE, OPTION hints) — only a single
# SELECT. It also has a driver bug where a `?` bound parameter silently
# matches zero rows once more than one JOIN is involved (confirmed against
# the live service: identical literal SQL returns correct rows, the same
# query with the Part No. filter as a bound parameter returns none). So
# fetch_plex_job inlines the filter as an escaped, allowlist-validated
# literal via {part_no_pattern} instead of a query parameter — the Job Group
# / Job Key filters from the original query are dropped entirely since they
# were always left at -1 (no-op) in practice.
PLEX_JOB_QUERY_TEMPLATE = """
SELECT
    AJG.Accounting_Job_Group,
    AJ.Accounting_Job_No,
    P.Part_No + '-' + P.Revision            AS [Okay PN],
    P.Name                                  AS [Tooling Line Item Description],
    J.Job_No                                AS [Tooling Job No.],

    ISNULL(SP.Price,0) *
    ISNULL(SR.Quantity,0)                   AS [Total Revenue],
    ISNULL(AR.Credit,0)                     AS [Invoiced Revenue],
    ISNULL(PO_Cost.Cost,0)                  AS [Vendor POs Cost],
    ISNULL(LC.Cost,0)                       AS [Labor Cost],
    ISNULL(PO_Cost.Cost,0) +
    ISNULL(LC.Cost,0)                       AS [Total Cost],
    (ISNULL(SP.Price,0) *
     ISNULL(SR.Quantity,0)) -
    (ISNULL(PO_Cost.Cost,0) +
     ISNULL(LC.Cost,0))                     AS [Profit or Loss]

FROM Part_v_Part P
JOIN Part_v_Part_Product_Type PT
    ON PT.PCN = P.Plexus_Customer_No
    AND PT.Product_Type_Key = P.Product_Type_Key
LEFT OUTER JOIN Part_v_Job J
    ON P.Plexus_Customer_No = J.PCN
    AND P.Part_Key = J.Part_Key
LEFT OUTER JOIN Accounting_v_Accounting_Job AJ
    ON J.PCN = AJ.PCN
    AND J.Accounting_Job_Key = AJ.Accounting_Job_Key
LEFT OUTER JOIN Accounting_v_Accounting_Job_Group AJG
    ON AJ.PCN = AJG.PCN
    AND AJ.Accounting_Job_Group_Key = AJG.Accounting_Job_Group_Key
LEFT OUTER JOIN Sales_v_PO_Line SPOL
    ON P.Plexus_Customer_No = SPOL.PCN
    AND P.Part_Key = SPOL.Part_Key
LEFT OUTER JOIN Sales_v_Price SP
    ON SPOL.PCN = SP.PCN
    AND SPOL.PO_Line_Key = SP.PO_Line_Key
LEFT OUTER JOIN (
    SELECT SPOL2.PCN, SPOL2.PO_Line_Key,
           SUM(SR2.Quantity) AS Quantity
    FROM Sales_v_PO_Line SPOL2
    JOIN Sales_v_Release SR2
        ON SPOL2.PCN = SR2.PCN
        AND SPOL2.PO_Line_Key = SR2.PO_Line_Key
    GROUP BY SPOL2.PCN, SPOL2.PO_Line_Key
) AS SR
    ON SPOL.PCN = SR.PCN
    AND SPOL.PO_Line_Key = SR.PO_Line_Key
LEFT OUTER JOIN (
    SELECT ARID.Plexus_Customer_No, ARID.Part_Key,
           SUM(ARID.Credit) AS Credit
    FROM Accounting_v_AR_Invoice_Dist ARID
    GROUP BY ARID.Plexus_Customer_No, ARID.Part_Key
) AS AR
    ON P.Plexus_Customer_No = AR.Plexus_Customer_No
    AND P.Part_Key = AR.Part_Key
LEFT OUTER JOIN (
    SELECT POL.Plexus_Customer_No, POL.For_Part_Key,
           SUM(POL.Unit_Price * POR2.Quantity) AS Cost
    FROM Purchasing_v_Line_Item POL
    LEFT OUTER JOIN Purchasing_v_Release POR2
        ON POL.Plexus_Customer_No = POR2.Plexus_Customer_No
        AND POL.Line_Item_Key = POR2.Line_Item_Key
    GROUP BY POL.Plexus_Customer_No, POL.For_Part_Key
) AS PO_Cost
    ON P.Plexus_Customer_No = PO_Cost.Plexus_Customer_No
    AND P.Part_Key = PO_Cost.For_Part_Key
LEFT OUTER JOIN (
    SELECT C.PCN, C.Part_Key,
           SUM(ROUND(C.Extended_Cost,2)) AS Cost
    FROM Common_v_Cost C
    WHERE C.Cost_Sub_Type_Key = 17924
    GROUP BY C.PCN, C.Part_Key
) AS LC
    ON P.Plexus_Customer_No = LC.PCN
    AND P.Part_Key = LC.Part_Key

WHERE PT.Product_Type IN (
    'Tooling','Inspection Device','NRE/Tooling',
    'Packaging','In-Development','Production',
    'Protype','Service'
)
AND P.Part_No LIKE '{part_no_pattern}'

ORDER BY
    AJG.Accounting_Job_Group,
    AJ.Accounting_Job_No,
    P.Part_No + '-' + P.Revision
"""

# Allowlist for Part No. filters — since the driver bug above rules out real
# parameter binding, this is the injection defense for the literal
# substitution into PLEX_JOB_QUERY_TEMPLATE instead. Matches the character
# set actually seen in Plex Part Nos (letters, digits, spaces, -_.).
PART_NO_FILTER_RE = re.compile(r"^[A-Za-z0-9 _.-]+$")


def fetch_plex_job(part_no_filter: str) -> pd.DataFrame:
    """Runs the Plex tooling job-cost query for a single Part No. filter (e.g.
    '924'), returning one DataFrame — the ODBC equivalent of one Drive file."""
    if not PART_NO_FILTER_RE.match(part_no_filter):
        raise ValueError(f"Invalid Part No. filter: {part_no_filter!r}")
    pattern = part_no_filter.replace("'", "''") + "%"
    query = PLEX_JOB_QUERY_TEMPLATE.format(part_no_pattern=pattern)
    conn = get_odbc_connection()
    try:
        return pd.read_sql(query, conn)
    finally:
        conn.close()


def fetch_plex_sheets(part_filters: list[str]) -> dict[str, pd.DataFrame]:
    """Runs the Plex query once per Part No. filter, returning {filter: DataFrame}
    — each entry plugs straight into the same cleaning/PO/final-results
    pipeline that a parsed upload sheet would."""
    return {part_no: fetch_plex_job(part_no) for part_no in part_filters}


PROJECT_REGISTRY_HEADERS = [
    "Customer",
    "Project",
    "Part Number",
    "Contingency/Management Reserve Used",
    "Expected Project End Date",
    "NOTES",
]


def get_project_registry_worksheet():
    import gspread

    client = get_gsheet_client()
    sheet = client.open_by_key(st.secrets["po_registry_sheet_id"])
    try:
        return sheet.worksheet("Project Registry")
    except gspread.WorksheetNotFound:
        worksheet = sheet.add_worksheet(
            title="Project Registry", rows=200, cols=len(PROJECT_REGISTRY_HEADERS)
        )
        worksheet.append_row(PROJECT_REGISTRY_HEADERS)
        return worksheet


def load_project_registry() -> pd.DataFrame:
    """Reads the hand-maintained Customer/Project/Part Number/Contingency/
    Expected End Date/NOTES rows. Project Balance is computed separately, not
    stored here. Returns an empty frame if Google Sheets isn't configured."""
    try:
        worksheet = get_project_registry_worksheet()
        records = worksheet.get_all_records()
        if not records:
            return pd.DataFrame(columns=PROJECT_REGISTRY_HEADERS)
        return pd.DataFrame(records)
    except Exception:
        return pd.DataFrame(columns=PROJECT_REGISTRY_HEADERS)


def parse_part_numbers(value) -> list[str]:
    """'931, 932, 985, 988' -> ['931', '932', '985', '988'], de-duplicated."""
    parts = re.split(r"[,\s]+", str(value).strip())
    seen = []
    for part in parts:
        if part and part not in seen:
            seen.append(part)
    return seen


def build_prefix_balances(final_sheets: dict[str, pd.DataFrame]) -> dict[str, float]:
    """Sums each row's Budget Left across every uploaded file, grouped by
    Okay PN prefix (same prefix used for the Step 4 'Project' label)."""
    combined = pd.concat(
        [df[df[df.columns[0]] != "TOTAL"] for df in final_sheets.values()],
        ignore_index=True,
    )
    prefixes = combined["Okay PN"].dropna().astype(str).map(extract_pn_prefix)
    return combined.assign(_prefix=prefixes).groupby("_prefix")["Budget Left"].sum().to_dict()


def build_open_projects(final_sheets: dict[str, pd.DataFrame]) -> pd.DataFrame:
    """Joins the hand-maintained Project Registry with a live-computed
    Project Balance (sum of Budget Left for that row's Part Number(s))."""
    registry_df = load_project_registry()
    if registry_df.empty:
        return registry_df.assign(**{"Project Balance": pd.Series(dtype=float)})

    prefix_balances = build_prefix_balances(final_sheets)
    registry_df["Project Balance"] = registry_df["Part Number"].apply(
        lambda pn: sum(prefix_balances.get(n, 0.0) for n in parse_part_numbers(pn))
    )
    ordered_cols = [
        "Customer",
        "Project",
        "Part Number",
        "Project Balance",
        "Contingency/Management Reserve Used",
        "Expected Project End Date",
        "NOTES",
    ]
    return registry_df[ordered_cols]


HISTORY_HEADERS = ["Date", "Customer", "Project", "Part Number", "Project Balance"]
HISTORY_KEY_COLS = ["Customer", "Project", "Part Number"]


def get_history_worksheet():
    import gspread

    client = get_gsheet_client()
    sheet = client.open_by_key(st.secrets["po_registry_sheet_id"])
    try:
        return sheet.worksheet("History")
    except gspread.WorksheetNotFound:
        worksheet = sheet.add_worksheet(title="History", rows=2000, cols=len(HISTORY_HEADERS))
        worksheet.append_row(HISTORY_HEADERS)
        return worksheet


def load_history() -> pd.DataFrame:
    """Returns every logged Project Balance snapshot, one row per Date x
    Customer/Project/Part Number. Fetches UNFORMATTED_VALUE and coerces
    Project Balance to numeric (dropping rows that fail) so a stray
    currency-formatted or blank cell can't break the trend chart or, worse,
    silently empty the whole history via save_history's overwrite — same
    class of bug fixed in load_po_registry. Returns an empty frame if Google
    Sheets isn't configured."""
    try:
        worksheet = get_history_worksheet()
        records = worksheet.get_all_records(value_render_option="UNFORMATTED_VALUE")
    except Exception:
        return pd.DataFrame(columns=HISTORY_HEADERS)

    if not records:
        return pd.DataFrame(columns=HISTORY_HEADERS)
    history = pd.DataFrame(records)
    history["Project Balance"] = pd.to_numeric(history["Project Balance"], errors="coerce")
    return history.dropna(subset=["Project Balance"])


def log_open_projects_snapshot(open_projects: pd.DataFrame) -> bool:
    """Appends today's Project Balance for each Open Projects row to the
    History tab, so trends can be charted over time. Re-running Generate
    Final Results on the same day replaces that day's rows instead of
    duplicating them. Returns False (without raising) if Sheets isn't
    configured or there's nothing to log."""
    if open_projects.empty:
        return False
    try:
        today = pd.Timestamp.now().strftime("%Y-%m-%d")
        history = load_history()
        history = history[history["Date"] != today]

        new_rows = open_projects[HISTORY_KEY_COLS + ["Project Balance"]].copy()
        new_rows.insert(0, "Date", today)

        combined = pd.concat([history, new_rows], ignore_index=True)[HISTORY_HEADERS]
        worksheet = get_history_worksheet()
        worksheet.clear()
        worksheet.update([HISTORY_HEADERS] + combined.values.tolist())
        return True
    except Exception:
        return False


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


def normalize_pn(value) -> str:
    """Normalizes Okay PN for registry matching: trims whitespace, collapses
    internal whitespace, strips a trailing dash (Plex exports sometimes add
    one, e.g. '932 A PD-01-' vs '932 A PD-01'), and ignores case."""
    text = re.sub(r"\s+", " ", str(value).strip().upper())
    return text.rstrip("- ")


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


def extract_pn_prefix(value: str) -> str:
    """Leading run of letters/digits, stopping at the first dash, dot, space,
    underscore, etc. — e.g. '924-1' and '924-01' both become '924'."""
    match = re.match(r"^[A-Za-z0-9]+", value)
    return match.group(0) if match else value


def derive_sheet_label(df: pd.DataFrame, fallback: str) -> str:
    """Human-readable heading for a sheet, e.g. '932/4066', based on its Okay
    PN prefixes — used instead of the raw uploaded filename, which is often an
    auto-generated export name like 'Query_2026_07_10-09-12-16'. Falls back to
    that filename if no prefixes can be derived (e.g. an empty sheet)."""
    data_rows = df[df[df.columns[0]] != "TOTAL"]
    prefixes = data_rows["Okay PN"].dropna().astype(str).map(extract_pn_prefix).unique()
    return "/".join(prefixes) if len(prefixes) else fallback


def build_project_summary(final_sheets: dict[str, pd.DataFrame]) -> pd.DataFrame:
    """One row per uploaded file, reusing each file's existing TOTAL row from
    Step 3. If a file contains multiple unique Tooling Job No. (or Okay PN
    prefix) values, they are joined with '/'."""
    rows = []
    for sheet_name, df in final_sheets.items():
        data_rows = df[df[df.columns[0]] != "TOTAL"]
        totals_row = df.iloc[-1]
        job_numbers = data_rows["Tooling Job No."].dropna().astype(str).unique()
        pn_prefixes = data_rows["Okay PN"].dropna().astype(str).map(extract_pn_prefix).unique()

        rows.append(
            {
                "File": sheet_name,
                "Project": "/".join(pn_prefixes),
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
        chart.x_axis.title = "Project"
        chart.legend = None
        chart.height, chart.width = 10, 20

        profit_col = project_summary.columns.get_loc("Profit or Loss ($)") + 1
        project_col = project_summary.columns.get_loc("Project") + 1
        data = Reference(worksheet, min_col=profit_col, min_row=1, max_row=n_rows + 1)
        cats = Reference(worksheet, min_col=project_col, min_row=2, max_row=n_rows + 1)
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

raw_sheets = None  # only (re)computed inside the change-gate below

if is_odbc_configured():
    st.caption(
        "Data is queried directly from Plex. Enter one Part No. filter per row "
        "below (e.g. 924) — add rows with the + at the bottom, or paste a whole "
        "column of them straight from Excel — then click 'Fetch from Plex'."
    )

    part_filters_df = st.data_editor(
        pd.DataFrame({"Part No. Filter": [""]}),
        key="part_filters_editor",
        num_rows="dynamic",
        hide_index=True,
        use_container_width=True,
    )
    fetch_clicked = st.button("Fetch from Plex")

    if fetch_clicked:
        part_filters = [
            str(v).strip() for v in part_filters_df["Part No. Filter"] if str(v).strip()
        ]
        if part_filters:
            with st.spinner(f"Querying Plex for {len(part_filters)} Part No. filter(s)..."):
                try:
                    st.session_state.plex_sheets = fetch_plex_sheets(part_filters)
                except Exception as e:
                    st.error(f"Couldn't query Plex: {e}")

    raw_sheets = st.session_state.get("plex_sheets", {})
    if not raw_sheets:
        st.info("Enter one or more Part No. filters above and click 'Fetch from Plex' to get started.")
        st.stop()
    sheet_key = list(raw_sheets.keys())
else:
    uploaded_files = st.file_uploader(
        "Upload CSV or Excel file(s)", type=["csv", "xlsx", "xls"], accept_multiple_files=True
    )
    if not uploaded_files:
        st.info("Upload a file to get started, or see PO_REGISTRY_SETUP.md to connect ODBC instead.")
        st.stop()
    sheet_key = [f.name for f in uploaded_files]

if st.session_state.get("loaded_sheet_key") != sheet_key:
    if raw_sheets is None:
        raw_sheets = load_all_sheets(uploaded_files)
    processed, missing_report = {}, {}

    for sheet_name, df in raw_sheets.items():
        missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
        if missing:
            missing_report[sheet_name] = missing
            continue
        processed[sheet_name] = consolidate_duplicates(df)

    po_registry = load_po_registry()
    st.session_state.registry_connected = is_registry_configured()

    st.session_state.loaded_sheet_key = sheet_key
    st.session_state.missing_report = missing_report
    st.session_state.editor_data = {
        name: df.assign(
            **{"Total PO $": df["Okay PN"].map(normalize_pn).map(po_registry).fillna(0.0)}
        )
        for name, df in processed.items()
    }
    st.session_state.pop("final_sheets", None)

for sheet_name, missing in st.session_state.missing_report.items():
    st.error(f"Sheet '{sheet_name}' is missing required columns {missing} — skipped.")

if not st.session_state.editor_data:
    st.stop()

st.success(f"Loaded {len(st.session_state.editor_data)} sheet(s) after cleaning duplicates.")
st.session_state.setdefault("registry_version", 0)

step2_header, step2_button = st.columns([4, 1])
with step2_header:
    st.header("Step 2: Enter Total PO $ for each row")
if st.session_state.get("registry_connected"):
    with step2_button:
        if st.button("Refresh from registry"):
            fresh_registry = load_po_registry()
            for name, df in st.session_state.editor_data.items():
                st.session_state.editor_data[name] = df.assign(
                    **{"Total PO $": df["Okay PN"].map(normalize_pn).map(fresh_registry).fillna(df["Total PO $"])}
                )
            st.session_state.registry_version += 1
            st.rerun()
    st.caption(
        "Only the 'Total PO $' column is editable. Values already seen for an Okay PN are "
        "pre-filled from your saved registry — just correct the ones that changed. If you edited "
        "the registry in Google Sheets after uploading, click 'Refresh from registry' to pull the "
        "latest values in without needing to re-upload."
    )
else:
    st.caption("Only the 'Total PO $' column is editable — all other columns are shown read-only for context.")
    st.warning(
        "PO $ registry isn't connected yet, so values aren't being remembered day to day. "
        "See PO_REGISTRY_SETUP.md to enable it."
    )

edited_data = {}
for sheet_name, df in st.session_state.editor_data.items():
    label = derive_sheet_label(df, sheet_name)
    with st.expander(label, expanded=len(st.session_state.editor_data) == 1):
        if label != sheet_name:
            st.caption(f"Source file: {sheet_name}")
        edited_data[sheet_name] = st.data_editor(
            df,
            key=f"editor_{sheet_name}_{st.session_state.registry_version}",
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

    if st.session_state.get("registry_connected"):
        registry = load_po_registry()
        for sheet_name, df in edited_data.items():
            baseline = st.session_state.editor_data[sheet_name]
            changed = df[df["Total PO $"] != baseline["Total PO $"]]
            updates = dict(zip(changed["Okay PN"].map(normalize_pn), changed["Total PO $"]))
            registry.update(updates)
        if save_po_registry(registry):
            st.toast(f"Saved {len(registry)} PO $ value(s) to the registry for next time.", icon="✅")
        else:
            st.warning("Couldn't save to the PO $ registry — values won't be remembered next time.")

        open_projects = build_open_projects(final_sheets)
        if log_open_projects_snapshot(open_projects):
            st.toast("Logged today's Project Balance snapshot to History.", icon="📈")

if "final_sheets" in st.session_state:
    tab_data, tab_report = st.tabs(["Data Cleaning", "Report"])

    with tab_data:
        st.header("Step 3: Final Results")
        for sheet_name, df_final in st.session_state.final_sheets.items():
            label = derive_sheet_label(df_final, sheet_name)
            st.subheader(label)
            if label != sheet_name:
                st.caption(f"Source file: {sheet_name}")
            st.dataframe(df_final, use_container_width=True)

        st.download_button(
            label="Download all sheets as Excel",
            data=build_excel(st.session_state.final_sheets),
            file_name="tooling_cleaned_all_sheets.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    with tab_report:
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

        project_order = project_summary["Project"].tolist()
        base_font = "system-ui, -apple-system, Segoe UI, sans-serif"

        bars = (
            alt.Chart(project_summary)
            .mark_bar(size=36, cornerRadiusTopLeft=4, cornerRadiusTopRight=4)
            .encode(
                x=alt.X("Project:N", sort=project_order, title=None, axis=alt.Axis(labelAngle=0, labelColor="#52514e")),
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
                    "Project",
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

        st.header("Step 5: Open Projects")
        open_projects_code = st.text_input("Enter code to view Open Projects", type="password")
        if open_projects_code != st.secrets.get("open_projects_code", "4045"):
            st.info("Enter the access code above to view Open Projects.")
        else:
            st.caption(
                "Customer, Project, Part Number, Contingency/Management Reserve Used, Expected Project End Date "
                "and NOTES are maintained by hand in the 'Project Registry' tab of your Google Sheet. "
                "Project Balance is computed live: sum of Budget Left for that row's Part Number(s)."
            )

            if not st.session_state.get("registry_connected"):
                st.warning("PO $ registry isn't connected, so Open Projects can't be computed. See PO_REGISTRY_SETUP.md.")
            else:
                open_projects = build_open_projects(st.session_state.final_sheets)
                if open_projects.empty:
                    st.info(
                        "No rows yet in the 'Project Registry' tab. Add Customer / Project / Part Number rows "
                        "there and they'll show up here with Project Balance filled in automatically."
                    )
                else:
                    st.dataframe(
                        open_projects,
                        use_container_width=True,
                        hide_index=True,
                        column_config={"Project Balance": st.column_config.NumberColumn(format="$%.2f")},
                    )
