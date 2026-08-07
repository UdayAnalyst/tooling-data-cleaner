import io

import altair as alt
import pandas as pd
import streamlit as st

from plex_data import (
    add_totals_row,
    build_daily_report_excel,
    build_excel,
    build_project_summary,
    build_project_summary_excel,
    clean_sheets,
    derive_sheet_label,
    extract_pn_prefix,
    fetch_local_files,
    fetch_plex_groups,
    get_credentials,
    get_or_create_worksheet,
    is_local_export_configured,
    is_odbc_configured,
    is_po_registry_configured,
    is_registry_configured,
    load_all_sheets,
    load_part_groups,
    load_po_registry,
    load_project_registry,
    normalize_pn,
    parse_part_numbers,
    STATUS_COLORS,
)

GDRIVE_SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]


@st.cache_resource
def get_drive_session():
    """An authenticated HTTP session for the Google Drive API, reusing the same
    service account as the PO/Project registries — just needs Drive scope added."""
    from google.auth.transport.requests import AuthorizedSession

    return AuthorizedSession(get_credentials(GDRIVE_SCOPES))


def is_drive_configured() -> bool:
    """True when a Drive folder is configured to read the daily Plex export
    from (daily_plex_export.py writes there) — the fallback source used when
    ODBC isn't available locally, e.g. on a Cloud deployment."""
    try:
        return "gcp_service_account" in st.secrets and "plex_export_folder_id" in st.secrets
    except Exception:
        return False


def fetch_drive_files() -> list[io.BytesIO]:
    """Downloads every CSV/Excel file in the configured Google Drive folder —
    populated daily by daily_plex_export.py — returning file-like objects with
    a `.name` attribute, a drop-in replacement for Streamlit's uploaded-file
    objects, compatible with load_all_sheets()."""
    session = get_drive_session()
    folder_id = st.secrets["plex_export_folder_id"]
    response = session.get(
        "https://www.googleapis.com/drive/v3/files",
        params={
            "q": f"'{folder_id}' in parents and trashed = false",
            "fields": "files(id, name)",
        },
    )
    response.raise_for_status()
    files = response.json().get("files", [])

    buffers = []
    for item in files:
        if not item["name"].lower().endswith((".csv", ".xlsx", ".xls")):
            continue
        response = session.get(
            f"https://www.googleapis.com/drive/v3/files/{item['id']}", params={"alt": "media"}
        )
        response.raise_for_status()
        buffer = io.BytesIO(response.content)
        buffer.name = item["name"]
        buffers.append(buffer)
    return buffers


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
    return get_or_create_worksheet("History", HISTORY_HEADERS, rows=2000)


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

st.title("Tooling Data Cleaning & Budget Tool")

if st.text_input("Enter access code to continue", type="password") != st.secrets.get("site_access_code"):
    st.info("Enter the access code above to continue.")
    st.stop()

raw_sheets = None  # only (re)computed inside the change-gate below

if is_local_export_configured():
    header_col, button_col = st.columns([4, 1])
    with header_col:
        st.caption(
            f"Files are read from the local Plex export folder ({st.secrets['plex_export_local_dir']}), "
            "last written by the daily export."
        )
    with button_col:
        if st.button("Refresh from local export"):
            st.session_state.pop("local_files", None)

    if "local_files" not in st.session_state:
        st.session_state.local_files = fetch_local_files()

    uploaded_files = st.session_state.local_files
    if not uploaded_files:
        st.info(f"No CSV/Excel files found in {st.secrets['plex_export_local_dir']}.")
        st.stop()
    sheet_key = [f.name for f in uploaded_files]
elif is_odbc_configured():
    st.caption("Data is queried directly from Plex, using the Part Number list from your Parts.xlsx.")

    if st.button("Fetch from Plex"):
        with st.spinner("Querying Plex..."):
            try:
                st.session_state.plex_sheets = fetch_plex_groups(load_part_groups())
            except Exception as e:
                st.error(f"Couldn't query Plex: {e}")

    raw_sheets = st.session_state.get("plex_sheets", {})
    if not raw_sheets:
        st.info("Click 'Fetch from Plex' to get started.")
        st.stop()
    sheet_key = list(raw_sheets.keys())
elif is_drive_configured():
    header_col, button_col = st.columns([4, 1])
    with header_col:
        st.caption("Files are fetched automatically from today's Plex export folder in Google Drive.")
    with button_col:
        if st.button("Refresh from Drive"):
            st.session_state.pop("drive_files", None)

    if "drive_files" not in st.session_state:
        with st.spinner("Fetching files from Google Drive..."):
            try:
                st.session_state.drive_files = fetch_drive_files()
            except Exception as e:
                st.error(f"Couldn't fetch files from Google Drive: {e}")
                st.session_state.drive_files = []

    uploaded_files = st.session_state.drive_files
    if not uploaded_files:
        st.info("No CSV/Excel files found in the configured Google Drive folder.")
        st.stop()
    sheet_key = [f.name for f in uploaded_files]
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
    processed, missing_report = clean_sheets(raw_sheets)

    po_registry = load_po_registry()
    st.session_state.registry_connected = is_registry_configured()
    st.session_state.po_registry_connected = is_po_registry_configured()

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
    st.header("Step 2: Total PO $ for each row")
if st.session_state.get("po_registry_connected"):
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
        "'Total PO $' is locked here — it's pulled from the local PO Registry file and can't be edited "
        "from this website. Update values in secret sheet.xlsx's 'PO Registry' tab, then click "
        "'Refresh from registry' to pull the latest values in without needing to re-upload."
    )
else:
    st.caption("'Total PO $' is locked and shown read-only — it can only be set via the local PO Registry file.")
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
            disabled=True,
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
