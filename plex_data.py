"""Plex ODBC + Project Registry helpers shared by main.py (the interactive
Streamlit app) and daily_plex_export.py (the scheduled morning export). Split
out from main.py because main.py executes its whole UI top-to-bottom as soon
as it's imported, so it can't be imported directly from a non-Streamlit
script — this module only defines functions/constants and has no top-level
Streamlit UI calls, so it's safe to import from either."""

import re

import pandas as pd
import streamlit as st

GSHEET_SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]


def get_credentials(scopes: list[str]):
    """Service-account credentials for the given OAuth scopes, built from the
    same gcp_service_account secret used by every Sheets/Drive caller."""
    from google.oauth2.service_account import Credentials

    return Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scopes)


@st.cache_resource
def get_gsheet_client():
    import gspread

    return gspread.authorize(get_credentials(GSHEET_SCOPES))


def is_registry_configured() -> bool:
    try:
        return "gcp_service_account" in st.secrets
    except Exception:
        return False


def get_or_create_worksheet(title: str, headers: list[str], rows: int = 200):
    """Returns the named worksheet tab in the shared registry spreadsheet
    (st.secrets["po_registry_sheet_id"]), creating it with a header row if it
    doesn't exist yet."""
    import gspread

    client = get_gsheet_client()
    sheet = client.open_by_key(st.secrets["po_registry_sheet_id"])
    try:
        return sheet.worksheet(title)
    except gspread.WorksheetNotFound:
        worksheet = sheet.add_worksheet(title=title, rows=rows, cols=len(headers))
        worksheet.append_row(headers)
        return worksheet


PROJECT_REGISTRY_HEADERS = [
    "Customer",
    "Project",
    "Part Number",
    "Contingency/Management Reserve Used",
    "Expected Project End Date",
    "NOTES",
]


def get_project_registry_worksheet():
    return get_or_create_worksheet("Project Registry", PROJECT_REGISTRY_HEADERS, rows=200)


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
    reruns. Only works on a machine that has the Plex ODBC driver and this
    DSN installed (Windows-only, set up via ODBC Data Source Administrator) —
    not available on Streamlit Community Cloud or other Linux hosts."""
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
