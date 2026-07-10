# Tooling Data Cleaning & Budget Tool

Single-file Streamlit app (`main.py`, no other source files) that ingests tooling
job cost/revenue spreadsheets, cleans them, and tracks project budgets over time.

Run locally: `streamlit run main.py`

## Data flow (mirrors the on-screen steps)

1. **Input** — either auto-fetched from a configured Google Drive folder
   (`fetch_drive_files`, used when `is_drive_configured()`), or manual upload via
   `st.file_uploader`. Files must contain `REQUIRED_COLUMNS`; sheets missing them
   are skipped with an error, not raised.
2. **Clean** — `load_all_sheets` reads each CSV/Excel sheet into a DataFrame, then
   `consolidate_duplicates` merges rows sharing the same `Okay PN` per fixed rules
   (sum `Total Revenue` only when it differs across the group, sum-or-zero the
   other cost columns, recompute `Profit or Loss`).
3. **Step 2 (editable)** — user fills in `Total PO $` per row in `st.data_editor`;
   pre-filled from the PO Registry by normalized `Okay PN` when connected.
4. **Step 3 (Final Results)** — `Budget Left = Total PO $ - Total Cost`, plus a
   TOTAL row (`add_totals_row`); downloadable as Excel (`build_excel`).
5. **Report tab (Steps 4-6)** — `build_project_summary` (one row per uploaded
   file, Profit/Loss vs. `Total PO $`), KPI metrics, an Altair chart, an Excel
   export with a native colored bar chart (`build_project_summary_excel`),
   **Open Projects** (`build_open_projects`) joining the hand-maintained Project
   Registry with a live-computed Project Balance per Part Number prefix, and
   **Step 6: Project Balance Trend** charting that same balance over time from
   the History tab (see below).

## Three Google Sheets-backed registries (all optional, app degrades gracefully)

- **PO Registry** (`Okay PN` -> `Total PO $`) — remembers PO $ entries across
  days so they don't need re-entering. Read/write via `load_po_registry` /
  `save_po_registry`; only saved on "Generate Final Results" for rows whose
  value actually changed (avoids clobbering concurrent edits).
- **Project Registry** — hand-maintained rows (Customer, Project, Part
  Number(s), contingency flag, expected end date, notes) in a "Project
  Registry" worksheet tab; `Project Balance` is *not* stored, it's computed live
  from `Budget Left` grouped by `Okay PN` prefix (`extract_pn_prefix`).
- **History** — one row per (Date, Customer, Project, Part Number) snapshot of
  Project Balance, logged automatically on "Generate Final Results" via
  `log_open_projects_snapshot`. Re-running on the same day replaces that day's
  rows rather than duplicating them. Feeds the Step 6 trend chart via
  `load_history`.

All three live in the same spreadsheet (`st.secrets["po_registry_sheet_id"]`),
as separate worksheet tabs, auto-created on first use if missing.

Setup instructions for the Sheets/Drive service account: `PO_REGISTRY_SETUP.md`.

## Key normalization helpers

- `normalize_pn` — uppercase, collapse whitespace, strip trailing dash. Needed
  because Plex exports sometimes add a trailing `-` inconsistently
  (`'932 A PD-01-'` vs `'932 A PD-01'`); registry lookups always go through this.
- `extract_pn_prefix` — leading alnum run before the first separator, e.g.
  `'924-1'` and `'924-01'` -> `'924'`. Used to group PNs into a "Project" both
  for the Step 4/Report labels and for Project Balance rollups.

## Config (not in git)

`.streamlit/secrets.toml` (gitignored) holds `gcp_service_account` (service
account JSON fields), `po_registry_sheet_id`, and `gdrive_folder_id`. Every
registry/Drive function checks `is_registry_configured()` / `is_drive_configured()`
first and fails soft (empty registry / manual upload fallback) rather than
raising, so the app works with zero config too.

## Conventions observed in this codebase

- No test suite currently exists.
- All logic lives in `main.py` — no helper modules; keep new functions there
  unless the file grows large enough to warrant a split.
- Functions that hit Google APIs catch broad `Exception` and return a safe
  empty/false value rather than raising, so the UI never hard-crashes on a
  misconfigured or unreachable registry — follow this pattern for any new
  Sheets/Drive calls.
- Gotcha (caused real data loss once, see git history around the
  "PO registry data loss" fix): `worksheet.get_all_records()` defaults to
  `FORMATTED_VALUE`, so a currency-formatted cell comes back as a string like
  `'$174,827.00'` or `'$ -'` instead of a number. Always pass
  `value_render_option="UNFORMATTED_VALUE"` when reading a numeric column, and
  parse/validate rows individually (never one try/except around the whole
  load) — any function that does a full-sheet `clear()` + rewrite (like
  `save_po_registry`, `log_open_projects_snapshot`) will otherwise silently
  wipe out everyone else's data the moment one row fails to parse.

## Pending / not yet pushed

- Commit `808c046` on `development` ("Add Project Balance history logging and
  trend chart (Step 6)") is local-only — not yet pushed to
  `origin/development` or merged into `main`. Revisit and push when ready;
  don't assume the deployed app has this feature yet.
