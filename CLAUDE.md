# Tooling Data Cleaning & Budget Tool

Streamlit app (`main.py`, plus two companion scripts — see below) that
ingests tooling job cost/revenue data, cleans it, and tracks project budgets
over time.

Run locally: `streamlit run main.py`

## Three files, not one — why

Started as a single-file app; grew a second entry point once Plex ODBC access
turned out to be Windows-machine-local only (see "Deployment topology"
below), which needed a script that runs *without* triggering main.py's UI.

- **`main.py`** — the interactive Streamlit app (`streamlit run main.py`).
- **`plex_data.py`** — shared, UI-free helpers: the Plex ODBC query/connection
  and Project Registry read functions. No top-level Streamlit calls, so it's
  safe to import from a plain script — `main.py` executes its whole UI
  top-to-bottom the moment it's imported, so it *can't* be imported directly;
  anything needed outside the interactive app has to live here instead.
- **`daily_plex_export.py`** — standalone scheduled script, not part of the
  interactive app (see "Daily Plex export" below).

## Data flow (mirrors the on-screen steps)

1. **Input**, in priority order:
   1. **Live Plex ODBC** (`fetch_plex_sheets` from `plex_data`, used when
      `is_odbc_configured()`) — user enters Part No. filters one per row in an
      `st.data_editor` column, one `PLEX_JOB_QUERY_TEMPLATE` run per filter via
      `fetch_plex_job`, each result is one "sheet." Only works on a machine
      with the Plex ODBC driver + DSN installed (this laptop) — see
      "Deployment topology."
   2. **Google Drive auto-fetch** (`fetch_drive_files`, used when
      `is_drive_configured()`) — reads whatever `daily_plex_export.py` last
      wrote to the configured Drive folder. This is the path a Cloud
      deployment actually uses, since it has no ODBC access.
   3. **Manual upload** via `st.file_uploader`, the fallback when neither of
      the above is configured.

   Whichever source is used, sheets must contain `REQUIRED_COLUMNS`; ones
   missing them are skipped with an error, not raised. Typing/entering rows
   one at a time in the ODBC Part No. editor is confirmed reliable; pasting a
   multi-line clipboard column into it landed as one broken cell in testing
   (2026-07-11) rather than spreading across rows — treat bulk paste there as
   unverified.
2. **Clean** — uploaded sheets are parsed by `load_all_sheets` (Plex query
   results are already DataFrames, no parsing needed); either way
   `consolidate_duplicates` then merges rows sharing the same `Okay PN` per
   fixed rules (sum `Total Revenue` only when it differs across the group,
   sum-or-zero the other cost columns, recompute `Profit or Loss`).
3. **Step 2 (editable)** — user fills in `Total PO $` per row in `st.data_editor`;
   pre-filled from the PO Registry by normalized `Okay PN` when connected.
4. **Step 3 (Final Results)** — `Budget Left = Total PO $ - Total Cost`, plus a
   TOTAL row (`add_totals_row`); downloadable as Excel (`build_excel`).
5. **Report tab (Steps 4-5)** — `build_project_summary` (one row per uploaded
   file, Profit/Loss vs. `Total PO $`), KPI metrics, an Altair chart, an Excel
   export with a native colored bar chart (`build_project_summary_excel`), and
   **Step 5: Open Projects** (`build_open_projects`, joining the hand-maintained
   Project Registry with a live-computed Project Balance per Part Number
   prefix) — gated behind an access code (`st.secrets["open_projects_code"]`,
   falls back to `"4045"` if unset; plain-text `st.text_input(type="password")`
   comparison, not real auth) since it surfaces PO $ / budget figures across
   all projects. The **Step 6: Project Balance
   Trend** chart (same Altair-over-`load_history` approach) was removed from
   the UI 2026-07-11 at the user's request; `log_open_projects_snapshot` still
   runs on every "Generate Final Results" and keeps writing to the History
   tab, so the data keeps accumulating if the chart comes back later.

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

Setup instructions for the Sheets service account: `PO_REGISTRY_SETUP.md`.

## Plex ODBC input (`plex_data.py`)

`fetch_plex_job(part_no_filter)` runs `PLEX_JOB_QUERY_TEMPLATE` (a fixed SQL
query against Plex's cloud reporting views, host `odbc.plex.com`) via a fresh
`pyodbc` connection per call (`get_odbc_connection`). The Part No. filter is
**not** bound as a query parameter — the DataDirect OpenAccess SDK driver
behind this DSN has a confirmed bug where a `?` bound parameter silently
matches zero rows once the query has more than one JOIN (verified live:
identical literal SQL returns correct rows, the parameterized version doesn't).
Instead the filter is validated against `PART_NO_FILTER_RE` (allowlist, no
quotes/SQL syntax) and inlined as an escaped literal via `{part_no_pattern}`.
`fetch_plex_sheets(part_filters)` runs one query per filter and returns
`{filter: DataFrame}`, one entry per filter, mirroring what one uploaded file
used to represent. The connection is `DSN=<odbc_dsn>;UID=<odbc_uid>;PWD=<odbc_pwd>`
— the DSN (set up in Windows' ODBC Data Source Administrator, e.g.
`Plex_ODBC_AVNA`) has the host/port/data source saved, but this driver does
*not* persist UID/PWD from the setup dialog, so both are passed explicitly on
every connection; `odbc_pwd` carries the IAM access token
(`authmethod=iam;accesstoken=...`).

## Deployment topology (verified 2026-07-11)

The Plex ODBC driver (DataDirect OpenAccess SDK) and the `Plex_ODBC_AVNA` DSN
are Windows-only, installed on this laptop via ODBC Data Source Administrator
— not something that travels with git, and not something Streamlit Community
Cloud (Linux, no custom system driver installs beyond apt via `packages.txt`)
can run. Progress DataDirect does publish a Unix/Linux OpenAccess client in
principle, but getting/licensing that binary and wiring up `unixODBC` inside
Streamlit Cloud's sandbox is unverified — treat cloud-direct ODBC as not
currently feasible without further investigation.

So: **live ODBC fetch only works when `main.py` runs on this Windows laptop.**
The Cloud-deployed app can't reach Plex directly, and instead falls back to
reading whatever `daily_plex_export.py` last dropped in the configured Drive
folder — see below.

## Daily Plex export (`daily_plex_export.py`)

Standalone script, run via Windows Task Scheduler (task name **"Plex Daily
Export"**, daily at 7:00 AM, `Execute` = the project's venv
`.venv\Scripts\python.exe`, `Argument` = `daily_plex_export.py`,
`WorkingDirectory` = the repo root so `secrets.toml` resolves). Not part of
the interactive app — imports only from `plex_data.py`, never from `main.py`.

What it does each run: reads every row's `Part Number` column from the
Project Registry sheet (`load_project_registry` + `parse_part_numbers`, so
`'931, 932, 985, 988'` becomes 4 separate filters), queries Plex once per
distinct Part No. via `fetch_plex_job`, and writes one CSV per filter
(`<Part No.>.csv`) into `st.secrets["plex_export_local_dir"]` — a folder
synced by Google Drive Desktop on this laptop. Also deletes any `.csv` in
that folder whose name doesn't match a current Part Number
(`prune_stale_files`), so removed projects don't linger. **Assumes the export
folder is dedicated to this script** — don't point `plex_export_local_dir` at
a folder with other files in it, or pruning will delete them.

`main.py`'s Drive-fetch (`is_drive_configured` / `fetch_drive_files`) reads
the synced copy of that same folder back via the Drive API, using
`st.secrets["plex_export_folder_id"]` (the Drive folder ID, not the local
path — get it from the folder's share link). The Cloud deployment needs its
*own* copy of `plex_export_folder_id` (and the other Plex/Sheets secrets
except `odbc_*`) set in Streamlit Community Cloud's own secrets UI — local
`secrets.toml` never reaches a Cloud deployment, it's gitignored and Cloud
has a separate secrets store entirely.

Run manually to test: `python daily_plex_export.py` (needs the same
`secrets.toml` as `main.py`, resolved relative to CWD).

**Gotcha (verified live 2026-07-10):** "Generate Final Results" always
recomputes `Project Balance` for *every* row in the Project Registry, not just
the Part No. filters loaded in the current session — projects with no
matching sheet loaded get `Project Balance = 0` (see `build_open_projects` /
`build_prefix_balances`), and `log_open_projects_snapshot` then overwrites
*today's* History row for all of them. Clicking "Generate Final Results" after
fetching only one or two Part No. filters (e.g. while testing) will zero out
today's History snapshot for every other project until someone re-runs it with
the full set of filters loaded. Confirmed by directly inspecting the History
sheet after a single-filter test fetch.

## Key normalization helpers

- `normalize_pn` — uppercase, collapse whitespace, strip trailing dash. Needed
  because Plex exports sometimes add a trailing `-` inconsistently
  (`'932 A PD-01-'` vs `'932 A PD-01'`); registry lookups always go through this.
- `extract_pn_prefix` — leading alnum run before the first separator, e.g.
  `'924-1'` and `'924-01'` -> `'924'`. Used to group PNs into a "Project" both
  for the Step 4/Report labels and for Project Balance rollups.

## Config (not in git)

`.streamlit/secrets.toml` (gitignored) holds `gcp_service_account` (service
account JSON fields) and `po_registry_sheet_id` for the Sheets registries;
`odbc_dsn` / `odbc_uid` / `odbc_pwd` for the local-only Plex ODBC connection
(see "Deployment topology" — omit these on a Cloud deployment, since the
driver isn't there anyway); `plex_export_folder_id` (Drive folder ID, used by
`main.py`'s Drive-fetch) and `plex_export_local_dir` (local filesystem path,
used only by `daily_plex_export.py`) for the daily-export bridge; and
`open_projects_code` gating the Open Projects section. Every registry/ODBC/
Drive function checks `is_registry_configured()` / `is_odbc_configured()` /
`is_drive_configured()` first and fails soft (empty registry / next source in
the priority order) rather than raising, so the app works with zero config
too.

## Conventions observed in this codebase

- No test suite currently exists.
- Interactive-app logic lives in `main.py`; shared/importable logic (used by
  both `main.py` and `daily_plex_export.py`) lives in `plex_data.py`. Keep new
  functions in `main.py` unless something outside the interactive app needs
  them too — see "Three files, not one."
- Functions that hit Google Sheets catch broad `Exception` and return a safe
  empty/false value rather than raising, so the UI never hard-crashes on a
  misconfigured or unreachable registry — follow this pattern for any new
  Sheets calls. ODBC calls (`fetch_plex_job`/`fetch_plex_sheets`) intentionally
  do *not* follow this pattern — a failed Plex query surfaces via `st.error`
  in the UI instead, since silently returning empty data here would look like
  "no jobs found" rather than "couldn't connect."
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
