"""Scheduled script (not part of the interactive Streamlit app): reads the
distinct Part Numbers out of the Project Registry Google Sheet, queries Plex
over ODBC for each one, and writes the results as CSV files into the local
Google Drive-synced folder (plex_export_local_dir in secrets.toml). Drive
Desktop syncs those files to the cloud, and main.py — when it can't reach
Plex directly, e.g. deployed on Streamlit Community Cloud — reads them back
via the Drive API (plex_export_folder_id in secrets.toml).

Only works when run on a machine with the Plex ODBC driver and DSN installed
(see plex_data.get_odbc_connection). Meant to run daily via Windows Task
Scheduler, "Start in" set to this file's directory so secrets.toml resolves.

Run manually: python daily_plex_export.py
"""

from pathlib import Path

import streamlit as st

from plex_data import fetch_plex_job, load_project_registry, parse_part_numbers

EXPORT_DIR = Path(st.secrets.get("plex_export_local_dir", r"C:\Users\upandey\Desktop\Gdrive query"))


def collect_part_numbers() -> list[str]:
    """Distinct Part Number prefixes across every row of the Project
    Registry's 'Part Number' column (which can hold several comma/space
    separated values per row, e.g. '931, 932, 985, 988')."""
    registry = load_project_registry()
    if registry.empty:
        return []
    numbers = []
    for value in registry["Part Number"].dropna():
        for pn in parse_part_numbers(value):
            if pn not in numbers:
                numbers.append(pn)
    return numbers


def prune_stale_files(current_part_numbers: list[str]) -> None:
    """Removes CSVs left over from Part Numbers no longer in the Project
    Registry, so main.py's Drive-fetch doesn't keep loading yesterday's
    now-removed projects. Only touches files this script would itself have
    written (<Part No.>.csv) — safe as long as EXPORT_DIR is dedicated to
    this export and not shared with other manually-placed files."""
    keep = {f"{pn}.csv" for pn in current_part_numbers}
    for path in EXPORT_DIR.glob("*.csv"):
        if path.name not in keep:
            path.unlink()
            print(f"  removed stale {path.name}")


def main() -> None:
    EXPORT_DIR.mkdir(parents=True, exist_ok=True)
    part_numbers = collect_part_numbers()
    if not part_numbers:
        print("No Part Numbers found in the Project Registry — nothing to export.")
        return

    prune_stale_files(part_numbers)

    print(f"Exporting {len(part_numbers)} Part No. filter(s) to {EXPORT_DIR}")
    for part_no in part_numbers:
        try:
            df = fetch_plex_job(part_no)
        except Exception as e:
            print(f"  {part_no}: FAILED — {e}")
            continue
        out_path = EXPORT_DIR / f"{part_no}.csv"
        df.to_csv(out_path, index=False)
        print(f"  {part_no}: {len(df)} row(s) -> {out_path.name}")


if __name__ == "__main__":
    main()
