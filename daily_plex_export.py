"""Scheduled script (not part of the interactive Streamlit app): reads the
Part Number list out of Parts.xlsx (st.secrets["parts_list_path"]), queries
Plex over ODBC for each one, and writes the results as CSV files into the
local Google Drive-synced folder (plex_export_local_dir in secrets.toml).
A '/'-joined row in Parts.xlsx (e.g. '985/931/988/930/931M') is one family —
queried per number but merged into a single CSV, same convention main.py
uses for its live Plex fetch (see plex_data.load_part_groups). Drive Desktop
syncs those files to the cloud, and main.py — when it can't reach Plex
directly, e.g. deployed on Streamlit Community Cloud — reads them back via
the Drive API (plex_export_folder_id in secrets.toml).

Only works when run on a machine with the Plex ODBC driver and DSN installed
(see plex_data.get_odbc_connection). Meant to run daily via Windows Task
Scheduler, "Start in" set to this file's directory so secrets.toml resolves.

Run manually: python daily_plex_export.py
"""

from pathlib import Path

import streamlit as st

from plex_data import fetch_plex_groups, load_part_groups

EXPORT_DIR = Path(st.secrets.get("plex_export_local_dir", r"C:\Users\upandey\Desktop\Gdrive query"))


def group_filename(group: list[str]) -> str:
    return "_".join(group) + ".csv"


def prune_stale_files(current_groups: list[list[str]]) -> None:
    """Removes CSVs left over from groups no longer in Parts.xlsx, so
    main.py doesn't keep loading yesterday's now-removed projects. Only
    touches files this script would itself have written — safe as long as
    EXPORT_DIR is dedicated to this export and not shared with other
    manually-placed files."""
    keep = {group_filename(group) for group in current_groups}
    for path in EXPORT_DIR.glob("*.csv"):
        if path.name not in keep:
            path.unlink()
            print(f"  removed stale {path.name}")


def main() -> None:
    EXPORT_DIR.mkdir(parents=True, exist_ok=True)
    groups = load_part_groups()
    if not groups:
        print("No Part Numbers found in Parts.xlsx — nothing to export.")
        return

    prune_stale_files(groups)

    print(f"Exporting {len(groups)} group(s) to {EXPORT_DIR}")
    for group in groups:
        try:
            sheets = fetch_plex_groups([group])
        except Exception as e:
            print(f"  {'/'.join(group)}: FAILED — {e}")
            continue
        df = next(iter(sheets.values()))
        out_path = EXPORT_DIR / group_filename(group)
        df.to_csv(out_path, index=False)
        print(f"  {'/'.join(group)}: {len(df)} row(s) -> {out_path.name}")


if __name__ == "__main__":
    main()
