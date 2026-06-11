import sys
from pathlib import Path

import pandas as pd
from tqdm import tqdm

MAX_SHOWN = 50


def get_files_dir() -> Path:
    # Resolved at call time so Path.cwd() reflects the user's working directory
    return Path.cwd() / "files"


def _normalize(val):
    if isinstance(val, float):
        rounded = round(val, 10)
        if rounded == int(rounded):
            return int(rounded)
        return rounded
    return val


def load_sheet_data(path: Path, sheet_name: str) -> list[tuple]:
    df = pd.read_excel(path, sheet_name=sheet_name, engine='openpyxl', header=None, dtype=object)
    df = df.where(df.notna(), None)
    return [tuple(_normalize(v) for v in row) for row in df.itertuples(index=False, name=None)]


def compare_sheets(old_file: Path, new_file: Path, sheet_name: str) -> bool:
    print(f"  Loading '{sheet_name}' from old file...", flush=True)
    old_rows = load_sheet_data(old_file, sheet_name)
    print(f"  Loading '{sheet_name}' from new file...", flush=True)
    new_rows = load_sheet_data(new_file, sheet_name)

    headers = old_rows[0] if old_rows else None
    has_headers = headers is not None

    differences: list[tuple[int, tuple, tuple]] = []

    max_rows = max(len(old_rows), len(new_rows))
    print(f"  Comparing {max_rows} rows...", flush=True)
    with tqdm(total=max_rows, desc=f"  Sheet '{sheet_name}'", unit="row", leave=True, file=sys.stdout) as bar:
        for i in range(max_rows):
            old_row = old_rows[i] if i < len(old_rows) else None
            new_row = new_rows[i] if i < len(new_rows) else None
            if old_row != new_row:
                differences.append((i + 1, old_row, new_row))
            bar.update(1)

    if not differences:
        print(f"  [OK] Sheet '{sheet_name}': no differences found ({len(old_rows)} rows)", flush=True)
        return True

    print(f"  [DIFF] Sheet '{sheet_name}': {len(differences)} row(s) differ (showing first {min(MAX_SHOWN, len(differences))})\n", flush=True)

    for row_num, old_row, new_row in differences[:MAX_SHOWN]:
        label = f"Row {row_num}"
        if has_headers and row_num == 1:
            label += " (header)"
        print(f"  --- {label} ---")

        if old_row is None:
            print(f"    old : <row missing>")
            print(f"    new : {new_row}\n")
            continue
        if new_row is None:
            print(f"    old : {old_row}")
            print(f"    new : <row missing>\n")
            continue

        max_cols = max(len(old_row), len(new_row))
        for col_idx in range(max_cols):
            old_val = old_row[col_idx] if col_idx < len(old_row) else "<missing>"
            new_val = new_row[col_idx] if col_idx < len(new_row) else "<missing>"

            if old_val != new_val:
                col_label = f"col {col_idx + 1}"
                if has_headers and headers and col_idx < len(headers) and headers[col_idx]:
                    col_label = str(headers[col_idx])
                print(f"    {col_label:20s}  old: {repr(old_val)}")
                print(f"    {'':20s}  new: {repr(new_val)}")
        print()

    sys.stdout.flush()
    return False


def run() -> int:
    try:
        files_dir = get_files_dir()
        old_file = files_dir / "old_file.xlsx"
        new_file = files_dir / "new_file.xlsx"

        for path, label in [(old_file, "old_file.xlsx"), (new_file, "new_file.xlsx")]:
            if not path.exists():
                print(f"ERROR: {label} not found at {path}", flush=True)
                return 1

        print("Reading sheet names from old_file.xlsx...", flush=True)
        old_sheets = set(pd.ExcelFile(old_file, engine='openpyxl').sheet_names)
        print("Reading sheet names from new_file.xlsx...", flush=True)
        new_sheets = set(pd.ExcelFile(new_file, engine='openpyxl').sheet_names)

        print("=" * 60)
        print("Comparing:")
        print(f"  old: {old_file.name}  ({len(old_sheets)} sheet(s))")
        print(f"  new: {new_file.name}  ({len(new_sheets)} sheet(s))")
        print("=" * 60)
        sys.stdout.flush()

        all_ok = True

        only_in_old = old_sheets - new_sheets
        only_in_new = new_sheets - old_sheets
        if only_in_old:
            print(f"\n[DIFF] Sheets only in old_file: {only_in_old}", flush=True)
            all_ok = False
        if only_in_new:
            print(f"\n[DIFF] Sheets only in new_file: {only_in_new}", flush=True)
            all_ok = False

        shared_sheets = old_sheets & new_sheets
        if not shared_sheets:
            shared_sheets = set(wb_old.sheetnames)

        print()
        for sheet in sorted(shared_sheets):
            ok = compare_sheets(old_file, new_file, sheet)
            if not ok:
                all_ok = False

        print("=" * 60)
        if all_ok:
            print("RESULT: Files are identical.")
        else:
            print("RESULT: Files differ. See details above.")
        sys.stdout.flush()

        return 0 if all_ok else 1

    except Exception as e:
        print(f"\nUNEXPECTED ERROR: {e}", file=sys.stderr, flush=True)
        raise


if __name__ == "__main__":
    run()
