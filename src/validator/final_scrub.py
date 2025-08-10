# src/validator/final_scrub.py
import sys
import os
import re
import datetime
from tkinter import Tk, filedialog

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from dateutil.parser import parse as date_parse

# Suffixes to preserve in proper-case names
SUFFIXES = ['Jr.', 'Sr.', 'II', 'III', 'IV', 'V']


def proper_case_name(name):
    if pd.isna(name):
        return name
    name = str(name).strip()

    # Handle names like "Doe, John Jr." → "John Doe Jr."
    if ',' in name:
        parts = name.split(',')
        if len(parts) >= 2:
            last = parts[0].strip()
            rest = parts[1].strip()
            name = f"{rest} {last}"

    tokens = name.split()
    out = []
    for token in tokens:
        if token in SUFFIXES:
            out.append(token)
        elif '-' in token:
            out.append('-'.join(t.capitalize() for t in token.split('-')))
        else:
            out.append(token.capitalize())
    return ' '.join(out)


def standardize_race(value):
    if pd.isna(value):
        return value
    value = str(value).strip().upper()
    mapping = {
        'W': 'White', 'WHITE': 'White',
        'B': 'Black', 'BLACK': 'Black',
        'H': 'Hispanic', 'HISPANIC': 'Hispanic',
        'O': 'Other', 'OTHER': 'Other'
    }
    return mapping.get(value, value.capitalize())


def transform_address(value):
    if pd.isna(value):
        return value
    val = str(value).strip()
    val = re.sub(r'\bHighway\b', 'Hwy', val, flags=re.IGNORECASE)
    val = re.sub(r'\bBoulevard\b', 'Blvd', val, flags=re.IGNORECASE)
    val = re.sub(r'\bNortheast\b', 'NE', val, flags=re.IGNORECASE)
    val = re.sub(r'\bSoutheast\b', 'SE', val, flags=re.IGNORECASE)
    val = re.sub(r'\bNorthwest\b', 'NW', val, flags=re.IGNORECASE)
    val = re.sub(r'\bSouthwest\b', 'SW', val, flags=re.IGNORECASE)
    return val


def clean_excel_file(file_path):
    """
    Cleans a single Excel file in-place.
    Returns a dict of stats and prints human-friendly lines.
    """
    stats = {
        "file": file_path,
        "start_rows": 0,
        "removed_disposed": 0,
        "removed_814_n_kentucky": 0,
        "removed_180_e_central": 0,
        "removed_general_delivery": 0,
        "duplicates_removed": 0,
        "rows_highlighted_today": 0,
        "end_rows": 0,
        "ok": False,
        "error": None,
    }

    try:
        df = pd.read_excel(file_path)
        stats["start_rows"] = len(df)

        # Clean Name
        if 'Name' in df.columns:
            df['Name'] = df['Name'].apply(proper_case_name)
            print("🧼 Cleaned 'Name' column")

        # Standardize Race
        if 'Race' in df.columns:
            df['Race'] = df['Race'].apply(standardize_race)
            print("🧬 Standardized 'Race' column")

        # Remove Status == Disposed
        if 'Status' in df.columns:
            before = len(df)
            df = df[~df['Status'].astype(str).str.strip().str.lower().eq('disposed')]
            stats["removed_disposed"] = before - len(df)
            if stats["removed_disposed"]:
                print(f"🗑️ Removed {stats['removed_disposed']} row(s) with Status 'Disposed'")

        # Clean Address 1
        if 'Address 1' in df.columns:
            # Remove specific addresses
            before = len(df)
            df = df[~df['Address 1'].astype(str).str.strip().str.lower().eq("814 north kentucky avenue")]
            stats["removed_814_n_kentucky"] = before - len(df)
            if stats["removed_814_n_kentucky"]:
                print(f"🏠 Removed {stats['removed_814_n_kentucky']} row(s) with '814 North Kentucky Avenue'")

            before = len(df)
            df = df[~df['Address 1'].astype(str).str.strip().str.lower().eq("180 east central avenue")]
            stats["removed_180_e_central"] = before - len(df)
            if stats["removed_180_e_central"]:
                print(f"🏢 Removed {stats['removed_180_e_central']} row(s) with '180 East Central Avenue'")

            before = len(df)
            df = df[~df['Address 1'].astype(str).str.lower().str.contains("general delivery")]
            stats["removed_general_delivery"] = before - len(df)
            if stats["removed_general_delivery"]:
                print(f"📦 Removed {stats['removed_general_delivery']} row(s) with 'General Delivery'")

            # Standardize keywords
            df['Address 1'] = df['Address 1'].apply(transform_address)
            print("📫 Standardized keywords in 'Address 1'")

        # Remove duplicate addresses (keep oldest Capture Date)
        if 'Address 1' in df.columns and 'Capture Date' in df.columns:
            df['Capture Date'] = pd.to_datetime(df['Capture Date'], errors='coerce').dt.date
            before = len(df)
            df = df.sort_values(by=['Address 1', 'Capture Date'])
            df = df.drop_duplicates(subset=['Address 1'], keep='first')
            stats["duplicates_removed"] = before - len(df)
            if stats["duplicates_removed"]:
                print(f"🔁 Removed {stats['duplicates_removed']} duplicate address row(s), kept oldest by Capture Date")

        # Reorder columns
        desired_order = [
            "Name", "Address 1", "Address 2", "City", "State", "Zip",
            "Case Number", "Status", "Sex", "Race", "Phone", "Public Defender"
        ]
        existing = [c for c in desired_order if c in df.columns]
        remaining = [c for c in df.columns if c not in existing]
        df = df[existing + remaining]
        print("📑 Reordered columns")

        # Final sort by Capture Date (oldest to newest)
        if 'Capture Date' in df.columns:
            df['Capture Date'] = pd.to_datetime(df['Capture Date'], errors='coerce')
            df = df.sort_values(by='Capture Date', ascending=True)
            print("📅 Sorted rows by 'Capture Date' (oldest to newest)")

        # Save cleaned data (in-place)
        df.to_excel(file_path, index=False)
        stats["end_rows"] = len(df)
        print("💾 Saved cleaned data to Excel")

        # Open workbook for formatting
        wb = load_workbook(file_path)
        ws = wb.active

        # Format Zip as text
        zip_col_index = None
        for idx, cell in enumerate(ws[1], start=1):
            if str(cell.value).strip().lower() == 'zip':
                zip_col_index = idx
                break
        if zip_col_index:
            for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                if len(row) >= zip_col_index:
                    row[zip_col_index - 1].number_format = "@"
            print("🏷️ Formatted 'Zip' column as text")

        # Highlight Capture Date == today
        date_col_index = None
        for idx, cell in enumerate(ws[1], start=1):
            if str(cell.value).strip().lower() == 'capture date':
                date_col_index = idx
                break

        rows_highlighted = 0
        green = PatternFill(start_color="CCFFCC", end_color="CCFFCC", fill_type="solid")
        today = datetime.date.today()

        if date_col_index:
            for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                if len(row) < date_col_index:
                    continue
                cell = row[date_col_index - 1]
                val = cell.value
                try:
                    if isinstance(val, datetime.datetime):
                        date_val = val.date()
                    elif isinstance(val, datetime.date):
                        date_val = val
                    elif isinstance(val, str):
                        date_val = date_parse(val).date()
                    else:
                        continue
                    if date_val == today:
                        for c in row:
                            c.fill = green
                        rows_highlighted += 1
                except Exception:
                    continue
            stats["rows_highlighted_today"] = rows_highlighted
            print(f"✅ Highlighted {rows_highlighted} row(s) with today's Capture Date")
        else:
            print("⚠️ 'Capture Date' column not found — skipping highlighting")

        wb.save(file_path)
        print(f"🎉 File cleaned and saved: {file_path}")

        stats["ok"] = True
        return stats

    except Exception as e:
        stats["error"] = str(e)
        print(f"❌ Error processing file: {e}")
        return stats


def _print_banner(title):
    print("\n" + "=" * 72)
    print(title)
    print("=" * 72)


def main(argv=None):
    argv = argv if argv is not None else sys.argv[1:]

    # Get files: CLI args or picker
    if argv:
        file_paths = [p for p in argv if os.path.isfile(p)]
        missing = [p for p in argv if not os.path.isfile(p)]
        if missing:
            _print_banner("⚠️  Skipping missing paths")
            for m in missing:
                print(m)
            print()
    else:
        Tk().withdraw()
        file_paths = filedialog.askopenfilenames(
            title="Select Excel Files",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )

    if not file_paths:
        print("❌ No files selected.")
        return 1

    all_stats = []
    for fp in file_paths:
        _print_banner(f"📄 Processing: {fp}")
        stats = clean_excel_file(fp)
        # Echo a compact summary per file
        if stats["ok"]:
            print("-- Summary --")
            print(f"Start rows: {stats['start_rows']}")
            print(f"Removed Disposed: {stats['removed_disposed']}")
            print(f"Removed '814 North Kentucky Avenue': {stats['removed_814_n_kentucky']}")
            print(f"Removed '180 East Central Avenue': {stats['removed_180_e_central']}")
            print(f"Removed 'General Delivery': {stats['removed_general_delivery']}")
            print(f"Duplicate addresses removed: {stats['duplicates_removed']}")
            print(f"Rows highlighted today: {stats['rows_highlighted_today']}")
            print(f"End rows: {stats['end_rows']}")
        else:
            print("-- Summary --")
            print("FAILED:", stats.get("error", "Unknown error"))
        all_stats.append(stats)

    # Optional overall summary
    _print_banner("📊 Overall Summary")
    ok = sum(1 for s in all_stats if s["ok"])
    fail = len(all_stats) - ok
    print(f"Files processed: {len(all_stats)} (ok: {ok}, failed: {fail})")
    return 0 if fail == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
