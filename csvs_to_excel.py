#!/usr/bin/env python3
"""
csvs_to_excel.py

Reads one or more CSV files and writes them to a single Excel workbook with:
 - a sheet per CSV (sheet name derived from filename)
 - an Index sheet with basic row counts
 - header formatting, autofilter, freeze panes, and adjusted column widths

Usage:
  python csvs_to_excel.py output.xlsx Speed_Limits.csv Traffic_Volumes.csv

If no CSV paths are given, this script will try to find
Speed_Limits.csv and Traffic_Volumes.csv in the current directory.
"""
import sys
from pathlib import Path
import pandas as pd

def auto_widths(df, min_width=10, max_width=50):
    widths = []
    for col in df.columns:
        max_len = df[col].astype(str).map(len).max() if not df.empty else 0
        header_len = len(str(col))
        w = max(header_len, max_len, min_width)
        w = min(w, max_width)
        widths.append(w + 2)  # small padding
    return widths

def write_workbook(out_path: Path, csv_paths):
    csv_paths = [Path(p) for p in csv_paths]
    writer = pd.ExcelWriter(out_path, engine='xlsxwriter', datetime_format='yyyy-mm-dd', date_format='yyyy-mm-dd')
    workbook = writer.book

    summary_rows = []

    for csv_path in csv_paths:
        if not csv_path.exists():
            print(f"Warning: {csv_path} not found, skipping.")
            continue
        try:
            df = pd.read_csv(csv_path)
        except Exception as e:
            print(f"Error reading {csv_path}: {e}. Skipping.")
            continue

        sheet_name = csv_path.stem[:31]  # Excel sheet name limit
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        worksheet = writer.sheets[sheet_name]

        # Format header
        header_fmt = workbook.add_format({'bold': True, 'bg_color': '#DCE6F1', 'border': 1})
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_fmt)

        # Autofilter and freeze pane
        worksheet.autofilter(0, 0, len(df.index), max(len(df.columns)-1, 0))
        worksheet.freeze_panes(1, 0)

        # Set column widths
        widths = auto_widths(df)
        for i, w in enumerate(widths):
            worksheet.set_column(i, i, w)

        summary_rows.append({
            'sheet': sheet_name,
            'rows': len(df),
            'columns': len(df.columns),
            'source_file': str(csv_path.name)
        })
        print(f"Wrote sheet: {sheet_name} ({len(df)} rows x {len(df.columns)} cols)")

    # Write Index/Summary sheet
    if summary_rows:
        idx_df = pd.DataFrame(summary_rows)
        idx_sheet = 'Index'
        idx_df.to_excel(writer, sheet_name=idx_sheet, index=False)
        ws = writer.sheets[idx_sheet]
        idx_header_fmt = workbook.add_format({'bold': True, 'bg_color': '#F2F2F2', 'border': 1})
        for col_num, value in enumerate(idx_df.columns.values):
            ws.write(0, col_num, value, idx_header_fmt)
        ws.autofilter(0, 0, len(idx_df), max(len(idx_df.columns)-1, 0))
        ws.freeze_panes(1, 0)
        for i, w in enumerate(auto_widths(idx_df)):
            ws.set_column(i, i, w)
    else:
        print("No CSVs were written; workbook will not contain an Index sheet.")

    writer.close()
    print(f"Saved workbook: {out_path}")

def main():
    if len(sys.argv) < 2:
        print("Usage: python csvs_to_excel.py output.xlsx [csv1.csv csv2.csv ...]")
        sys.exit(1)

    out = Path(sys.argv[1])
    csvs = sys.argv[2:]
    if not csvs:
        # default to the two files present in your repo root
        csvs = ['Speed_Limits.csv', 'Traffic_Volumes.csv']
    write_workbook(out, csvs)

if __name__ == '__main__':
    main()
