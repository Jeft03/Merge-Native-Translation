# merge_native_translation.py
# What it does:
#   1. Asks you to select a folder
#   2. Finds every .csv and .xlsx file in that folder
#   3. Reads each file into a table (DataFrame)
#   4. Keeps only the Native and Translation columns
#   5. Removes duplicate Native + Translation pairs
#   6. Saves one master file in the same folder
#
# Needs: Python 3 + "pip install pandas openpyxl"

import os
import tkinter as tk
from tkinter import filedialog, messagebox

import pandas as pd

# ---------- Settings you can change ----------
KEEP_COLUMNS = ["Native", "Translation"]   # column names to keep (case and spaces are ignored)
OUTPUT_NAME = "master_native_translation"  # output file name, without extension
OUTPUT_FORMAT = "csv"                      # "csv" (default) or "xlsx"
# ---------------------------------------------


def pick_folder():
    """Step 1: show a 'select folder' window. Works on any user's laptop."""
    root = tk.Tk()
    root.withdraw()                 # hide the empty main window
    root.attributes("-topmost", True)
    folder = filedialog.askdirectory(title="Select the folder with your CSV / Excel files")
    root.destroy()
    return folder


def find_files(folder):
    """Step 2: list all .csv and .xlsx files, skipping our own output and Excel temp files."""
    files = []
    for name in sorted(os.listdir(folder)):
        lower = name.lower()
        if name.startswith("~$"):                 # temp file Excel creates while a file is open
            continue
        if lower.startswith(OUTPUT_NAME.lower()):  # don't read the master file from a previous run
            continue
        if lower.endswith(".csv") or lower.endswith(".xlsx"):
            files.append(os.path.join(folder, name))
    return files


def read_csv_any_encoding(path):
    """Read a CSV/TSV. Tries UTF-8 first, then UTF-16. Comma or tab is detected automatically."""
    for enc in ["utf-8-sig", "utf-16"]:
        try:
            return pd.read_csv(path, encoding=enc, sep=None, engine="python",
                               dtype=str, keep_default_na=False)
        except (UnicodeDecodeError, UnicodeError):
            continue
    raise ValueError("could not read as UTF-8 or UTF-16")


def read_file(path):
    """Step 3: read one file into a list of tables (an Excel file can have several sheets)."""
    if path.lower().endswith(".csv"):
        return [read_csv_any_encoding(path)]
    sheets = pd.read_excel(path, sheet_name=None, dtype=str, keep_default_na=False)
    return list(sheets.values())


def keep_wanted_columns(df):
    """Step 4: keep only Native and Translation. Returns None if the table doesn't have them."""
    # Match column names ignoring upper/lower case and extra spaces, e.g. " native " == "Native"
    lookup = {str(col).strip().lower(): col for col in df.columns}
    found = []
    for wanted in KEEP_COLUMNS:
        key = wanted.strip().lower()
        if key not in lookup:
            return None
        found.append(lookup[key])
    out = df[found].copy()
    out.columns = KEEP_COLUMNS
    return out


def main():
    folder = pick_folder()
    if not folder:
        print("No folder selected. Stopped.")
        return

    files = find_files(folder)
    if not files:
        messagebox.showwarning("Nothing to do", "No .csv or .xlsx files found in:\n" + folder)
        return

    tables = []
    skipped = []
    used_files = set()
    for path in files:
        name = os.path.basename(path)
        try:
            for df in read_file(path):
                trimmed = keep_wanted_columns(df)
                if trimmed is None:
                    skipped.append(name + " (no Native/Translation column)")
                    continue
                tables.append(trimmed)
                used_files.add(name)
                print(f"Read {len(trimmed)} rows from {name}")
        except Exception as e:
            skipped.append(f"{name} ({e})")

    if not tables:
        messagebox.showerror("Failed", "No usable files.\n\nSkipped:\n" + "\n".join(skipped))
        return

    master = pd.concat(tables, ignore_index=True)
    total_rows = len(master)

    # Clean: remove spaces at start/end, drop rows where both cells are empty
    for col in KEEP_COLUMNS:
        master[col] = master[col].astype(str).str.strip()
    master = master[(master[KEEP_COLUMNS[0]] != "") | (master[KEEP_COLUMNS[1]] != "")]

    # Step 5: remove duplicate Native + Translation pairs
    master = master.drop_duplicates(subset=KEEP_COLUMNS)

    # Step 6: save
    if OUTPUT_FORMAT == "xlsx":
        out_path = os.path.join(folder, OUTPUT_NAME + ".xlsx")
        master.to_excel(out_path, index=False)
    else:
        out_path = os.path.join(folder, OUTPUT_NAME + ".csv")
        master.to_csv(out_path, index=False, encoding="utf-8-sig")  # utf-8-sig so Excel shows all languages

    summary = (f"Files read: {len(used_files)}\n"
               f"Rows before dedupe: {total_rows}\n"
               f"Rows in master: {len(master)}\n\n"
               f"Saved to:\n{out_path}")
    if skipped:
        summary += "\n\nSkipped:\n" + "\n".join(skipped)
    print(summary)
    messagebox.showinfo("Done", summary)


if __name__ == "__main__":
    main()
