"""
Excel Processor (Dates preserved as text)
-----------------------------------------
- Adds "Minutes" column (Hours * 60) right after Hours.
- Adds "EQ" column at the end.
- Preserves Dates column exactly as in input (string).
- For consecutive rows with same (CID, Name, Dates):
    * Only rows whose HEHE contains "DS" are considered.
    * Task IDs are parsed (numbers with >=6 digits).
    * Row's minutes are split evenly among tasks.
    * Totals are accumulated across the group.
    * Each DS row shows the cumulative totals in EQ.
"""

import re
import pandas as pd
from pathlib import Path
from tkinter import filedialog, Tk, messagebox

# ---------------- Helpers ---------------- #

def pick_file_dialog():
    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    file_path = filedialog.askopenfilename(
        title="Select input Excel file",
        filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
    )
    root.destroy()
    return file_path

def parse_task_ids(hehe_text, min_len=6):
    """Extract long numeric task IDs (>= min_len digits) from the HEHE string."""
    if not isinstance(hehe_text, str):
        return []
    parts = [p.strip() for p in hehe_text.strip("{} ").split("|")]
    for part in parts:
        ids = re.findall(r"\d+", part)
        ids = [tid for tid in ids if len(tid) >= min_len]
        if ids:
            return ids
    ids = re.findall(r"\d+", hehe_text)
    return [tid for tid in ids if len(tid) >= min_len]

def distribute_minutes(total_minutes, n_tasks):
    if n_tasks <= 0:
        return []
    base = total_minutes // n_tasks
    rem = total_minutes % n_tasks
    return [base + (1 if i < rem else 0) for i in range(n_tasks)]

def format_EQ(task_order, totals):
    return " | ".join(f"{tid}-{totals[tid]} mins" for tid in task_order)

# ---------------- Main Processing ---------------- #

def process_excel(input_path):
    # Read, but force Dates column as string to preserve original format
    df = pd.read_excel(
        input_path,
        engine="openpyxl",
        dtype={"Dates": str}
    )

    # Add Minutes next to Hours
    if "Hours" not in df.columns:
        raise KeyError("Expected 'Hours' column not found")
    hours_idx = df.columns.get_loc("Hours")
    df.insert(
        hours_idx + 1,
        "Minutes",
        (pd.to_numeric(df["Hours"], errors="coerce").fillna(0) * 60).astype(int)
    )

    # Add EQ at the end
    df["EQ"] = ""

    # Group consecutive rows by (CID, Name, Dates)
    n = len(df)
    i = 0
    while i < n:
        start = i
        j = i + 1
        while j < n and all(df.at[j, col] == df.at[start, col] for col in ["CID", "Name", "Dates"]):
            j += 1
        group_idxs = list(range(start, j))

        task_totals = {}
        task_order = []

        for idx in group_idxs:
            hehe_val = df.at[idx, "HEHE"]
            if not isinstance(hehe_val, str) or "DS" not in hehe_val:
                continue  # leave EQ blank

            task_ids = parse_task_ids(hehe_val, min_len=6)
            if not task_ids:
                continue

            minutes = df.at[idx, "Minutes"]
            per_task = distribute_minutes(minutes, len(task_ids))

            for t, tid in enumerate(task_ids):
                if tid not in task_totals:
                    task_totals[tid] = 0
                    task_order.append(tid)
                task_totals[tid] += per_task[t]

            df.at[idx, "EQ"] = format_EQ(task_order, task_totals)

        i = j

    # Save output
    out_path = Path(input_path).with_name(f"processed_{Path(input_path).name}")
    df.to_excel(out_path, index=False, engine="openpyxl")
    return out_path

# ---------------- Run ---------------- #

if __name__ == "__main__":
    try:
        file_path = pick_file_dialog()
        if not file_path:
            messagebox.showinfo("No file chosen", "No input file selected. Exiting.")
            raise SystemExit()
        out_file = process_excel(file_path)
        messagebox.showinfo("Done", f"Processed file saved:\n{out_file}")
    except Exception as e:
        messagebox.showerror("Error", str(e))
        raise
