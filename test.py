import re
import pandas as pd
from pathlib import Path
from tkinter import filedialog, Tk, messagebox
from collections import defaultdict
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill


# ---------------- File Picker ---------------- #
def pick_file_dialog():
    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    file_path = filedialog.askopenfilename(
        title="Select Excel file",
        filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
    )
    root.destroy()
    return file_path


# ---------------- Helper 1 ---------------- #
def parse_task_ids(hehe_text):
    """Extract task IDs (8 digits or 8 digits + 1 letter)."""
    if not isinstance(hehe_text, str):
        return []
    return re.findall(r"\b\d{8}[A-Z]?\b", hehe_text)


# ---------------- Helper 2 ---------------- #
def distribute_minutes(total_minutes, n_tasks):
    """Distribute minutes equally among tasks."""
    if n_tasks <= 0:
        return []
    per_task = total_minutes / n_tasks
    return [per_task] * n_tasks


# ---------------- Helper 3 ---------------- #
def format_task_output(task_ids, per_task_minutes):
    """Format task-time output."""
    return " ".join([f"{tid}-{mins:.2f} mins" for tid, mins in zip(task_ids, per_task_minutes)])


# ---------------- Helper 4 ---------------- #
def parse_task_minutes_from_cell(cell_value):
    """Extract (task_id, minutes) pairs from formatted EQ/PF cells."""
    if not isinstance(cell_value, str) or not cell_value.strip():
        return []
    pairs = []
    for token in cell_value.split():
        if "-" in token:
            tid, mins_part = token.split("-")
            mins = re.findall(r"[\d.]+", mins_part)
            if mins:
                pairs.append((tid, float(mins[0])))
    return pairs


# ---------------- Main Processing ---------------- #
def process_excel(input_path):
    wb = load_workbook(input_path)
    sheet1 = wb.active
    df = pd.read_excel(input_path)

    merged = defaultdict(lambda: {"Minutes": 0, "Hehe": "", "Difficulty": "", "EQ": "", "PF": ""})

    # ---------- Merge consecutive rows ---------- #
    for i in range(len(df)):
        cid, name, date = df.loc[i, "CID"], df.loc[i, "Name"], df.loc[i, "Dates"]
        hehe, minutes, difficulty = (
            df.loc[i, "HEHE"],
            df.loc[i, "Minutes"],
            df.loc[i, "Difficulty"],
        )

        key = (cid, name, date)
        merged[key]["Minutes"] += minutes if not pd.isna(minutes) else 0
        merged[key]["Hehe"] = hehe
        merged[key]["Difficulty"] = difficulty

        # --- EQ (DS) Handling --- #
        if isinstance(hehe, str) and "DS" in hehe.upper():
            tasks = parse_task_ids(hehe)
            if tasks:
                per_task = distribute_minutes(merged[key]["Minutes"], len(tasks))
                merged[key]["EQ"] = format_task_output(tasks, per_task)

        # --- PF Handling --- #
        elif isinstance(hehe, str) and "PF" in hehe.upper():
            tasks = parse_task_ids(hehe)
            if tasks:
                per_task = distribute_minutes(merged[key]["Minutes"], len(tasks))
                merged[key]["PF"] = format_task_output(tasks, per_task)

    # ---------- Update Sheet 1 ---------- #
    df["EQ"] = ""
    df["PF"] = ""
    df["Count"] = ""

    for i in range(len(df)):
        cid, name, date, hehe = df.loc[i, ["CID", "Name", "Dates", "HEHE"]]
        key = (cid, name, date)
        df.loc[i, "EQ"] = merged[key]["EQ"]
        df.loc[i, "PF"] = merged[key]["PF"]

        # Count
        if isinstance(hehe, str) and ("DS" in hehe.upper() or "PF" in hehe.upper()):
            tasks = parse_task_ids(hehe)
            df.loc[i, "Count"] = len(tasks)
        else:
            df.loc[i, "Count"] = ""

    # Write Sheet 1
    for col_num, col_name in enumerate(df.columns, 1):
        sheet1.cell(row=1, column=col_num).value = col_name
    for i, row in df.iterrows():
        for j, value in enumerate(row, 1):
            sheet1.cell(row=i + 2, column=j).value = value

    # ---------------- Sheet 2 ---------------- #
    sheet2 = wb.create_sheet("Sheet2")
    sheet2.append(["CID", "Name", "PF Task", "PF Minutes", "PF Diff",
                   "EQ Task", "EQ Minutes", "EQ Hours", "EQ Diff"])

    sheet2_rows = []
    for _, row in df.iterrows():
        cid, name, difficulty = row["CID"], row["Name"], row["Difficulty"]

        pf_items = parse_task_minutes_from_cell(row["PF"])
        eq_items = parse_task_minutes_from_cell(row["EQ"])

        max_len = max(len(pf_items), len(eq_items))
        for i in range(max_len):
            pf_task, pf_min = (pf_items[i] if i < len(pf_items) else ("", ""))
            eq_task, eq_min = (eq_items[i] if i < len(eq_items) else ("", ""))

            pf_diff = difficulty if pf_task else ""
            eq_diff = difficulty if eq_task else ""
            eq_hrs = round(eq_min / 60, 2) if eq_min else ""

            sheet2_rows.append([
                cid, name,
                pf_task, pf_min, pf_diff,
                eq_task, eq_min, eq_hrs, eq_diff
            ])

    for r in sheet2_rows:
        sheet2.append(r)

    # ---------------- Sheet 3 ---------------- #
    sheet3 = wb.create_sheet("Sheet3")
    sheet3.append(["EQ"])
    sheet3.append(["Name", "Difficulty", "Total Tasks", "Total Hrs", "Avg Hrs/Task"])

    df2 = pd.DataFrame(sheet2_rows, columns=["CID", "Name", "PF Task", "PF Minutes", "PF Diff",
                                             "EQ Task", "EQ Minutes", "EQ Hours", "EQ Diff"])
    df2 = df2[df2["EQ Task"] != ""]

    for name, group in df2.groupby("Name"):
        sheet3.append([name])
        for diff, g in group.groupby("EQ Diff"):
            if not diff:
                continue
            total_tasks = len(g)
            total_hrs = g["EQ Hours"].astype(float).sum()
            avg_hrs = total_hrs / total_tasks if total_tasks else 0
            sheet3.append(["", diff, total_tasks, f"{total_hrs:.2f}", f"{avg_hrs:.2f}"])
        sheet3.append([])

    # ---------------- Highlight "Total" Rows ---------------- #
    yellow_fill = PatternFill(start_color="FFFFFF00", end_color="FFFFFF00", fill_type="solid")
    bold_font = Font(bold=True)
    for sheet in [sheet2, sheet3]:
        for row in sheet.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.strip().lower().startswith("total"):
                    cell.font = bold_font
                    cell.fill = yellow_fill

    # ---------------- Save ---------------- #
    wb.save(input_path)
    print(f"✅ Processing complete! Updated workbook saved: {input_path}")


# ---------------- Run ---------------- #
if __name__ == "__main__":
    try:
        file_path = pick_file_dialog()
        if not file_path:
            messagebox.showinfo("No file chosen", "No input file selected. Exiting.")
            raise SystemExit()
        process_excel(file_path)
        messagebox.showinfo("Done", f"Processed file saved:\n{file_path}")
    except Exception as e:
        messagebox.showerror("Error", str(e))
        raise
