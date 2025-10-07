import re
import pandas as pd
from pathlib import Path
from tkinter import filedialog, Tk, messagebox
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill


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
    """Return list of task ids: 8 digits or 8 digits + 1 letter (case-insensitive)."""
    if not isinstance(hehe_text, str):
        return []
    return re.findall(r"\b\d{8}[A-Za-z]?\b", hehe_text)


# ---------------- Helper 2 ---------------- #
def distribute_minutes(total_minutes, n_tasks):
    """Even float distribution across tasks."""
    if n_tasks <= 0:
        return []
    per = float(total_minutes) / float(n_tasks)
    return [per] * n_tasks


# ---------------- Helper 3 ---------------- #
def format_task_output(task_order, totals):
    """Return formatted string like 12345678-10.5 mins | 23456789-9.75 mins"""
    return " | ".join(f"{tid}-{totals[tid]:.2f} mins" for tid in task_order)


# ---------------- Helper 4 ---------------- #
def parse_task_minutes_from_cell(cell_value):
    """
    Parse strings like 'count = 3 | 12345678-11.25 mins | 23456789-11.25 mins'
    into {taskid: minutes}. Returns numeric minutes (float).
    """
    out = {}
    if not isinstance(cell_value, str):
        return out
    parts = [p.strip() for p in cell_value.split("|")]
    for part in parts:
        if not part:
            continue
        if part.lower().startswith("count"):
            continue
        if "-" not in part:
            continue
        tid, mins = part.split("-", 1)
        tid = tid.strip()
        m = re.findall(r"[\d\.]+", mins)
        if not m:
            continue
        try:
            val = float(m[0])
        except:
            continue
        out[tid] = out.get(tid, 0.0) + val
    return out


# ---------------- Main Processing ---------------- #
def process_excel(input_path):
    """
    Reads input Excel and produces:
      - Sheet1 with Minutes, EQ, PF, Count.
      - Sheet2 with PF/EQ tasks, minutes, hours, difficulty.
      - Sheet3 with EQ summary (Low/Medium/High by employee).
      Output file: processed_<filename>.xlsx
    """

    df = pd.read_excel(input_path, engine="openpyxl", dtype={"Dates": str})

    # Insert Minutes after Hours
    if "Hours" not in df.columns:
        raise KeyError("Expected column 'Hours' not found in input file.")
    hours_idx = df.columns.get_loc("Hours")
    df.insert(
        hours_idx + 1,
        "Minutes",
        (pd.to_numeric(df["Hours"], errors="coerce").fillna(0) * 60.0).astype(float)
    )

    # Prepare EQ, PF, Count columns
    df["EQ"] = ""
    df["PF"] = ""
    df["Count"] = ""

    # --- Process consecutive rows by (CID, Name, Dates) --- #
    n = len(df)
    i = 0
    while i < n:
        start = i
        j = i + 1
        while j < n and all(df.at[j, col] == df.at[start, col] for col in ("CID", "Name", "Dates")):
            j += 1
        group_idxs = list(range(start, j))

        eq_totals, eq_order = {}, []
        pf_totals, pf_order = {}, []

        for idx in group_idxs:
            hehe_val = df.at[idx, "HEHE"]
            if not isinstance(hehe_val, str):
                continue
            hehe_upper = hehe_val.upper()
            minutes = float(df.at[idx, "Minutes"]) if pd.notna(df.at[idx, "Minutes"]) else 0.0

            # DS/EQ
            if ("DS" in hehe_upper) or ("EQ" in hehe_upper):
                task_ids = parse_task_ids(hehe_val)
                if not task_ids:
                    continue
                per_task = distribute_minutes(minutes, len(task_ids))
                for t, tid in enumerate(task_ids):
                    if tid not in eq_totals:
                        eq_totals[tid] = 0.0
                        eq_order.append(tid)
                    eq_totals[tid] += per_task[t]
                df.at[idx, "EQ"] = f"count = {len(eq_order)} | " + format_task_output(eq_order, eq_totals)
                df.at[idx, "Count"] = len(eq_order)

            # PF
            elif "PF" in hehe_upper:
                task_ids = parse_task_ids(hehe_val)
                if not task_ids:
                    continue
                per_task = distribute_minutes(minutes, len(task_ids))
                for t, tid in enumerate(task_ids):
                    if tid not in pf_totals:
                        pf_totals[tid] = 0.0
                        pf_order.append(tid)
                    pf_totals[tid] += per_task[t]
                df.at[idx, "PF"] = f"count = {len(pf_order)} | " + format_task_output(pf_order, pf_totals)
                df.at[idx, "Count"] = len(pf_order)

        i = j

    # ---------- Build Sheet2 ---------- #
    sheet2_rows = []
    for (cid, name), g in df.groupby(["CID", "Name"], sort=False):
        seen_dates = []
        for d in g["Dates"].tolist():
            if d not in seen_dates:
                seen_dates.append(d)

        pf_tasks, pf_diff = {}, {}
        eq_tasks, eq_diff = {}, {}

        for d in seen_dates:
            rows_date = g[g["Dates"] == d]
            if rows_date.empty:
                continue
            last_row = rows_date.iloc[-1]
            diff_val = last_row.get("Difficulty", "")
            pf_map = parse_task_minutes_from_cell(last_row.get("PF", ""))
            for tid, mins in pf_map.items():
                pf_tasks[tid] = pf_tasks.get(tid, 0.0) + mins
                pf_diff[tid] = diff_val

            eq_map = parse_task_minutes_from_cell(last_row.get("EQ", ""))
            for tid, mins in eq_map.items():
                eq_tasks[tid] = eq_tasks.get(tid, 0.0) + mins
                eq_diff[tid] = diff_val

        pf_items, eq_items = list(pf_tasks.items()), list(eq_tasks.items())
        max_len = max(len(pf_items), len(eq_items))

        for k in range(max_len):
            pf_task, pf_min = (pf_items[k] if k < len(pf_items) else ("", ""))
            eq_task, eq_min = (eq_items[k] if k < len(eq_items) else ("", ""))

            pf_d = pf_diff.get(pf_task, "") if pf_task else ""
            eq_d = eq_diff.get(eq_task, "") if eq_task else ""
            pf_hr = round(pf_min / 60.0, 2) if pf_min != "" else ""
            eq_hr = round(eq_min / 60.0, 2) if eq_min != "" else ""

            sheet2_rows.append([
                cid, name,
                pf_task, round(pf_min, 2) if pf_min != "" else "", pf_hr, pf_d,
                eq_task, round(eq_min, 2) if eq_min != "" else "", eq_hr, eq_d
            ])

        pf_total_min = sum(pf_tasks.values()) if pf_tasks else 0.0
        eq_total_min = sum(eq_tasks.values()) if eq_tasks else 0.0
        pf_total_hr = round(pf_total_min / 60.0, 2) if pf_tasks else ""
        eq_total_hr = round(eq_total_min / 60.0, 2) if eq_tasks else ""

        sheet2_rows.append([
            cid, name,
            f"Total ({len(pf_tasks)})", round(pf_total_min, 2), pf_total_hr, "",
            f"Total ({len(eq_tasks)})", round(eq_total_min, 2), eq_total_hr, ""
        ])
        sheet2_rows.append([""] * 10)

    df_sheet2 = pd.DataFrame(sheet2_rows, columns=[
        "CID", "Name",
        "PF Task", "PF Minutes", "PF Hours", "PF Difficulty",
        "EQ Task", "EQ Minutes", "EQ Hours", "EQ Difficulty"
    ])

    # ---------- Build Sheet3 ---------- #
    sheet3_rows = []
    headers = ["Name", "Difficulty", "Total Tasks", "Total Hrs", "Avg Hrs/Task"]
    base_order = ["Low", "Medium", "High"]

    for (cid, name), g in df_sheet2.groupby(["CID", "Name"], sort=False):
        eq_rows = g[g["EQ Task"].notna() & (g["EQ Task"].astype(str) != "") &
                    (~g["EQ Task"].astype(str).str.startswith("Total"))]

        counts, hrs_sum = {}, {}
        extras = []
        for _, r in eq_rows.iterrows():
            diff = str(r.get("EQ Difficulty", "") or "")
            try:
                hrs = float(r["EQ Hours"]) if r["EQ Hours"] not in ("", None) else 0.0
            except:
                try:
                    mins = float(r["EQ Minutes"]) if r["EQ Minutes"] not in ("", None) else 0.0
                    hrs = mins / 60.0
                except:
                    hrs = 0.0
            counts[diff] = counts.get(diff, 0) + 1
            hrs_sum[diff] = hrs_sum.get(diff, 0.0) + hrs
            if diff not in base_order and diff not in extras:
                extras.append(diff)

        order = base_order + extras
        first_row = True
        for diff in order:
            tcount = counts.get(diff, 0)
            total_hrs = round(hrs_sum.get(diff, 0.0), 2) if tcount else 0.0
            avg_hrs = round(total_hrs / tcount, 2) if tcount else 0.0
            sheet3_rows.append([
                name if first_row else "",
                diff,
                tcount,
                total_hrs,
                avg_hrs
            ])
            first_row = False
        sheet3_rows.append([""] * 5)

    df_sheet3 = pd.DataFrame(sheet3_rows, columns=headers)

    # ---------- Save all ---------- #
    out_path = Path(input_path).with_name(f"processed_{Path(input_path).name}")
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Sheet1", index=False)
        df_sheet2.to_excel(writer, sheet_name="Sheet2", index=False)
        pd.DataFrame([["EQ"]]).to_excel(writer, sheet_name="Sheet3", index=False, header=False)
        df_sheet3.to_excel(writer, sheet_name="Sheet3", startrow=1, index=False)

    wb = load_workbook(out_path)
    ws3 = wb["Sheet3"]
    ws3.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
    cell = ws3["A1"]
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.font = Font(bold=True)
    wb.save(out_path)

    return out_path


# ---------------- Run Script ---------------- #
if __name__ == "__main__":
    try:
        file_path = pick_file_dialog()
        if not file_path:
            messagebox.showinfo("No file chosen", "No input file selected. Exiting.")
            raise SystemExit()
        out_path = process_excel(file_path)
        messagebox.showinfo("Done", f"Processed file saved:\n{out_path}")
    except Exception as e:
        messagebox.showerror("Error", str(e))
        raise
