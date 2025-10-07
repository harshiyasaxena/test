def process_excel(input_path):
    """
    Minimal-change, self-contained process_excel function that:
      - Reads input Excel (Sheet1) and computes Minutes, EQ, PF, Count as before.
      - Builds Sheet2 as a side-by-side PF/EQ summary (CID, Name, PF Task, PF Minutes, PF Hours, PF Difficulty,
        EQ Task, EQ Minutes, EQ Hours, EQ Difficulty) using last-row cumulative values per date.
      - Builds Sheet3: for each employee (CID+Name) it aggregates EQ tasks from Sheet2 by Difficulty (Low/Medium/High)
        and produces rows:
            Name | Difficulty | Total Tasks | Total Hrs | Avg Hrs/Task
        (Name is printed only on the first difficulty row of each employee).
      - Writes Sheet1, Sheet2 and Sheet3 to processed_<inputfilename>.xlsx next to the input file.
    """

    import re
    import pandas as pd
    from pathlib import Path
    from openpyxl import load_workbook
    from openpyxl.styles import Font, Alignment, PatternFill

    # ---------- Small helpers (kept minimal) ----------
    def parse_task_ids(hehe_text):
        """Return list of task ids: 8 digits or 8 digits + 1 letter (case-insensitive)."""
        if not isinstance(hehe_text, str):
            return []
        return re.findall(r"\b\d{8}[A-Za-z]?\b", hehe_text)

    def distribute_minutes(total_minutes, n_tasks):
        """Even float distribution across tasks."""
        if n_tasks <= 0:
            return []
        per = float(total_minutes) / float(n_tasks)
        return [per] * n_tasks

    def format_task_output(task_order, totals):
        return " | ".join(f"{tid}-{totals[tid]:.2f} mins" for tid in task_order)

    def parse_task_minutes_from_cell(cell_value):
        """
        Parse strings like 'count = 3 | 12345678-11.25 mins | 23456789-11.25 mins' into {taskid: minutes}
        Returns numeric minutes (float).
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

    # ---------- Read input and prepare Sheet1 ----------

    df = pd.read_excel(input_path, engine="openpyxl", dtype={"Dates": str})

    # Insert Minutes immediately after Hours (float, keep fractions)
    if "Hours" not in df.columns:
        raise KeyError("Expected column 'Hours' not found in the input file.")
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

    # Process consecutive rows by (CID, Name, Dates) and fill EQ / PF as cumulative per-day (per earlier logic)
    n = len(df)
    i = 0
    while i < n:
        start = i
        j = i + 1
        while j < n and all(df.at[j, col] == df.at[start, col] for col in ("CID", "Name", "Dates")):
            j += 1
        group_idxs = list(range(start, j))

        # Separate accumulators for EQ and PF (do not mix)
        eq_totals = {}
        eq_order = []
        pf_totals = {}
        pf_order = []

        for idx in group_idxs:
            hehe_val = df.at[idx, "HEHE"]
            if not isinstance(hehe_val, str):
                continue
            hehe_upper = hehe_val.upper()
            minutes = float(df.at[idx, "Minutes"]) if pd.notna(df.at[idx, "Minutes"]) else 0.0

            # DS/EQ case
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
                # write cumulative EQ into this row
                df.at[idx, "EQ"] = f"count = {len(eq_order)} | " + format_task_output(eq_order, eq_totals)
                df.at[idx, "Count"] = len(eq_order)

            # PF case
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

            else:
                # neither DS/EQ nor PF - leave blank
                continue

        i = j

    # ---------- Build Sheet2 (side-by-side PF/EQ), using last-row per date and keep numeric minutes/hours ----------

    sheet2_rows = []
    for (cid, name), g in df.groupby(["CID", "Name"], sort=False):
        # keep date order as appeared
        seen_dates = []
        for d in g["Dates"].tolist():
            if d not in seen_dates:
                seen_dates.append(d)

        pf_tasks = {}       # tid -> total minutes across dates (using last-row per date)
        pf_difficulty = {}  # tid -> difficulty (from the last row for that date)
        eq_tasks = {}
        eq_difficulty = {}

        for d in seen_dates:
            rows_date = g[g["Dates"] == d]
            if rows_date.empty:
                continue
            last_row = rows_date.iloc[-1]  # last row for that date (cumulative totals)

            # difficulty value on that last row (applies to all tasks from that row)
            diff_val = last_row.get("Difficulty", "")
            # PF tasks from last_row
            pf_map = parse_task_minutes_from_cell(last_row.get("PF", ""))
            for tid, mins in pf_map.items():
                pf_tasks[tid] = pf_tasks.get(tid, 0.0) + mins
                # store difficulty (if multiple dates with same task id, last seen difficulty will be used)
                pf_difficulty[tid] = diff_val

            # EQ tasks from last_row
            eq_map = parse_task_minutes_from_cell(last_row.get("EQ", ""))
            for tid, mins in eq_map.items():
                eq_tasks[tid] = eq_tasks.get(tid, 0.0) + mins
                eq_difficulty[tid] = diff_val

        # prepare items and align PF & EQ side-by-side
        pf_items = list(pf_tasks.items())
        eq_items = list(eq_tasks.items())
        max_len = max(len(pf_items), len(eq_items))

        for k in range(max_len):
            pf_task, pf_min = (pf_items[k] if k < len(pf_items) else ("", ""))
            eq_task, eq_min = (eq_items[k] if k < len(eq_items) else ("", ""))

            pf_diff = pf_difficulty.get(pf_task, "") if pf_task else ""
            eq_diff = eq_difficulty.get(eq_task, "") if eq_task else ""

            pf_hours = round(pf_min / 60.0, 2) if pf_min != "" and pf_min is not None else ""
            eq_hours = round(eq_min / 60.0, 2) if eq_min != "" and eq_min is not None else ""

            sheet2_rows.append([
                cid, name,
                pf_task, (round(pf_min, 2) if pf_min != "" else ""), pf_hours, pf_diff,
                eq_task, (round(eq_min, 2) if eq_min != "" else ""), eq_hours, eq_diff
            ])

        # totals (numeric minutes + numeric hours)
        pf_total_minutes = sum(pf_tasks.values()) if pf_tasks else 0.0
        eq_total_minutes = sum(eq_tasks.values()) if eq_tasks else 0.0
        pf_total_hours = round(pf_total_minutes / 60.0, 2) if pf_tasks else ""
        eq_total_hours = round(eq_total_minutes / 60.0, 2) if eq_tasks else ""

        sheet2_rows.append([
            cid, name,
            f"Total ({len(pf_tasks)})", (round(pf_total_minutes, 2) if pf_tasks else ""), pf_total_hours, "",
            f"Total ({len(eq_tasks)})", (round(eq_total_minutes, 2) if eq_tasks else ""), eq_total_hours, ""
        ])
        sheet2_rows.append(["", "", "", "", "", "", "", "", "", ""])

    df_sheet2 = pd.DataFrame(sheet2_rows, columns=[
        "CID", "Name",
        "PF Task", "PF Minutes", "PF Hours", "PF Difficulty",
        "EQ Task", "EQ Minutes", "EQ Hours", "EQ Difficulty"
    ])

    # ---------- Build Sheet3 from Sheet2 (aggregate EQ by Difficulty) ----------
    # For each employee (CID+Name) aggregate EQ tasks (use EQ Hours numeric column).
    sheet3_rows = []
    # top label row will be written separately (we'll write as first row into Excel)
    headers = ["Name", "Difficulty", "Total Tasks", "Total Hrs", "Avg Hrs/Task"]

    # difficulties to always show in this order (include only these; if other labels exist, they will be appended after)
    base_difficulty_order = ["Low", "Medium", "High"]

    for (cid, name), g in df_sheet2.groupby(["CID", "Name"], sort=False):
        # gather EQ task rows for this employee (exclude rows where EQ Task is empty or starts with "Total")
        eq_rows = g[g["EQ Task"].notna() & (g["EQ Task"].astype(str) != "") & (~g["EQ Task"].astype(str).str.startswith("Total"))]

        # Build counts & hours per difficulty (using EQ Difficulty and EQ Hours)
        counts = {}
        hours_sum = {}
        # gather unique difficulties encountered to keep order
        extra_difficulties = []
        for _, r in eq_rows.iterrows():
            diff = r.get("EQ Difficulty", "")
            # prefer string
            diff = diff if isinstance(diff, str) else str(diff)
            try:
                hrs = float(r["EQ Hours"]) if (r["EQ Hours"] not in ("", None)) else 0.0
            except:
                # try parse from EQ Minutes if EQ Hours missing
                try:
                    mins = float(r["EQ Minutes"]) if (r["EQ Minutes"] not in ("", None)) else 0.0
                    hrs = mins / 60.0
                except:
                    hrs = 0.0
            counts[diff] = counts.get(diff, 0) + 1
            hours_sum[diff] = hours_sum.get(diff, 0.0) + hrs
            if diff not in base_difficulty_order and diff not in extra_difficulties:
                extra_difficulties.append(diff)

        # Ensure base order appears (even if zeros)
        difficulty_order = base_difficulty_order + extra_difficulties

        # prepare rows for this employee: print name only once
        first_row = True
        any_equ_data = False
        for diff in difficulty_order:
            task_count = counts.get(diff, 0)
            total_hrs = round(hours_sum.get(diff, 0.0), 2) if task_count else 0.0
            avg_hrs = round((total_hrs / task_count), 2) if task_count else 0.0
            # Only append rows if there is data or we want to show zeros for Low/Medium/High
            # We will show Low/Medium/High rows always (as the user wanted), even if count==0
            sheet3_rows.append([
                name if first_row else "",
                diff,
                task_count,
                total_hrs,
                avg_hrs
            ])
            first_row = False
            any_equ_data = any_equ_data or (task_count > 0)
        # add a blank separator row (optional)
        sheet3_rows.append(["", "", "", "", ""])

    # Build DataFrame for Sheet3
    df_sheet3 = pd.DataFrame(sheet3_rows, columns=headers)

    # ---------- Save all three sheets ----------
    out_path = Path(input_path).with_name(f"processed_{Path(input_path).name}")
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Sheet1", index=False)
        df_sheet2.to_excel(writer, sheet_name="Sheet2", index=False)
        # write label row "EQ" then the header+data for sheet3
        pd.DataFrame([["EQ"]]).to_excel(writer, sheet_name="Sheet3", index=False, header=False)
        df_sheet3.to_excel(writer, sheet_name="Sheet3", startrow=1, index=False)

    # Optional: style Sheet3 top label
    wb = load_workbook(out_path)
    ws3 = wb["Sheet3"]
    try:
        ws3.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
        cell = ws3["A1"]
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.font = Font(bold=True)
    except Exception:
        pass
    wb.save(out_path)

    return out_path
