def process_file(path, save_overwrite=True, verbose=True):
    wb = load_workbook(path)

    if "EPL Data" not in wb.sheetnames or "Report Data" not in wb.sheetnames:
        raise RuntimeError("Workbook must contain both 'EPL Data' and 'Report Data' sheets.")

    ws_EPL = wb["EPL Data"]
    ws_report = wb["Report Data"]

    # ==================================================
    # 1. Read EPL parents (AS-IS)
    # ==================================================
    EPL_parents = read_EPL_parents(ws_EPL)
    parent_set = set(canonical_parent_key(p) for p in EPL_parents if p)

    if verbose:
        print("\n[EPL PARENTS]")
        for p in parent_set:
            print(" ", p)

    # ==================================================
    # 2. Build EPL children qty (AS-IS)
    # ==================================================
    EPL_children_qty = {}
    current_parent = None

    for row in ws_EPL.iter_rows(min_row=2):
        v = row[EPL_Part_COL - 1].value
        qty_val = row[EPL_QTY_COL - 1].value
        desc_val = row[EPL_DESC_COL - 1].value

        qty_empty = qty_val is None or str(qty_val).strip() == ""
        desc_empty = desc_val is None or str(desc_val).strip() == ""

        if v and qty_empty and desc_empty:
            current_parent = canonical_parent_key(v)
            continue

        if not current_parent:
            continue

        child_key = str(v).strip().upper() if v else ""
        if not child_key or child_key.startswith("CA"):
            continue

        EPL_children_qty.setdefault(current_parent, []).append(
            (child_key, "" if qty_val is None else str(qty_val).strip())
        )

    # ==================================================
    # 3. Build report index
    # ==================================================
    report_rows = build_report_index(ws_report)

    # ==================================================
    # 4. Build parent-child map (CORRECTED)
    # ==================================================
    parent_children_map = build_parent_children_map(
        report_rows, ws_report, parent_set, verbose=verbose
    )

    # ==================================================
    # 5. Match against EPL (UNCHANGED LOGIC)
    # ==================================================
    GREEN = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
    RED   = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

    def _to_int_str(val):
        if val is None:
            return None
        try:
            return str(int(float(str(val).strip())))
        except:
            return None

    for parent, children in parent_children_map.items():
        epl_list = list(EPL_children_qty.get(parent, []))

        if verbose:
            print(f"\n[MATCH] Parent={parent} EPL_children={epl_list}")

        for occ in children:
            occ_fk = occ["fk"]
            occ_upa_row = occ["upa_row"]
            occ_upa_val = occ["upa"]

            matched = None

            # direct match
            for i, (child_key, child_qty) in enumerate(epl_list):
                if normalize_spaces(occ_fk).upper() == child_key:
                    matched = (i, child_key, child_qty)
                    break

            # TP fallback
            if matched is None:
                for tp_row in occ["tp_rows"]:
                    alt = ws_report.cell(row=tp_row, column=12).value
                    if alt:
                        alt_u = normalize_spaces(str(alt)).upper()
                        for i, (child_key, child_qty) in enumerate(epl_list):
                            if alt_u == child_key:
                                matched = (i, child_key, child_qty)
                                break
                    if matched:
                        break

            # LP fallback
            if matched is None:
                for lp_row in occ["lp_rows"]:
                    alt = ws_report.cell(row=lp_row, column=12).value
                    if alt:
                        alt_u = normalize_spaces(str(alt)).upper()
                        for i, (child_key, child_qty) in enumerate(epl_list):
                            if alt_u == child_key:
                                matched = (i, child_key, child_qty)
                                break
                    if matched:
                        break

            target_cell = ws_report.cell(row=occ_upa_row, column=REPORT_UPA_COL)

            if matched:
                _, _, child_qty = matched
                if _to_int_str(occ_upa_val) == _to_int_str(child_qty):
                    target_cell.fill = GREEN
                else:
                    target_cell.fill = RED
            else:
                target_cell.fill = RED

    # ==================================================
    # 6. Save
    # ==================================================
    if save_overwrite:
        wb.save(path)

    if verbose:
        print("\n[DONE] Process completed successfully")



def build_parent_children_map(report_rows, ws_report, parent_set, verbose=True):
    """
    Build parent -> children mapping using:
    - IN column for depth
    - ONLY EPL parents as valid parents
    - AIR never becomes parent
    """

    parent_children_map = {}

    # Track last valid EPL parent at each IN level
    last_parent_at_level = {}

    # Track last child at each level (for TP / LP continuation)
    last_child_at_level = {}

    if verbose:
        print("\n[BUILD PARENT-CHILD MAP — EPL PARENTS ONLY]\n")

    for r in report_rows:
        row_no = r["row"]
        part   = r["part"]
        level  = r["level"]

        if not part or level is None:
            continue

        part_u = canonical_parent_key(part)

        # -------------------------------
        # Find structural parent (EPL only)
        # -------------------------------
        structural_parent = last_parent_at_level.get(level - 1)

        # AIR → GF redirection
        effective_parent = structural_parent
        if part_u.startswith("AIR") and structural_parent:
            effective_parent = "GF" + structural_parent.replace("IR", "", 1)

        # -------------------------------
        # Attach child to parent
        # -------------------------------
        child_entry = {
            "fk": part_u,
            "upa": r["upa"],
            "upa_row": row_no,
            "tp_rows": [],
            "lp_rows": [],
            "used": False
        }

        if effective_parent:
            parent_children_map.setdefault(effective_parent, []).append(child_entry)

        # -------------------------------
        # TP / LP alt-pn rows (unchanged)
        # -------------------------------
        dsi_val = ws_report.cell(row=row_no, column=13).value  # DSI column
        if dsi_val:
            dsi = str(dsi_val).strip().upper()
            prev_child = last_child_at_level.get(level)
            if prev_child:
                if dsi == "TP":
                    prev_child["tp_rows"].append(row_no)
                elif dsi == "LP":
                    prev_child["lp_rows"].append(row_no)

        # -------------------------------
        # Update last_parent_at_level
        # ONLY if this part is an EPL parent
        # -------------------------------
        if part_u in parent_set and not part_u.startswith("AIR"):
            last_parent_at_level[level] = part_u

            if verbose:
                print(f"[PARENT] row={row_no} IN={level} PARENT={part_u}")

        last_child_at_level[level] = child_entry

        if verbose:
            print(
                f"[ROW {row_no}] IN={level} PART={part_u} "
                f"STRUCT_PARENT={structural_parent} "
                f"EFFECTIVE_PARENT={effective_parent}"
            )

    return parent_children_map
      
                                
