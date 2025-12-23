def canonical_parent_key(p):
    if not p or not isinstance(p, str):
        return None
    return p.strip().upper()

def build_report_index(ws_report):
    """
    Builds a flat index but ALSO keeps IN (column I) for hierarchy building.
    """
    rows = []
    for idx, row in enumerate(ws_report.iter_rows(min_row=2), start=2):
        fk = row[REPORT_Part_COL - 1].value      # J
        UPA = row[REPORT_UPA_COL - 1].value      # N
        IN_val = row[8].value                    # I

        fk_u = str(fk).strip().upper() if fk and str(fk).strip() else None
        UPA_s = "" if UPA is None else str(UPA).strip()

        try:
            level = int(IN_val)
        except:
            level = None

        rows.append({
            "row": idx,
            "part": fk_u,
            "upa": UPA_s,
            "level": level
        })
    return rows


def collect_children_for_parent(report_rows, parent_token, ws_report):
    """
    Collects children based on IN hierarchy.
    AIR* parts are redirected to GF<parent>.
    """
    children = []

    # find parent row
    parent_idx = None
    parent_level = None

    for r in report_rows:
        if r["part"] == parent_token:
            parent_idx = r["row"]
            parent_level = r["level"]
            break

    if parent_idx is None or parent_level is None:
        return [], None

    current_child = None

    for r in report_rows:
        if r["row"] <= parent_idx:
            continue

        if r["level"] is None:
            continue

        # stop when hierarchy goes back or equal
        if r["level"] <= parent_level:
            break

        part = r["part"]
        if not part:
            continue

        # AIR → GF redirection
        effective_parent = parent_token
        if part.startswith("AIR"):
            base = parent_token.replace("IR", "", 1)
            effective_parent = "GF" + base

        entry = {
            "fk": part,
            "upa": r["upa"],
            "upa_row": r["row"],
            "tp_rows": [],
            "lp_rows": [],
            "used": False
        }

        children.append(entry)
        current_child = entry

        # TP / LP rows belong to last child
        dsi_val = ws_report.cell(row=r["row"], column=13).value
        if dsi_val and current_child:
            dsi = str(dsi_val).strip().upper()
            if dsi == "TP":
                current_child["tp_rows"].append(r["row"])
            elif dsi == "LP":
                current_child["lp_rows"].append(r["row"])

    return children, parent_idx

def process_file(path, save_overwrite=True, verbose=True):
    wb = load_workbook(path)
    if "EPL Data" not in wb.sheetnames or "Report Data" not in wb.sheetnames:
        raise RuntimeError("Workbook must contain both 'EPL Data' and 'Report Data' sheets.")

    ws_EPL = wb["EPL Data"]
    ws_report = wb["Report Data"]

    EPL_parents = read_EPL_parents(ws_EPL)
    parent_set = set(EPL_parents)

    EPL_children_qty = {}
    current_parent = None

    for row in ws_EPL.iter_rows(min_row=2):
        v = row[EPL_Part_COL - 1].value
        qty_val = row[EPL_QTY_COL - 1].value
        qty_empty = qty_val is None or str(qty_val).strip() == ""
        desc_empty = row[EPL_DESC_COL - 1].value in (None, "")

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

    canonice_parents = set(EPL_children_qty.keys())

    # --- NEW: hierarchy-aware report index ---
    report_rows = build_report_index(ws_report)

    parent_children_map = {}
    parent_row_index_map = {}

    for p in canonice_parents:
        children, parent_ws_row = collect_children_for_parent(report_rows, p, ws_report)
        parent_children_map[p] = children
        parent_row_index_map[p] = parent_ws_row

    GREEN = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
    RED = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

    for parent_token, occurrences in parent_children_map.items():
        epl_list = list(EPL_children_qty.get(parent_token, []))

        for occ in occurrences:
            occ_fk = occ["fk"]
            occ_upa_row = occ["upa_row"]
            occ_upa_val = occ["upa"]

            matched = None
            for ei, (child_key, child_qty) in enumerate(epl_list):
                if normalize_spaces(occ_fk) == child_key:
                    matched = (ei, child_qty)
                    break

            if matched:
                ei, child_qty = matched
                UPA_int = _to_int_str(occ_upa_val)
                qty_int = _to_int_str(child_qty)

                cell = ws_report.cell(row=occ_upa_row, column=REPORT_UPA_COL)
                cell.fill = GREEN if UPA_int == qty_int else RED

                epl_list.pop(ei)

    if save_overwrite:
        wb.save(path)
