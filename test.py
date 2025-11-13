def check_in_direct_indirect_PN(sheet2, target_pn):
    """
    Returns True if target_pn appears under the 'Direct/Indirect' → 'PN' column
    within its own PN group section in Sheet2.
    """
    ws = sheet2
    target_norm = normalize_PN(str(target_pn))
    found_group = False

    for row in ws.iter_rows(min_row=1, max_col=ws.max_column):
        first_val = str(row[0].value).strip().upper() if row[0].value else ""

        # --- detect start of group for this PN ---
        if first_val == "PN":
            pn_val = str(row[1].value).strip() if row[1].value else None
            if pn_val and normalize_PN(pn_val) == target_norm:
                found_group = True
            else:
                found_group = False
            continue

        # --- if inside correct PN group ---
        if found_group and first_val in ("PF", "DIRECT/INDIRECT", "DIRECT", "INDIRECT"):
            # locate the PN column header in this section
            for cell in row:
                if str(cell.value).strip().upper() == "PN":
                    pn_col = cell.column
                    r = cell.row + 1
                    while r <= ws.max_row:
                        next_val = ws.cell(row=r, column=1).value
                        if next_val and str(next_val).strip().upper() == "PN":
                            break  # stop when next PN group starts
                        val = ws.cell(row=r, column=pn_col).value
                        if val and normalize_PN(str(val)) == target_norm:
                            return True
                        r += 1
            break  # stop scanning once Direct/Indirect section processed

    return False
