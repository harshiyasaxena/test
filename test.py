def check_in_direct_indirect_PN(sheet2, target_pn):
    """Check if the given PN appears under 'PN' column in 'Direct/Indirect' section."""
    found_PN_section = False
    current_PN_header = None
    for row in sheet2.iter_rows(min_row=1):
        a_val = str(row[0].value).strip() if row[0].value else None
        if not a_val:
            continue

        if a_val.upper() == "PN":
            found_PN_section = True
            continue

        if found_PN_section:
            # when we encounter the PN value, start tracking
            current_PN_header = a_val
            found_PN_section = False
            continue

        # detect Direct/Indirect section
        if a_val and a_val.strip().upper() == "DIRECT/ INDIRECT":
            # scan downward for a PN column in this section
            for cell in row:
                if str(cell.value).strip().upper() == "PN":
                    pn_col = cell.column
                    v_row = cell.row + 1
                    ws = cell.parent
                    while v_row <= ws.max_row:
                        first_col_val = str(ws.cell(row=v_row, column=1).value).strip() if ws.cell(row=v_row, column=1).value else ""
                        if first_col_val and first_col_val.upper() == "PN":
                            break
                        val = ws.cell(row=v_row, column=pn_col).value
                        if val and normalize_PN(str(val)) == normalize_PN(target_pn):
                            return True
                        v_row += 1
    return False
