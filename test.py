def get_child_PN_for_IW1(sheet2, parent_pn):
    """
    For a given parent_pn (from sheet3 col J), go to sheet2 and:
      - find the PN block whose PN matches parent_pn (using normalize_PN)
      - inside its 'Direct/Indirect' section, find the row containing header 'IW'
      - then scan all rows below it, until next 'PN' block starts
      - for every row where IW == 1, collect:
            MU candidates  -> value under 'PN'
            MN candidates  -> value under 'RP/NP PN'
    Return (mu_list, mn_list)
    """
    ws = sheet2

    def norm_pn(x):
        if x is None:
            return None
        s = str(x).strip()
        if s.upper().startswith("IR"):
            s = s[2:].strip()
        return s.upper()

    pending = False
    last_pn = None
    in_direct = False
    DI_header = None
    PN_col = None
    RPNP_col = None
    IW_col = None

    mu_list = []
    mn_list = []

    for row in ws.iter_rows(min_row=1):
        first = str(row[0].value).strip() if row[0].value else None

        # --- Detect new PN block header in column A ---
        if first and first.upper() == "PN":
            # if we were already inside desired parent block and have collected something, we can stop
            if last_pn and norm_pn(last_pn) == norm_pn(parent_pn) and (mu_list or mn_list):
                break

            pending = True
            in_direct = False     # reset for next block
            DI_header = None
            PN_col = RPNP_col = IW_col = None
            continue

        # --- The row immediately after "PN" is the part number itself ---
        if pending:
            last_pn = first
            pending = False
            continue

        # If we are not in the block of the desired parent_pn, skip
        if not last_pn or norm_pn(last_pn) != norm_pn(parent_pn):
            continue

        # --- Enter Direct/Indirect header row ---
        if first and first.lower() == "direct/indirect":
            in_direct = True
            DI_header = row
            PN_col = RPNP_col = IW_col = None
            continue

        # --- Inside Direct/Indirect section ---
        if in_direct:
            # If we encounter a new PN block while in Direct/Indirect, we are done
            if first and first.upper() == "PN":
                break

            # Detect the column indices once from DI_header row (0-based indices for 'row[...]')
            if PN_col is None or IW_col is None:
                for idx, c in enumerate(DI_header):
                    val = str(c.value).strip() if c.value else None
                    if val == "PN":
                        PN_col = idx
                    elif val == "IW":
                        IW_col = idx
                    elif val == "RP/NP PN":
                        RPNP_col = idx

            if IW_col is None:
                continue

            # Read IW value in this row
            iw_raw = row[IW_col].value
            try:
                iw_val = int(float(str(iw_raw).strip()))
            except Exception:
                iw_val = None

            # Only care about rows where IW == 1
            if iw_val == 1:
                if PN_col is not None:
                    pn_val = row[PN_col].value
                    if pn_val:
                        mu_list.append(str(pn_val).strip())

                if RPNP_col is not None:
                    mn_val = row[RPNP_col].value
                    if mn_val:
                        mn_list.append(str(mn_val).strip())

    return mu_list, mn_list


        if dsi_val and dsi_val.upper() in ("MU", "MN"):
            # Get all MU/MN candidates (for IW == 1) from sheet2, for this parent PN
            mu_list, mn_list = get_child_PN_for_IW1(self.sheet2, part_no4)

            extracted_L = extract_last_token(row[11].value)
            norm_token = normalize_PN(extracted_L)

            norm_mu_set = {normalize_PN(x) for x in mu_list if x}
            norm_mn_set = {normalize_PN(x) for x in mn_list if x}

            # --- MU: check against any PN with IW=1 ---
            if dsi_val.upper() == "MU":
                if norm_token not in norm_mu_set:
                    for cell in row:
                        cell.fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
                    add_error_comment_to_PN_cell(self.sheet3, part_no4, "MU part no mismatch")
                    continue

            # --- MN: check against any RP/NP PN with IW=1 ---
            if dsi_val.upper() == "MN":
                if norm_token not in norm_mn_set:
                    for cell in row:
                        cell.fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
                    add_error_comment_to_PN_cell(self.sheet3, part_no4, "MN part no mismatch")
                    continue

            # If MU/MN token matched, then apply your IW_val logic as before:
            if IW_val != 1:
                if check_in_direct_indirect_PN(self.sheet2, part_no4):
                    for cell in row:
                        cell.fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
                    continue

                if t_val == "D":
                    for cell in row:
                        cell.fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
                else:
                    PN_row_t = PN_to_sheet3_t.get(part_no4)
                    if PN_row_t == "D":
                        for cell in row:
                            cell.fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
                    else:
                        for cell in row:
                            cell.fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
                        if row[9].value:
                            add_error_comment_to_PN_cell(self.sheet3, part_no4, text="MU/MN rule violation")
            else:
                # IW_val == 1, everything OK
                for cell in row:
                    cell.fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
            
