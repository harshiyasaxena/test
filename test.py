def get_child_PN_for_IW1(sheet2, parent_pn):
    ws = sheet2
    pending = False
    last_pn = None
    in_direct = False
    DI_header = None
    PN_col = None
    RPNP_col = None
    IW_col = None

    mu_list = []
    mn_list = []

    print(f"\n===== DEBUG: Finding IW=1 children for parent: {parent_pn} =====")

    for row in ws.iter_rows(min_row=1):
        first = str(row[0].value).strip() if row[0].value else None

        if first == "PN":
            pending = True
            continue

        if pending:
            last_pn = first
            pending = False
            print(f"[DEBUG] Found parent PN block: {last_pn}")
            continue

        if last_pn != parent_pn:
            continue

        if first and first.lower() == "direct/indirect":
            in_direct = True
            DI_header = row
            print("[DEBUG] Entered Direct/Indirect section")
            continue

        if in_direct:
            # Detect column indexes
            if PN_col is None or IW_col is None:
                for idx, c in enumerate(DI_header):
                    val = str(c.value).strip() if c.value else None
                    if val == "PN":
                        PN_col = idx
                    if val == "IW":
                        IW_col = idx
                    if val == "RP/NP PN":
                        RPNP_col = idx

                print(f"[DEBUG] PN_col={PN_col}, IW_col={IW_col}, RPNP_col={RPNP_col}")

            # Stop when next PN block starts
            if first == "PN":
                print("[DEBUG] Next PN block reached, stopping IW scan.")
                break

            iw_cell = row[IW_col].value if IW_col is not None else None
            try:
                iw_val = int(float(str(iw_cell).strip()))
            except:
                iw_val = None

            print(f"[DEBUG] Row IW value: {iw_val}")

            if iw_val == 1:
                mu_child = row[PN_col].value if PN_col is not None else None
                mn_child = row[RPNP_col].value if RPNP_col is not None else None

                print(f"[DEBUG] Found IW=1 → MU={mu_child}, MN={mn_child}")

                if mu_child:
                    mu_list.append(str(mu_child).strip())
                if mn_child:
                    mn_list.append(str(mn_child).strip())

    print(f"[RESULT] mu_list={mu_list}, mn_list={mn_list}\n")
    return mu_list, mn_list

if dsi_val and dsi_val.upper() in ("MU", "MN"):
    
    # debug print
    print(f"\n[DEBUG-S3] Checking PN={part_no4}, DSI={dsi_val}, Nomen(L)={row[11].value}")

    mu_list, mn_list = get_child_PN_for_IW1(self.sheet2, part_no4)

    # debug
    print(f"[DEBUG-S3] MU list = {mu_list}")
    print(f"[DEBUG-S3] MN list = {mn_list}")

    extracted_L = extract_last_token(row[11].value)
    print(f"[DEBUG-S3] Extracted token from column L = {extracted_L}")

    # ---------------------- MU comparison -----------------------
    if dsi_val.upper() == "MU":
        if normalize_PN(extracted_L) not in {normalize_PN(x) for x in mu_list}:
            for cell in row:
                cell.fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
            add_error_comment_to_PN_cell(self.sheet3, part_no4, "MU part no mismatch")
            print("[DEBUG-S3] MU mismatch -> RED")
            continue
        else:
            print("[DEBUG-S3] MU matched -> OK (no color change)")

    # ---------------------- MN comparison -----------------------
    if dsi_val.upper() == "MN":
        if normalize_PN(extracted_L) not in {normalize_PN(x) for x in mn_list}:
            for cell in row:
                cell.fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
            add_error_comment_to_PN_cell(self.sheet3, part_no4, "MN part no mismatch")
            print("[DEBUG-S3] MN mismatch -> RED")
            continue
        else:
            print("[DEBUG-S3] MN matched -> OK (no color change)")
    
    
