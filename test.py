def validate_PF_block(sheet3, blk, PF_desc_steps):
    global PF_duplicate_blocks

    red = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    green = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")

    pn = blk["pn"]
    rows = blk["rows"]
    parent_row = rows[0]

    # ⭐ SKIP entire block if parent PN row is deleted
    parent_t = sheet3.cell(row=parent_row, column=4).value
    if parent_t and str(parent_t).strip().upper() == "D":
        return

    # ⭐ We build TWO lists:
    PF_raw = []     # detects PF presence (includes deleted)
    PF_active = []  # used for comparison (excludes deleted)

    for r in rows:
        dsi = sheet3.cell(row=r, column=13).value
        if dsi and str(dsi).strip().upper() == "PF":
            PF_raw.append(r)   # ALWAYS append here (even if deleted)

            # check t_val
            t_val = sheet3.cell(row=r, column=4).value
            if t_val and str(t_val).strip().upper() == "D":
                continue      # skip deleted PF from active list

            PF_active.append(r)

    pf_exists_in_fcr = len(PF_raw) > 0
    pf_active_count = len(PF_active)
    pf_exists_in_steps = PF_desc_steps is not None

    steps_norm = normalize(PF_desc_steps or "")

    # ⭐ CASE 1 — PF exists in FCR but not in STEPS
    if pf_exists_in_fcr and not pf_exists_in_steps:
        sheet3.cell(row=parent_row, column=10).fill = red
        add_error_comment_to_cell(sheet3, parent_row, 10,
                                  "PF found in FCR but not in STEP DATA")

        # only mark ACTIVE (non-deleted) rows
        for r in PF_active:
            sheet3.cell(row=r, column=12).fill = red
            sheet3.cell(row=r, column=13).fill = red
        return

    # ⭐ CASE 2 — PF expected but not in FCR (raw PF list empty)
    if pf_exists_in_steps and not pf_exists_in_fcr:
        sheet3.cell(row=parent_row, column=10).fill = red
        add_error_comment_to_cell(sheet3, parent_row, 10,
                                  "Missing PF in this occurrence")
        return

    # ⭐ CASE 3 — Multiple PF rows (only ACTIVE rows count)
    if pf_active_count > 1:
        sheet3.cell(row=parent_row, column=10).fill = red
        add_error_comment_to_cell(sheet3, parent_row, 10,
                                  "Multiple PF rows found in this occurrence")
        PF_duplicate_blocks.add((pn, parent_row))

        matched_found = False
        for r in PF_active:
            nomen = sheet3.cell(row=r, column=12).value
            nomen_norm = normalize(nomen or "")

            if not matched_found and nomen_norm == steps_norm:
                sheet3.cell(row=r, column=12).fill = green
                sheet3.cell(row=r, column=13).fill = green
                matched_found = True
            else:
                sheet3.cell(row=r, column=12).fill = red
                sheet3.cell(row=r, column=13).fill = red

        return

    # Single PF case handled in compare_and_color
    return
