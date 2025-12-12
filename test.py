def validate_PF_block(sheet3, blk, PF_desc_steps):
    global PF_duplicate_blocks

    red = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    green = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")

    pn = blk["pn"]
    rows = blk["rows"]
    parent_row = rows[0]

    # ⭐ Skip entire block if parent row is deleted
    parent_t_val = sheet3.cell(row=parent_row, column=4).value
    if parent_t_val and str(parent_t_val).strip().upper() == "D":
        return

    # ----------------------------------------------------
    # ⭐ We make 2 lists:
    # PF_rows_raw  -> ALL PF rows (including deleted)
    # PF_rows      -> only non-deleted PF rows (for matching)
    # ----------------------------------------------------
    PF_rows_raw = []
    PF_rows = []

    for r in rows:
        dsi_val = sheet3.cell(row=r, column=13).value
        if dsi_val and str(dsi_val).strip().upper() == "PF":
            PF_rows_raw.append(r)  # include deleted PF here

            # skip if deleted
            t_val = sheet3.cell(row=r, column=4).value
            if t_val and str(t_val).strip().upper() == "D":
                continue

            PF_rows.append(r)  # only non-deleted rows here

    pf_raw_count = len(PF_rows_raw)    # used for "PF exists in report"
    pf_count = len(PF_rows)            # used for matching only
    pf_in_steps = PF_desc_steps is not None

    steps_norm = normalize(PF_desc_steps or "")

    # ----------------------------------------------------
    # ⭐ CASE 1 — PF present in FCR (even deleted) but not in STEP DATA
    # ----------------------------------------------------
    if pf_raw_count > 0 and not pf_in_steps:
        sheet3.cell(row=parent_row, column=10).fill = red
        add_error_comment_to_cell(sheet3, parent_row, 10,
                                  "PF found in FCR but not in STEP DATA")

        # only mark non-deleted PF rows red
        for r in PF_rows:
            sheet3.cell(row=r, column=12).fill = red
            sheet3.cell(row=r, column=13).fill = red
        return

    # ----------------------------------------------------
    # ⭐ CASE 2 — PF expected but missing (no PF rows at all except deleted)
    # ----------------------------------------------------
    if pf_in_steps and pf_raw_count == 0:
        sheet3.cell(row=parent_row, column=10).fill = red
        add_error_comment_to_cell(sheet3, parent_row, 10,
                                  "Missing PF in this occurrence")
        return

    # ----------------------------------------------------
    # ⭐ CASE 3 — Multiple PF rows (ONLY non-deleted count)
    # ----------------------------------------------------
    if pf_count > 1:
        sheet3.cell(row=parent_row, column=10).fill = red
        add_error_comment_to_cell(sheet3, parent_row, 10,
                                  "Multiple PF rows found in this occurrence")

        PF_duplicate_blocks.add((pn, parent_row))

        matched_found = False

        for r in PF_rows:  # only non-deleted rows
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

    # (Single PF case handled by compare_and_color)
    return
