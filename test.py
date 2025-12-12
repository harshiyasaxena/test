def build_PN_blocks_with_dsi(sheet3):

    blocks = []
    current_pn = None
    current_block_rows = []
    current_dsis = []
    current_block_deleted = False   # will skip whole block if parent row deleted

    for row_cells in sheet3.iter_rows(min_row=2):

        pn_val = row_cells[9].value
        dsi_val = row_cells[12].value
        t_val = row_cells[3].value   # T_VAL is column 4

        # -------------------------------------------------------
        # NEW BLOCK STARTS 
        # -------------------------------------------------------
        if pn_val:

            # save previous block if not deleted
            if current_pn is not None and not current_block_deleted:
                blocks.append({
                    "pn": current_pn,
                    "dsis": current_dsis,
                    "rows": current_block_rows
                })

            # start new block
            current_pn = str(pn_val).strip()
            current_block_rows = []
            current_dsis = []

            # check if parent row of PN is deleted
            current_block_deleted = (
                t_val and str(t_val).strip().upper() == "D"
            )

        # -------------------------------------------------------
        # Add row to block
        # -------------------------------------------------------
        if current_pn:

            current_block_rows.append(row_cells[0].row)

            # ⭐ ALWAYS RECORD DSI PRESENCE (even if deleted)
            # So we know PF/MD exists in report
            if dsi_val:
                current_dsis.append(str(dsi_val).strip())

            # NOTE:
            # Deleted DSIs will be ignored later in validator when comparing

    # -------------------------------------------------------
    # Append final block (if not deleted)
    # -------------------------------------------------------
    if current_pn and not current_block_deleted:
        blocks.append({
            "pn": current_pn,
            "dsis": current_dsis,
            "rows": current_block_rows
        })

    return blocks


def validate_PF_block(sheet3, blk, PF_desc_steps):
    global PF_duplicate_blocks

    red = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    green = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")

    pn = blk["pn"]
    rows = blk["rows"]
    parent_row = rows[0]

    # ----------------------------------------
    # Skip entire block if parent row is deleted
    # ----------------------------------------
    parent_t_val = sheet3.cell(row=parent_row, column=4).value
    if parent_t_val and str(parent_t_val).strip().upper() == "D":
        return

    # ----------------------------------------
    # Collect PF rows
    # PF_raw = all PF rows (even deleted)
    # PF_active = only non-deleted PF rows
    # ----------------------------------------
    PF_raw = []
    PF_active = []

    for r in rows:
        dsi_val = sheet3.cell(row=r, column=13).value
        if not dsi_val or str(dsi_val).strip().upper() != "PF":
            continue

        PF_raw.append(r)  # presence detection

        t_val = sheet3.cell(row=r, column=4).value
        if t_val and str(t_val).strip().upper() == "D":
            continue  # skip deleted PF rows

        PF_active.append(r)

    pf_raw_count = len(PF_raw)
    pf_active_count = len(PF_active)
    pf_in_steps = PF_desc_steps is not None

    steps_norm = normalize(PF_desc_steps or "")

    # ----------------------------------------
    # CASE 1: PF exists in FCR but NOT in STEPS
    # ----------------------------------------
    if pf_raw_count > 0 and not pf_in_steps:
        sheet3.cell(row=parent_row, column=10).fill = red
        add_error_comment_to_cell(sheet3, parent_row, 10,
                                  "PF found in FCR but not in STEP DATA")

        # highlight only active PF rows (not deleted)
        for r in PF_active:
            sheet3.cell(row=r, column=12).fill = red
            sheet3.cell(row=r, column=13).fill = red

        return

    # ----------------------------------------
    # CASE 2: PF expected in steps but NOT found in FCR
    # deleted PF rows do NOT count
    # ----------------------------------------
    if pf_in_steps and pf_raw_count == 0:
        sheet3.cell(row=parent_row, column=10).fill = red
        add_error_comment_to_cell(sheet3, parent_row, 10,
                                  "Missing PF in this occurrence")
        return

    # ----------------------------------------
    # CASE 3: Multiple PF rows (only active PF count matters)
    # ----------------------------------------
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

    # single PF case → compare handled by compare_and_color
    return
