for row in self.sheet3.iter_rows(min_row=2):  # assuming header in row 1
    part_no3 = str(row[0].value).strip() if row[0].value else None
    dsi3 = str(row[10].value).strip() if row[10].value else None  # DSI column
    nomen3 = str(row[11].value).strip() if row[11].value else None  # description column

    if not part_no3 or not dsi3 or dsi3.upper() != "QD":
        continue

    # ---- Check if this PN exists in Sheet2 mapping ----
    if part_no3 not in PN_to_desc3:
        continue

    desc_list = PN_to_desc3[part_no3]
    if not isinstance(desc_list, list):
        desc_list = [desc_list]

    # ---- Get all QD rows in Sheet3 for this PN ----
    qd_rows_for_pn = [
        r for r in self.sheet3.iter_rows(min_row=2)
        if str(r[0].value).strip() == part_no3 and str(r[10].value).strip().upper() == "QD"
    ]

    # ---- Strict count match check ----
    if len(qd_rows_for_pn) != len(desc_list):
        # Mark the PN cell red
        pn_cell = row[0]
        pn_cell.fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
        # Add an error comment
        add_error_comment_to_PN_cell(
            self.sheet3,
            part_no3,
            text=f"QD count mismatch: Sheet2 has {len(desc_list)} desc, Sheet3 has {len(qd_rows_for_pn)} QD rows"
        )
        # Skip matching for this PN altogether
        continue

    # ---- If counts match, compare each QD vs IW description ----
    for qd_row, desc_text in zip(qd_rows_for_pn, desc_list):
        nomen3 = str(qd_row[11].value).strip() if qd_row[11].value else ""
        if normalize_PN(nomen3) == normalize_PN(desc_text):
            qd_row[11].fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
            qd_row[12].fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
        else:
            qd_row[11].fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
            qd_row[12].fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
            add_error_comment_to_PN_cell(
                self.sheet3,
                part_no3,
                text=f"QD description mismatch for PN {part_no3}: expected '{desc_text}'"
            )
