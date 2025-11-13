if dsi3 and dsi3.upper() == "QD" and part_no3 in PN_to_desc3:
    desc_list = PN_to_desc3[part_no3]
    if not isinstance(desc_list, list):
        desc_list = [desc_list]

    # 🔹 NEW: get all QD rows for this PN in sheet3
    qd_rows_for_pn = [
        r for r in self.sheet3.iter_rows(min_row=2)
        if str(r[0].value).strip() == part_no3 and str(r[10].value).strip().upper() == "QD"
    ]

    # 🔹 NEW: check if counts match
    if len(qd_rows_for_pn) != len(desc_list):
        # mark only the PN cell red + add comment
        row[0].fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
        add_error_comment_to_PN_cell(
            self.sheet3,
            part_no3,
            text=f"QD count mismatch: Sheet2 has {len(desc_list)} desc, Sheet3 has {len(qd_rows_for_pn)} QD rows"
        )
        continue  # skip further comparison

    # (existing matching logic below stays same)
    match_found = any(normalize_PN(nomen3) == normalize_PN(desc) for desc in desc_list)

    if match_found:
        row[11].fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
        row[12].fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
    else:
        row[11].fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
        row[12].fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
        add_error_comment_to_PN_cell(self.sheet3, part_no3, text="QD description unmatched")
