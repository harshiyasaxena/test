def add_error_comment_to_cell(sheet, row_index, col_index, text, author="System"):
    """
    Attach `text` as a comment to sheet.cell(row_index, col_index).
    Does not duplicate identical lines if already present.
    """
    cell = sheet.cell(row=row_index, column=col_index)
    existing = cell.comment.text if cell.comment else None
    if existing:
        # avoid adding duplicate message lines
        existing_lines = [ln.strip() for ln in existing.splitlines() if ln.strip()]
        if text.strip() in existing_lines:
            return
        new_text = existing + "\n" + text
    else:
        new_text = text

    width_pt, height_pt = _comment_size_for_text(new_text)
    try:
        cell.comment = Comment(new_text, author=author, width=width_pt, height=height_pt)
    except TypeError:
        # older openpyxl may not accept width/height in constructor
        c = Comment(new_text, author=author)
        try:
            c.width = width_pt
            c.height = height_pt
        except Exception:
            pass
        cell.comment = c


def mark_missing_PF_per_block(sheet3, PN_to_desc):
    red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    blocks = build_PN_blocks_with_dsi(sheet3)

    for blk in blocks:
        pn = blk["pn"]
        dsis = blk["dsis"]

        # PF expected for this PN?
        if pn not in PN_to_desc:
            continue

        # PF missing in this block?
        if "PF" not in [d.upper() for d in dsis]:
            # mark only the FIRST ROW (the parent row) of this block
            parent_row_index = blk["rows"][0]
            sheet3.cell(row=parent_row_index, column=10).fill = red_fill  # column 10 = Part No (J)
            add_error_comment_to_cell(sheet3, parent_row_index, 10, "Missing PF in this occurrence")


def compare_and_color(sheet3, PN_to_desc, dsi_match,
                      error_msg_unmatched, normalize_fn):
    last_part_no = None
    parent_row_index = None
    parent_t_val = None

    for row in sheet3.iter_rows(min_row=2):
        if row[9].value:
            # new parent found -> update both PN and its row index
            last_part_no = normalize_fn(row[9].value)
            parent_row_index = row[9].row
            parent_t_val = str(row[3].value).strip() if row[3].value else None

        part_no = last_part_no
        nomen = str(row[11].value).strip() if row[11].value else None
        dsi = str(row[12].value).strip() if row[12].value else None

        if parent_t_val == "D":
            continue

        if dsi and dsi.lower() == dsi_match.lower() and part_no in PN_to_desc:
            desc = PN_to_desc[part_no]

            if dsi_match.upper() in ("PF", "MD", "QDQ"):
                left = normalize(nomen) if nomen is not None else ""
                right = normalize(desc) if desc is not None else ""
            else:
                left = normalize_PN(nomen)
                right = normalize_PN(desc)

            if left == right:
                row[11].fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
                row[12].fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
            else:
                row[11].fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
                row[12].fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
                # ADD comment only on this block's parent cell (not globally)
                if parent_row_index:
                    add_error_comment_to_cell(sheet3, parent_row_index, 10, error_msg_unmatched)
