def build_MD_desc_from_steps(sheet2):
    """
    Returns:
       MD_map = {
           "12345678": ["desc1", "desc2", ...],
           "98765432": ["desc1"],
           ...
       }
    """
    MD_map = {}
    last_pn = None
    pending_pn = False

    for row in sheet2.iter_rows(min_row=1):
        first = str(row[0].value).strip() if row[0].value else None

        if not first:
            continue

        # Detect new PN start
        if first.upper() == "PN":
            pending_pn = True
            continue

        if pending_pn:
            last_pn = first
            pending_pn = False
            continue

        # Collect MD lines for this PN
        if first.upper() == "MD" and last_pn:
            desc_parts = []
            for cell in row[3:12]:  # columns 4-12 contain MD description
                if cell.value:
                    desc_parts.append(str(cell.value).strip())
            desc = " ".join(desc_parts).strip()

            MD_map.setdefault(last_pn, []).append(desc)

    return MD_map


def build_MD_blocks_from_fcr(sheet3):
    """
    Returns:
      [
        {"pn": "12345678", "md_desc": ["a","b"], "rows": [2,3,4]},
        {"pn": "23456789", "md_desc": ["x"],     "rows": [5,6]},
        ...
      ]
    """
    blocks = []
    current_pn = None
    current_rows = []
    current_md_desc = []

    for row_cells in sheet3.iter_rows(min_row=2):
        pn_val = row_cells[9].value
        dsi_val = row_cells[12].value
        nomen_val = row_cells[11].value

        # New block starts
        if pn_val:
            # store previous block
            if current_pn is not None:
                blocks.append({
                    "pn": current_pn,
                    "md_desc": current_md_desc,
                    "rows": current_rows
                })

            current_pn = str(pn_val).strip()
            current_rows = []
            current_md_desc = []

        # add row to block
        if current_pn:
            current_rows.append(row_cells[0].row)

            # collect MD rows
            if dsi_val and str(dsi_val).strip().upper() == "MD":
                if nomen_val:
                    current_md_desc.append(str(nomen_val).strip())

    # append last block
    if current_pn:
        blocks.append({
            "pn": current_pn,
            "md_desc": current_md_desc,
            "rows": current_rows
        })

    return blocks


def validate_MD_per_block(sheet3, MD_map_steps):
    red = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    blocks = build_MD_blocks_from_fcr(sheet3)

    for blk in blocks:
        pn = blk["pn"]
        fcr_md_list = blk["md_desc"]
        parent_row = blk["rows"][0]

        steps_md_list = MD_map_steps.get(pn, [])

        steps_count = len(steps_md_list)
        fcr_count = len(fcr_md_list)

        # CASE 1: FCR has MD but steps has NO MD
        if fcr_count > 0 and steps_count == 0:
            sheet3.cell(row=parent_row, column=10).fill = red
            add_error_comment_to_cell(sheet3, parent_row, 10, "MD present in FCR but not in STEP DATA")
            continue

        # CASE 2: Steps has MD but FCR has none
        if steps_count > 0 and fcr_count == 0:
            sheet3.cell(row=parent_row, column=10).fill = red
            add_error_comment_to_cell(sheet3, parent_row, 10, "Missing MD in FCR")
            continue

        # CASE 3: Counts mismatch
        if steps_count != fcr_count:
            sheet3.cell(row=parent_row, column=10).fill = red
            add_error_comment_to_cell(
                sheet3, parent_row, 10,
                f"MD count mismatch: expected {steps_count}, got {fcr_count}"
            )
            continue

        # CASE 4: Compare descriptions (unordered)
        # Use sets but consider duplicates -> use multisets (Counter)
        from collections import Counter
        if Counter([d.lower() for d in steps_md_list]) != Counter([d.lower() for d in fcr_md_list]):
            sheet3.cell(row=parent_row, column=10).fill = red
            add_error_comment_to_cell(sheet3, parent_row, 10, "MD description mismatch")
            continue

