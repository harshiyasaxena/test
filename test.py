def build_PN_blocks_with_dsi(sheet3):
    """
    Returns a list of blocks:
    [
      {"pn": "12345678", "dsis": ["AB","BC","CD","PF"], "rows": [2,3,4,5]},
      {"pn": "23456789", "dsis": ["AB","BC"],           "rows": [6,7]},
      {"pn": "12345678", "dsis": ["AB","BC"],           "rows": [8,9]}
    ]
    """
    blocks = []
    current_pn = None
    current_block_rows = []
    current_dsis = []

    for row_cells in sheet3.iter_rows(min_row=2):
        pn_val = row_cells[9].value
        dsi_val = row_cells[12].value

        if pn_val:  # new block starts
            if current_pn is not None:
                blocks.append({
                    "pn": current_pn,
                    "dsis": current_dsis,
                    "rows": current_block_rows
                })
            
            current_pn = str(pn_val).strip()
            current_dsis = []
            current_block_rows = []

        # always add row to current block (pn or blank)
        if current_pn:
            current_block_rows.append(row_cells[0].row)
            if dsi_val:
                current_dsis.append(str(dsi_val).strip())
    
    # append last block
    if current_pn:
        blocks.append({
            "pn": current_pn,
            "dsis": current_dsis,
            "rows": current_block_rows
        })

    return blocks



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
            # mark only the FIRST ROW of this block
            row_index = blk["rows"][0]
            sheet3.cell(row=row_index, column=10).fill = red_fill
            add_error_comment_to_PN_cell(sheet3, pn, "Missing PF in this occurrence")
