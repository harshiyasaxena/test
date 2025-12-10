def build_PN_blocks(sheet3, normalize_fn=normalize_PN):
    """
    Splits sheet3 into independent PN blocks.
    Each block begins at a non-blank PN cell and continues until the next PN.
    Blank rows belong to the last seen PN only.
    """
    blocks = []
    current_block = []
    current_pn = None

    for row in sheet3.iter_rows(min_row=2):
        raw = row[9].value

        if raw:  # A new PN → close previous block
            pn = normalize_fn(raw)

            if current_block:
                blocks.append((current_pn, current_block))

            current_pn = pn
            current_block = [row]

        else:
            if current_pn:
                current_block.append(row)

    # Add last block
    if current_block:
        blocks.append((current_pn, current_block))

    return blocks



def compare_and_color_block(block_rows, PN_to_desc, dsi_match,
                            error_msg_unmatched, normalize_fn):

    last_part_no = None
    parent_t_val = None

    for row in block_rows:

        if row[9].value:
            last_part_no = normalize_fn(row[9].value)
            parent_t_val = str(row[3].value).strip() if row[3].value else None

        part_no = last_part_no
        nomen = str(row[11].value).strip() if row[11].value else ""
        dsi = str(row[12].value).strip() if row[12].value else ""

        if parent_t_val == "D":
            continue

        if dsi.upper() == dsi_match.upper() and part_no in PN_to_desc:

            desc = PN_to_desc[part_no]

            # PF/MD/QD require full normalize text
            if dsi_match.upper() in ("PF", "MD", "QD", "QDQ"):
                left = normalize(nomen)
                right = normalize(desc)
            else:
                left = normalize_fn(nomen)
                right = normalize_fn(desc)

            if left == right:
                row[11].fill = PatternFill("00FF00")
                row[12].fill = PatternFill("00FF00")
            else:
                row[11].fill = PatternFill("FF0000")
                row[12].fill = PatternFill("FF0000")
                add_error_comment_to_PN_cell(
                    self.sheet3, part_no, error_msg_unmatched
                )


PN_blocks = build_PN_blocks(self.sheet3, normalize_PN)

for pn, block_rows in PN_blocks:

    if pn not in PN_to_desc:
        continue

    # Check if this block contains PF
    block_has_pf = any(
        str(r[12].value).strip().upper() == "PF"
        for r in block_rows
    )

    # Step 1 — Missing PF in this block only
    if not block_has_pf:
        first_row = block_rows[0]
        first_row[9].fill = PatternFill("FF0000")
        add_error_comment_to_PN_cell(self.sheet3, pn, "Missing PF in FCR")
        continue

    # Step 2 — Compare the PF rows inside this block only
    compare_and_color_block(
        block_rows,
        PN_to_desc,
        "PF",
        "PF description unmatched",
        normalize_PN
    )
