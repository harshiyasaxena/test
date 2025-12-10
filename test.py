def build_PN_has_dsi(sheet3, dsi_check_values, dsi_eq_case_insensitive=True):
    """
    Returns: PN_has = { PN : [True/False per block] }
    Block = from a PN row until next PN row.
    """
    PN_has = {}        # PN → list of True/False per block
    last_pn = None
    current_block = -1

    for row in sheet3.iter_rows(min_row=2):
        pn_cell = row[9].value

        # A new PN encountered → new block begins
        if pn_cell:
            last_pn = str(pn_cell).strip()
            PN_has.setdefault(last_pn, [])
            PN_has[last_pn].append(False)     # default: no DSI match in this block yet
            current_block = len(PN_has[last_pn]) - 1

        if not last_pn:
            continue

        dsi_val = str(row[12].value).strip() if row[12].value else None
        if not dsi_val:
            continue

        if dsi_eq_case_insensitive:
            match = dsi_val.lower() in {v.lower() for v in dsi_check_values}
        else:
            match = dsi_val in dsi_check_values

        if match:
            PN_has[last_pn][current_block] = True

    return PN_has


def mark_missing_PN_red_for_desc(PN_to_desc, PN_has_set, sheet3,
                                 error_msg, normalize_fn, block_map):
    """
    Marks PN red only for the block where PF/MD is actually missing.
    PN_has_set = {PN: [True/False per block]}
    block_map = row_index → block number
    """
    red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

    for idx, row in enumerate(sheet3.iter_rows(min_row=2)):
        pn_val = row[9].value
        if not pn_val:
            continue

        norm_PN = normalize_fn(pn_val)
        current_block = block_map[idx]    # block number for this occurrence

        # block-wise detection
        blocks = PN_has_set.get(norm_PN, [])
        missing = (current_block >= len(blocks)) or (blocks[current_block] is False)

        if missing:
            # mark ONLY the PN cell of this block
            row[9].fill = red_fill
            add_error_comment_to_PN_cell(sheet3, norm_PN, error_msg, row[9])


def mark_missing_PN_red_by_presence(PN_set, PN_has_set,
                                    sheet3, error_msg, block_map,
                                    normalize_fn=normalize_PN, check_dsi=False):
    """
    PN_set     = all PNs expected to have a DSI (from step data)
    PN_has_set = {PN: [True/False per block]}
    block_map  = row_index → block number
    """
    red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

    for idx, row in enumerate(sheet3.iter_rows(min_row=2)):
        pn_val = row[9].value
        if not pn_val:
            continue

        norm_PN = normalize_fn(pn_val)
        current_block = block_map[idx]

        if norm_PN not in PN_set:
            continue

        blocks = PN_has_set.get(norm_PN, [])
        missing = (current_block >= len(blocks)) or (blocks[current_block] is False)

        if missing:
            row[9].fill = red_fill
            add_error_comment_to_PN_cell(sheet3, norm_PN, error_msg, row[9])


PN_block_map = []
block_counter = {}
current_block = None

for row in self.sheet3.iter_rows(min_row=2):
    pn = row[9].value
    if pn:
        pn = str(pn).strip()
        block_counter[pn] = block_counter.get(pn, 0) + 1
        current_block = block_counter[pn] - 1
    PN_block_map.append(current_block)
