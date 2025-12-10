def build_PN_has_dsi(sheet3, dsi_check_values, dsi_eq_case_insensitive=True):
    """
    Returns:  PN_blocks = {
         "12345678": [ {"PF","MD"}, {"MD"} ],   # block 1 has PF+MD, block 2 has MD only
         ...
    }
    """

    PN_blocks = {}
    current_PN = None
    current_block = set()

    for row in sheet3.iter_rows(min_row=2):
        
        # New Part Number → start new block
        if row[9].value:
            pn = str(row[9].value).strip()

            # Save previous block before starting new one
            if current_PN is not None:
                PN_blocks.setdefault(current_PN, []).append(current_block)

            # Reset for new block
            current_PN = pn
            current_block = set()

        # Collect DSI in current block
        dsi_tmp = str(row[12].value).strip() if row[12].value else None
        if dsi_tmp:
            if dsi_eq_case_insensitive:
                if dsi_tmp.lower() in {d.lower() for d in dsi_check_values}:
                    current_block.add(dsi_tmp.upper())
            else:
                if dsi_tmp in dsi_check_values:
                    current_block.add(dsi_tmp.upper())

    # Save last block
    if current_PN is not None:
        PN_blocks.setdefault(current_PN, []).append(current_block)

    return PN_blocks


def add_error_comment_to_PN_cell(sheet3, PN_val, text, row):
    """
    Modified: comment added ONLY to the PN cell of THIS block,
    not to all PN occurrences across sheet.
    """
    pn_cell = row[9]   # Only current PN cell, not all occurrences

    existing = pn_cell.comment.text if pn_cell.comment else None
    if existing:
        if text.strip() in existing:
            return
        new_text = existing + "\n" + text
    else:
        new_text = text

    # Auto-size comment
    width_pt = min(600, len(new_text) * 7 + 10)
    height_pt = (new_text.count("\n") + 1) * 15 + 10

    try:
        pn_cell.comment = Comment(new_text, "System", width=width_pt, height=height_pt)
    except:
        c = Comment(new_text, "System")
        pn_cell.comment = c


def mark_missing_PN_red_for_desc(PN_to_desc, PN_blocks, sheet3,
                                 error_msg, normalize_fn):

    red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

    for PN_val in PN_to_desc:
        norm_PN = normalize_fn(PN_val)

        if norm_PN not in PN_blocks:
            continue

        blocks = PN_blocks[norm_PN]

        block_index = -1
        current_block_rows = []

        last_seen_pn = None

        for row in sheet3.iter_rows(min_row=2):

            if row[9].value:
                pn_here = normalize_fn(row[9].value)

                # New block start
                if pn_here == norm_PN:
                    if current_block_rows:
                        # Check previous block
                        if "PF" not in blocks[block_index]:
                            for r in current_block_rows:
                                r[9].fill = red_fill
                            add_error_comment_to_PN_cell(sheet3, PN_val,
                                                         error_msg, current_block_rows[0])

                    block_index += 1
                    current_block_rows = []

            # Still in same PN block?
            if row[9].value is None and last_seen_pn == norm_PN:
                current_block_rows.append(row)

            last_seen_pn = normalize_fn(row[9].value) if row[9].value else last_seen_pn

        # Check last block
       


def mark_missing_PN_red_by_presence(PN_set, PN_blocks, sheet3,
                                    error_msg, check_dsi=False):
    """
    BLOCKWISE VERSION:
    - PN_set = PN list from STEP DATA (PN_to_desc keys)
    - PN_blocks = blockwise DSIs detected in REPORT DATA from build_PN_has_dsi()
    - If a required DSI is missing in *a specific block*, only that block is marked.

    check_dsi=True → skip deleted items (T column = 'D')
    """

    red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

    for PN_val in PN_set:
        norm = normalize_PN(PN_val)

        # If PN not present in sheet3 → no block → skip
        if norm not in PN_blocks:
            continue

        pn_block_list = PN_blocks[norm]   # list of sets (each block)

        block_index = -1
        current_block_rows = []
        last_seen = None

        for row in sheet3.iter_rows(min_row=2):
            # New block begins
            if row[9].value:
                pn_here = normalize_PN(row[9].value)

                if pn_here == norm:
                    # Check previous block
                    if current_block_rows and block_index < len(pn_block_list):
                        block_has = pn_block_list[block_index]

                        if len(block_has) == 0:   # NO DSI found in this block
                            # but skip deleted item blocks
                            t_val = str(current_block_rows[0][3].value).strip() if current_block_rows[0][3].value else ""
                            if not (check_dsi and t_val == "D"):
                                for r in current_block_rows:
                                    r[9].fill = red_fill

                                add_error_comment_to_PN_cell(
                                    sheet3, PN_val, error_msg, current_block_rows[0]
                                )

                    block_index += 1
                    current_block_rows = []

            # Continue collecting rows of block
            if last_seen == norm and row[9].value is None:
                current_block_rows.append(row)

            last_seen = normalize_PN(row[9].value) if row[9].value else last_seen

        # Check final block
        if current_block_rows and block_index < len(pn_block_list):
            block_has = pn_block_list[block_index]

            if len(block_has) == 0:
                t_val = str(current_block_rows[0][3].value).strip() if current_block_rows[0][3].value else ""
                if not (check_dsi and t_val == "D"):
                    for r in current_block_rows:
                        r[9].fill = red_fill
                    add_error_comment_to_PN_cell(sheet3, PN_val, error_msg, current_block_rows[0])


def compare_and_color(sheet3, PN_to_desc, dsi_match,
                      error_msg_unmatched, normalize_fn):

    last_pn = None
    current_block_rows = []

    for row in sheet3.iter_rows(min_row=2):

        # Detect new block start
        if row[9].value:
            # Reset block tracking
            last_pn = normalize_fn(row[9].value)
            current_block_rows = []

        if last_pn is None:
            continue

        dsi = str(row[12].value).strip() if row[12].value else None
        nomen = str(row[11].value).strip() if row[11].value else None
        t_val = str(row[3].value).strip() if row[3].value else None

        # Skip deleted
        if t_val == "D":
            continue

        # Only match rows having this DSI
        if not (dsi and dsi.upper() == dsi_match.upper()):
            continue

        # PN must exist in STEP data dictionary
        if last_pn not in PN_to_desc:
            # unmatched occurrence → error
            row[11].fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
            row[12].fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
            add_error_comment_to_PN_cell(sheet3, last_pn, f"{dsi_match} found in FCR but not in STEP", row)
            continue

        desc = PN_to_desc[last_pn]

        # For PF/MD/QD compare text; for MU/MN compare normalized PN
        if dsi_match.upper() in ("PF", "MD", "QD"):
            left = normalize(nomen)
            right = normalize(desc)
        else:
            left = normalize_fn(nomen)
            right = normalize_fn(desc)

        # Match?
        if left == right:
            row[11].fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
            row[12].fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
        else:
            row[11].fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
            row[12].fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

            add_error_comment_to_PN_cell(sheet3, last_pn, error_msg_unmatched, row)


# -----------------------------------------------------------
#   PF — BLOCKWISE STEP–vs–REPORT VALIDATION
# -----------------------------------------------------------

# 1⃣  Extract PF descriptions from STEP DATA
PN_to_PF_desc, _ = build_PN_desc(
    self.sheet2,
    header_marker="PN",
    desc_marker="PF",
    follow_marker_check="PF",
    desc_cols_slice=slice(3, 12)
)

# 2⃣  Build blockwise PF presence from REPORT DATA
#     result = { "12345678": [ {"PF"}, {}, {"PF"} ] }
#     means PN 12345678 has 3 independent blocks
PN_blocks_PF = build_PN_blocks(self.sheet3, ["PF"])


# -----------------------------------------------------------
# 3⃣  BLOCKWISE PRESENCE CHECK
#     If a PN has PF in STEP DATA → every block MUST have PF
#     If a PN has NO PF in STEP DATA → NO block must contain PF
# -----------------------------------------------------------

for PN_val, blocks in PN_blocks_PF.items():

    normPN = normalize_PN(PN_val)
    step_has_pf = normPN in {normalize_PN(x) for x in PN_to_PF_desc}

    for block_index, block_dsi_set in enumerate(blocks):

        # Extract the rows belonging to this block
        block_rows = get_block_rows(self.sheet3, PN_val, block_index)

        # Skip deleted block
        if block_rows:
            t_val = str(block_rows[0][3].value).strip() if block_rows[0][3].value else ""
            if t_val == "D":
                continue

        if step_has_pf:
            # STEP → PF required, but REPORT block missing PF
            if "PF" not in block_dsi_set:
                for r in block_rows:
                    r[9].fill = RED
                add_error_comment_to_PN_cell(
                    self.sheet3, PN_val,
                    "PF missing in this block (but PF exists in STEP DATA)",
                    pn_cell=block_rows[0][9]
                )
        else:
            # STEP → PF NOT present, but REPORT block contains PF → ERROR
            if "PF" in block_dsi_set:
                for r in block_rows:
                    r[9].fill = RED
                add_error_comment_to_PN_cell(
                    self.sheet3, PN_val,
                    "PF found in Report Data but not in Step Data",
                    pn_cell=block_rows[0][9]
                )


# -----------------------------------------------------------
# 4⃣  BLOCKWISE PF DESCRIPTION MATCH
# -----------------------------------------------------------

for row in self.sheet3.iter_rows(min_row=2):

    if row[9].value:  # Start of new block
        last_pn = normalize_PN(row[9].value)
        block_index = 0
        block_rows = get_block_rows(self.sheet3, last_pn, block_index)

    dsi = str(row[12].value).strip() if row[12].value else ""
    nomen = str(row[11].value).strip() if row[11].value else ""
    t_val = str(row[3].value).strip() if row[3].value else ""

    # Skip deleted blocks
    if t_val == "D":
        continue

    if dsi.upper() != "PF":
        continue

    # PN must exist in STEP DATA PF dictionary
    if last_pn not in PN_to_PF_desc:
        row[11].fill = RED
        row[12].fill = RED
        add_error_comment_to_PN_cell(
            self.sheet3, last_pn,
            "PF found in FCR but PF description not present in STEP DATA",
            pn_cell=row[9]
        )
        continue

    expected_desc = normalize(PN_to_PF_desc[last_pn])
    actual_desc = normalize(nomen)

    # Compare
    if actual_desc == expected_desc:
        row[11].fill = GREEN
        row[12].fill = GREEN
    else:
        row[11].fill = RED
        row[12].fill = RED
        add_error_comment_to_PN_cell(
            self.sheet3, last_pn,
            "PF description unmatched",
            pn_cell=row[9]
        )


# -----------------------------------------------------------
#                     MD BLOCK (UPDATED)
# -----------------------------------------------------------

# 1️⃣ Extract MD descriptions from STEP DATA
PN_to_MD_desc, PN_header_cell6 = build_PN_desc(
    self.sheet2,
    header_marker="PN",
    desc_marker="MD",
    follow_marker_check="PF",      # your format
    desc_cols_slice=slice(3, 12)
)

# 2️⃣ Build MD presence map in report data (not blockwise)
PN_has_MD_in_sheet3 = build_PN_has_dsi(self.sheet3, ["MD"])

# 3️⃣ Build MD block map:  { "123…": [ {"MD"}, {}, {"MD"} ] }
PN_blocks_MD = build_PN_blocks(self.sheet3, ["MD"])

# -----------------------------------------------------------
# 4️⃣ BLOCKWISE PRESENCE CHECK
# -----------------------------------------------------------

for pn, blocks in PN_blocks_MD.items():

    norm_pn = normalize_PN(pn)
    step_has_md = norm_pn in {normalize_PN(x) for x in PN_to_MD_desc}

    for bi, block_set in enumerate(blocks):

        # rows belonging to the MD block
        block_rows = get_block_rows(self.sheet3, pn, bi)
        if not block_rows:
            continue

        # Skip deleted block
        t_val = str(block_rows[0][3].value).strip() if block_rows[0][3].value else ""
        if t_val == "D":
            continue

        # CASE A — STEP DATA HAS MD → block must contain MD
        if step_has_md:
            if "MD" not in block_set:
                pn_cell = block_rows[0][9]
                pn_cell.fill = RED
                add_error_comment_to_PN_cell(
                    self.sheet3, pn,
                    "MD missing in this block (MD exists in Step Data)",
                    pn_cell=pn_cell
                )
        else:
            # CASE B — STEP DATA HAS NO MD → block must not contain MD
            if "MD" in block_set:
                pn_cell = block_rows[0][9]
                pn_cell.fill = RED
                add_error_comment_to_PN_cell(
                    self.sheet3, pn,
                    "MD found in Report Data but not present in Step Data",
                    pn_cell=pn_cell
                )

# -----------------------------------------------------------
# 5️⃣ BLOCKWISE MD DESCRIPTION MATCHING
# -----------------------------------------------------------

last_pn = None

for row in self.sheet3.iter_rows(min_row=2):

    if row[9].value:  # new PN block starts
        last_pn = normalize_PN(row[9].value)

    dsi = str(row[12].value).strip() if row[12].value else ""
    nomen = str(row[11].value).strip() if row[11].value else ""
    t_val = str(row[3].value).strip() if row[3].value else ""

    # Skip deleted rows
    if t_val == "D":
        continue

    # Only MD rows
    if dsi.upper() != "MD":
        continue

    # If STEP DATA has no MD description for this PN
    if last_pn not in PN_to_MD_desc:
        row[11].fill = RED
        row[12].fill = RED
        add_error_comment_to_PN_cell(
            self.sheet3, last_pn,
            "MD present in Report Data but MD not defined in Step Data",
            pn_cell=row[9]
        )
        continue

    expected_desc = normalize(PN_to_MD_desc[last_pn])
    actual_desc = normalize(nomen)

    # Match descriptions
    if actual_desc == expected_desc:
        row[11].fill = GREEN
        row[12].fill = GREEN
    else:
        row[11].fill = RED
        row[12].fill = RED
        add_error_comment_to_PN_cell(
            self.sheet3, last_pn,
            "MD description unmatched",
            pn_cell=row[9]
        )

# -----------------------------------------------------------
#                     SD BLOCK (UPDATED)
# -----------------------------------------------------------

# 1️⃣ Extract SD descriptions from STEP DATA
PN_to_SD_desc = build_PN_desc_simple(
    self.sheet2,
    header_marker="PN",
    desc_marker="14",             # your sheet2 pattern for SD
    desc_cols_slice=slice(3, 12)
)

# 2️⃣ Map which PNs have SD anywhere in sheet3 (not blockwise)
PN_has_SD_in_sheet3 = build_PN_has_dsi(self.sheet3, ["SD"])

# 3️⃣ Build SD block map from sheet3
PN_blocks_SD = build_PN_blocks(self.sheet3, ["SD"])

# -----------------------------------------------------------
# 4️⃣ BLOCKWISE PRESENCE CHECKING
# -----------------------------------------------------------

for pn, blocks in PN_blocks_SD.items():

    norm_pn = normalize_PN(pn)
    step_has_sd = norm_pn in {normalize_PN(x) for x in PN_to_SD_desc}

    for bi, block_set in enumerate(blocks):

        block_rows = get_block_rows(self.sheet3, pn, bi)
        if not block_rows:
            continue

        # skip deleted blocks
        t_val = str(block_rows[0][3].value).strip() if block_rows[0][3].value else ""
        if t_val == "D":
            continue

        # CASE A — Step Data HAS SD → block must contain SD
        if step_has_sd:
            if "SD" not in block_set:
                pn_cell = block_rows[0][9]
                pn_cell.fill = RED
                add_error_comment_to_PN_cell(
                    self.sheet3,
                    pn,
                    "SD missing in this block (SD exists in Step Data)",
                    pn_cell=pn_cell
                )

        # CASE B — Step Data does NOT have SD → block must NOT contain SD
        else:
            if "SD" in block_set:
                pn_cell = block_rows[0][9]
                pn_cell.fill = RED
                add_error_comment_to_PN_cell(
                    self.sheet3,
                    pn,
                    "SD found in Report Data but not present in Step Data",
                    pn_cell=pn_cell
                )

# -----------------------------------------------------------
# 5️⃣ BLOCKWISE SD DESCRIPTION MATCHING
# -----------------------------------------------------------

last_pn = None

for row in self.sheet3.iter_rows(min_row=2):

    if row[9].value:
        last_pn = normalize_PN(row[9].value)

    dsi = str(row[12].value).strip() if row[12].value else ""
    nomen = str(row[11].value).strip() if row[11].value else ""
    t_val = str(row[3].value).strip() if row[3].value else ""

    # skip deleted
    if t_val == "D":
        continue

    # only SD rows
    if dsi.upper() != "SD":
        continue

    # If SD not present in Step Data → description unmatched automatically
    if last_pn not in PN_to_SD_desc:
        row[11].fill = RED
        row[12].fill = RED
        add_error_comment_to_PN_cell(
            self.sheet3, last_pn,
            "SD exists in Report Data but not in Step Data",
            pn_cell=row[9]
        )
        continue

    expected = normalize(PN_to_SD_desc[last_pn])
    actual = normalize(nomen)

    # Match descriptions
    if expected == actual:
        row[11].fill = GREEN
        row[12].fill = GREEN
    else:
        row[11].fill = RED
        row[12].fill = RED
        add_error_comment_to_PN_cell(
            self.sheet3,
            last_pn,
            "SD description unmatched",
            pn_cell=row[9]
        )


