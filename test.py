# ------------------ START: Block-wise DSI helpers & DSI processing ------------------

def add_error_comment_to_PN_cell(sheet3, PN_val, text, pn_cell, _comment_size_for_text_fn=None):
    """
    Attach comment ONLY to this block's PN cell (pn_cell is row[9] of that block).
    _comment_size_for_text_fn: function to compute comment width/height (use your existing helper)
    """
    if pn_cell is None:
        return

    existing = pn_cell.comment.text if pn_cell.comment else None

    if existing:
        existing_lines = {ln.strip() for ln in existing.splitlines() if ln.strip()}
        if text.strip() in existing_lines:
            return
        new_text = existing + "\n" + text
    else:
        new_text = text

    # determine width/height if helper available
    if _comment_size_for_text_fn:
        width_pt, height_pt = _comment_size_for_text_fn(new_text)
    else:
        # sensible defaults if helper missing
        width_pt, height_pt = (300, 40)

    try:
        pn_cell.comment = Comment(new_text, author="System", width=width_pt, height=height_pt)
    except TypeError:
        c = Comment(new_text, author="System")
        try:
            c.width = width_pt
            c.height = height_pt
        except Exception:
            pass
        pn_cell.comment = c


def build_PN_has_dsi(sheet3, dsi_check_values, normalize_fn):
    """
    Return: dict pn -> list of booleans per block
      e.g. { "12345678": [True, False, True], ... }
    Block boundaries: row where sheet3.colJ (index 9) has a value starts a new block.
    """
    PN_has = {}
    last_pn = None
    current_block_idx = -1

    for row in sheet3.iter_rows(min_row=2):
        pn_cell = row[9].value
        if pn_cell:
            last_pn = normalize_fn(pn_cell)
            PN_has.setdefault(last_pn, [])
            PN_has[last_pn].append(False)
            current_block_idx = len(PN_has[last_pn]) - 1

        if not last_pn:
            continue

        dsi_val = row[12].value
        if not dsi_val:
            continue
        dsi_str = str(dsi_val).strip()

        # case-insensitive compare
        if any(dsi_str.lower() == v.lower() for v in dsi_check_values):
            PN_has[last_pn][current_block_idx] = True

    return PN_has


def mark_missing_PN_red_for_desc(PN_to_desc, PN_has_blocks, sheet3,
                                 error_msg, normalize_fn, block_map,
                                 comment_size_fn=None):
    """
    Mark PN (the PN cell of the block) red when expected DSI (PF/MD etc.) is missing for that block.
    PN_to_desc: mapping from STEP DATA PN -> expected description (those PNs that must have DSI)
    PN_has_blocks: output of build_PN_has_dsi (block-aware)
    block_map: list of block index per sheet3 row (0-based). length == number of rows starting at min_row=2
    """
    red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

    for idx, row in enumerate(sheet3.iter_rows(min_row=2)):
        pn_cell = row[9]
        if not pn_cell.value:
            continue
        norm_pn = normalize_fn(pn_cell.value)
        current_block = block_map[idx]
        # skip if PN not in step-data list (we only check PNs expected to have desc)
        if norm_pn not in PN_to_desc:
            continue
        blocks = PN_has_blocks.get(norm_pn, [])
        missing = (current_block >= len(blocks)) or (blocks[current_block] is False)
        if missing:
            # color only this block's PN cell and attach comment here
            try:
                pn_cell.fill = red_fill
            except Exception:
                pass
            add_error_comment_to_PN_cell(sheet3, norm_pn, error_msg, pn_cell, comment_size_fn)


def mark_missing_PN_red_by_presence(PN_set, PN_has_blocks, sheet3, error_msg,
                                    block_map, normalize_fn, comment_size_fn=None):
    """
    General presence-based missing checker (used for SD/QD/MU/MN/RB/RO):
    PN_set: set/list of PNs (from step data) expected to have the DSI
    PN_has_blocks: dict pn -> [True/False...]
    block_map: row-index -> block number (from the block map builder)
    """
    red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    for idx, row in enumerate(sheet3.iter_rows(min_row=2)):
        pn_cell = row[9]
        if not pn_cell.value:
            continue
        norm_pn = normalize_fn(pn_cell.value)
        if norm_pn not in PN_set:
            continue
        current_block = block_map[idx]
        blocks = PN_has_blocks.get(norm_pn, [])
        missing = (current_block >= len(blocks)) or (blocks[current_block] is False)
        if missing:
            try:
                pn_cell.fill = red_fill
            except Exception:
                pass
            add_error_comment_to_PN_cell(sheet3, norm_pn, error_msg, pn_cell, comment_size_fn)


def compare_and_color(sheet3, PN_to_desc, dsi_match, error_msg_unmatched,
                      normalize_fn, block_map, PN_has_blocks, comment_size_fn=None):
    """
    Block-aware compare and color function.
    - PN_to_desc: mapping PN -> expected description (from step/STEP DATA)
    - dsi_match: e.g. "PF" or "MD" etc. (case-insensitive)
    - PN_has_blocks: dict pn -> [True/False per block] (output of build_PN_has_dsi)
    - block_map: row-index -> block number
    """
    green_fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
    red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

    last_pn = None
    parent_t_val = None

    for idx, row in enumerate(sheet3.iter_rows(min_row=2)):
        # update PN when a PN cell appears
        if row[9].value:
            last_pn = normalize_fn(row[9].value)
            parent_t_val = str(row[3].value).strip() if row[3].value else None

        if last_pn is None:
            continue

        # Only compare if this PN has an expected description
        if last_pn not in PN_to_desc:
            continue

        current_block = block_map[idx]
        if current_block is None:
            continue

        # If this block does NOT have the DSI (according to PN_has_blocks) skip comparison
        blocks = PN_has_blocks.get(last_pn, [])
        if current_block >= len(blocks) or blocks[current_block] is False:
            continue

        # Now check the row's DSI and nomenclature
        dsi = str(row[12].value).strip() if row[12].value else None
        nomen = str(row[11].value).strip() if row[11].value else ""

        if dsi and dsi.lower() == dsi_match.lower():
            # For PF/MD/QDQ you used normalized comparison in original; preserve that behavior:
            if dsi_match.upper() in ("PF", "MD", "QDQ"):
                left = " ".join(nomen.split())
                right = " ".join(str(PN_to_desc[last_pn]).split()) if PN_to_desc[last_pn] else ""
            else:
                left = normalize_fn(nomen) if nomen else ""
                right = normalize_fn(PN_to_desc[last_pn]) if PN_to_desc[last_pn] else ""

            if left == right:
                try:
                    row[11].fill = green_fill
                    row[12].fill = green_fill
                except Exception:
                    pass
            else:
                try:
                    row[11].fill = red_fill
                    row[12].fill = red_fill
                except Exception:
                    pass
                # attach comment to the PN cell of this row's block
                add_error_comment_to_PN_cell(sheet3, last_pn, error_msg_unmatched, row[9], comment_size_fn)


# ------------------ END: Block-wise DSI helpers & common functions ------------------


# ------------------ DSI processing snippet to paste into your main_code() ------------------
# Place this snippet where you previously handled PF/MD/SD/QD/MU/MN/RB/RO. It uses:
#   - normalize_PN (must exist)
#   - build_PN_desc and other STEP DATA builders (must exist)
#   - _comment_size_for_text (optional, pass it in if available)

# 1) Build block map once for sheet3 (list aligned with rows min_row=2)
PN_block_map = []
block_counter = {}
current_block = None
for row in self.sheet3.iter_rows(min_row=2):
    pn_val = row[9].value
    if pn_val:
        pn_norm = normalize_PN(pn_val)
        block_counter[pn_norm] = block_counter.get(pn_norm, 0) + 1
        current_block = block_counter[pn_norm] - 1
    PN_block_map.append(current_block)

# 2) Build DSI presence maps for each DSI type (block-aware)
PN_has_PF = build_PN_has_dsi(self.sheet3, ["PF"], normalize_PN)
PN_has_MD = build_PN_has_dsi(self.sheet3, ["MD"], normalize_PN)
PN_has_SD = build_PN_has_dsi(self.sheet3, ["SD"], normalize_PN)
PN_has_QD = build_PN_has_dsi(self.sheet3, ["QD"], normalize_PN)
PN_has_MU_MN = build_PN_has_dsi(self.sheet3, ["MU", "MN"], normalize_PN)
PN_has_RB_RO = build_PN_has_dsi(self.sheet3, ["RB", "RO"], normalize_PN)

# 3) PF: build PN->desc mapping from sheet2 (your existing function)
PN_to_desc_pf, PN_header_cell_pf = build_PN_desc(self.sheet2, "PN", "PF", "PF", slice(3,12))
# mark missing PF per-block
mark_missing_PN_red_for_desc(PN_to_desc_pf, PN_has_PF, self.sheet3,
                             "Missing PF in FCR", normalize_PN, PN_block_map,
                             comment_size_fn=_comment_size_for_text if '_comment_size_for_text' in globals() else None)
# compare PF descriptions block-wise
compare_and_color(self.sheet3, PN_to_desc_pf, "PF", "PF description unmatched",
                  normalize_PN, PN_block_map, PN_has_PF,
                  comment_size_fn=_comment_size_for_text if '_comment_size_for_text' in globals() else None)

# 4) MD
PN_to_desc_md, PN_header_cell_md = build_PN_desc(self.sheet2, "PN", "MD", "PF", slice(3,12))
mark_missing_PN_red_for_desc(PN_to_desc_md, PN_has_MD, self.sheet3,
                             "Missing MD in FCR", normalize_PN, PN_block_map,
                             comment_size_fn=_comment_size_for_text if '_comment_size_for_text' in globals() else None)
compare_and_color(self.sheet3, PN_to_desc_md, "MD", "MD description unmatched",
                  normalize_PN, PN_block_map, PN_has_MD,
                  comment_size_fn=_comment_size_for_text if '_comment_size_for_text' in globals() else None)

# 5) SD presence-based (PN_to_desc from build_PN_desc_simple)
PN_to_desc_sd = build_PN_desc_simple(self.sheet2, "PN", "14", slice(3,12))
# For SD we used mark_missing_PN_red_by_presence in earlier code — use it block-wise:
mark_missing_PN_red_by_presence(set(PN_to_desc_sd.keys()), PN_has_SD, self.sheet3,
                                "Missing SD in FCR", PN_block_map, normalize_PN,
                                comment_size_fn=_comment_size_for_text if '_comment_size_for_text' in globals() else None)
compare_and_color(self.sheet3, PN_to_desc_sd, "SD", "SD description unmatched",
                  normalize_PN, PN_block_map, PN_has_SD,
                  comment_size_fn=_comment_size_for_text if '_comment_size_for_text' in globals() else None)

# 6) QD
PN_to_desc_qd, PN_iw_codes = build_PN_to_desc_iw_iwdesc(self.sheet2)
mark_missing_PN_red_by_presence(set(PN_to_desc_qd.keys()), PN_has_QD, self.sheet3,
                                "Missing QD in FCR", PN_block_map, normalize_PN,
                                comment_size_fn=_comment_size_for_text if '_comment_size_for_text' in globals() else None)
compare_and_color(self.sheet3, PN_to_desc_qd, "QD", "QD description unmatched",
                  normalize_PN, PN_block_map, PN_has_QD,
                  comment_size_fn=_comment_size_for_text if '_comment_size_for_text' in globals() else None)

# 7) MU/MN (presence and compare)
PN_to_desc_mu_mn, PN_with_iw_one = build_PN_to_desc_iw_variants(self.sheet2, header_marker="PN", iw_marker="IW", target_value=1, stop_values=("PN",))
mark_missing_PN_red_by_presence(set(PN_to_desc_mu_mn.keys()), PN_has_MU_MN, self.sheet3,
                                "Missing MU/MN in FCR", PN_block_map, normalize_PN,
                                comment_size_fn=_comment_size_for_text if '_comment_size_for_text' in globals() else None)
compare_and_color(self.sheet3, PN_to_desc_mu_mn, "MU", "MU part no unmatched",
                  normalize_PN, PN_block_map, PN_has_MU_MN,
                  comment_size_fn=_comment_size_for_text if '_comment_size_for_text' in globals() else None)
compare_and_color(self.sheet3, PN_to_desc_mu_mn, "MN", "MN part no unmatched",
                  normalize_PN, PN_block_map, PN_has_MU_MN,
                  comment_size_fn=_comment_size_for_text if '_comment_size_for_text' in globals() else None)

# 8) RB/RO
PN_to_desc_rb_ro, PN_with_iw_two = build_PN_to_desc_iw_variants(self.sheet2, header_marker="PN", iw_marker="IW", target_value=2, stop_values=("PN",))
mark_missing_PN_red_by_presence(set(PN_to_desc_rb_ro.keys()), PN_has_RB_RO, self.sheet3,
                                "Missing RB/RO in FCR", PN_block_map, normalize_PN,
                                comment_size_fn=_comment_size_for_text if '_comment_size_for_text' in globals() else None)
compare_and_color(self.sheet3, PN_to_desc_rb_ro, "RB", "RB part no unmatched",
                  normalize_PN, PN_block_map, PN_has_RB_RO,
                  comment_size_fn=_comment_size_for_text if '_comment_size_for_text' in globals() else None)
compare_and_color(self.sheet3, PN_to_desc_rb_ro, "RO", "RO part no unmatched",
                  normalize_PN, PN_block_map, PN_has_RB_RO,
                  comment_size_fn=_comment_size_for_text if '_comment_size_for_text' in globals() else None)

# ------------------ End DSI processing snippet ------------------


# If you want the complete program's execution block (entrypoint), keep your existing one.
# For convenience, here's the typical bottom lines:
if __name__ == "__main__":
    app = App()
    app.run()
