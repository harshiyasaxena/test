# === PF & MD BLOCK VALIDATION PER OCCURRENCE ===

required_DSIs = {}
for pn in PN_to_desc.keys():   # PF exists in sheet2
    required_DSIs.setdefault(normalize_PN(pn), set()).add("PF")

for pn in PN_to_desc6.keys():  # MD exists in sheet2
    required_DSIs.setdefault(normalize_PN(pn), set()).add("MD")

current_pn = None
block_rows = []
seen_DSIs = set()

for row in self.sheet3.iter_rows(min_row=2):

    if row[9].value:  # new PN / start of new block
        if current_pn:
            needed = required_DSIs.get(current_pn, set())
            if needed and not needed.issubset(seen_DSIs):
                # skip deleted parent rows
                parent_t = None
                for r in block_rows:
                    if r[9].value:
                        parent_t = str(r[3].value).strip() if r[3].value else None
                        break
                if parent_t != "D":
                    missing = needed - seen_DSIs
                    for r in block_rows:
                        if r[9].value:  # color only PN cell
                            r[9].fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
                            add_error_comment_to_PN_cell(self.sheet3,
                                                         current_pn,
                                                         f"Missing {', '.join(missing)} in this block")

        # reset for next block
        current_pn = normalize_PN(row[9].value)
        block_rows = []
        seen_DSIs = set()

    block_rows.append(row)
    dsi_here = str(row[12].value).strip().upper() if row[12].value else None
    if dsi_here:
        seen_DSIs.add(dsi_here)

# final block compute
if current_pn:
    needed = required_DSIs.get(current_pn, set())
    if needed and not needed.issubset(seen_DSIs):
        parent_t = None
        for r in block_rows:
            if r[9].value:
                parent_t = str(r[3].value).strip() if r[3].value else None
                break
        if parent_t != "D":
            missing = needed - seen_DSIs
            for r in block_rows:
                if r[9].value:
                    r[9].fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
                    add_error_comment_to_PN_cell(self.sheet3,
                                                 current_pn,
                                                 f"Missing {', '.join(missing)} in this block")
