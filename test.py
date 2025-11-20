def extract_pn_from_L(value):
    if not value:
        return None
    parts = str(value).strip().split()
    if len(parts) < 2:
        return None
    return parts[-1].strip()  # last token is always PN

def get_child_PN_for_IW1(sheet2, parent_pn):
    ws = sheet2
    pending = False
    last_pn = None
    in_direct = False
    DI_header = None
    PN_col = None
    IW_col = None

    for row in ws.iter_rows(min_row=1):
        first = str(row[0].value).strip() if row[0].value else None

        if first == "PN":
            pending = True
            continue

        if pending:
            last_pn = first
            pending = False
            continue

        if last_pn != parent_pn:
            continue

        if first and first.lower() == "direct/indirect":
            in_direct = True
            DI_header = row
            continue

        if in_direct:
            if PN_col is None or IW_col is None:
                # detect PN and IW columns
                for idx, c in enumerate(DI_header):
                    val = str(c.value).strip() if c.value else None
                    if val == "PN" and PN_col is None:
                        PN_col = idx
                    if val == "IW" and IW_col is None:
                        IW_col = idx

            # next parent block detected -> stop
            if first == "PN":
                return None

            iw_cell = row[IW_col].value if IW_col is not None else None
            try:
                iw_val = int(float(str(iw_cell).strip()))
            except:
                iw_val = None

            if iw_val == 1:
                child = row[PN_col].value
                return str(child).strip() if child else None

    return None


if dsi_val and dsi_val.upper() in ("MU", "MN"):

    # Ignore T == D case
    if t_val == "D":
        for cell in row:
            cell.fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
        continue

    # Extract PN from column L
    extracted = extract_pn_from_L(row[11].value)

    # Get expected PN from sheet2 (IW=1)
    expected = get_child_PN_for_IW1(self.sheet2, part_no4)

    # Compare
    if extracted and expected and normalize_PN(extracted) == normalize_PN(expected):
        for cell in row:
            cell.fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
    else:
        for cell in row:
            cell.fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
        add_error_comment_to_PN_cell(
            self.sheet3,
            part_no4,
            f"MU mismatch: Expected {expected}, Found {extracted}"
        )

    continue
