def build_PN_to_desc_iw_iwdesc(sheet):
    PN_to_desc3 = {}
    PN_iw_with_codes = set()
    last_PN3 = None
    pending_PN3 = None

    for row in sheet.iter_rows(min_row=1):
        a_val = str(row[0].value).strip() if row[0].value else None
        if not a_val:
            continue

        # detect new PN block
        if a_val.upper() == "PN":
            pending_PN3 = True
            continue

        # capture the PN value after "PN"
        if pending_PN3:
            last_PN3 = a_val
            pending_PN3 = False
            continue

        # locate "Direct/Indirect" section
        if a_val.upper() == "DIRECT/ INDIRECT" and last_PN3:
            ws = row[0].parent

            # find "IW" column in this row
            for cell in row:
                if not (cell.value and str(cell.value).strip().upper() == "IW"):
                    continue

                iw_col = cell.column
                v_row = cell.row + 1  # start scanning below IW
                max_row = ws.max_row

                # iterate until next PN block begins
                while v_row <= max_row:
                    first_col_val = (
                        str(ws.cell(row=v_row, column=1).value).strip()
                        if ws.cell(row=v_row, column=1).value
                        else ""
                    )

                    # stop when new PN section starts
                    if first_col_val and first_col_val.upper() == "PN":
                        break

                    val_cell = ws.cell(row=v_row, column=iw_col)
                    if not val_cell.value:
                        v_row += 1
                        continue

                    code_val = str(val_cell.value).strip()

                    # skip standard IW codes (1 or 2)
                    if code_val in ("1", "2"):
                        v_row += 1
                        continue

                    # record this PN as having IW codes (non-1/2)
                    PN_iw_with_codes.add(last_PN3)

                    # get next row (expected description)
                    desc_row = v_row + 1
                    if desc_row > max_row:
                        v_row += 2
                        continue

                    # read description from columns F–K
                    desc_parts = []
                    for c in range(6, 12):
                        v = ws.cell(row=desc_row, column=c).value
                        if v and str(v).strip():
                            desc_parts.append(str(v).strip())

                    desc_text = " ".join(desc_parts).strip()
                    if desc_text:
                        PN_to_desc3.setdefault(last_PN3, []).append(desc_text)
                        print(f"Mapped PN '{last_PN3}' -> '{desc_text}' (IW code: {code_val})")

                    # move 2 rows forward (skip desc row)
                    v_row = desc_row + 1

    return PN_to_desc3, PN_iw_with_codes
