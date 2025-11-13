def check_pn_in_direct_indirect_table(sheet2, target_pn):
    """
    Returns True if target_pn is found under the 'PN' column 
    of the Direct/Indirect section corresponding to that part number.
    """
    for i, row in enumerate(sheet2.iter_rows(min_row=1, values_only=True)):
        # Locate the main PN header block
        if str(row[0]).strip().upper() == "PN" and i + 1 < sheet2.max_row:
            main_pn = str(sheet2.cell(row=i + 2, column=1).value).strip() if sheet2.cell(row=i + 2, column=1).value else ""
            if normalize_PN(main_pn) == normalize_PN(target_pn):
                # We are in the correct PN block for this part number
                for j in range(i + 2, sheet2.max_row + 1):
                    a_val = sheet2.cell(row=j, column=1).value
                    if a_val and str(a_val).strip().upper() == "DIRECT/INDIRECT":
                        # Found the Direct/Indirect section
                        header_row = j
                        pn_col_idx = None

                        # Find which column in this header row has "PN"
                        for cell in sheet2[header_row]:
                            if cell.value and str(cell.value).strip().upper() == "PN":
                                pn_col_idx = cell.column
                                print(f"✅ Found PN column at {pn_col_idx} for PN {target_pn}")
                                break

                        if pn_col_idx is None:
                            print(f"⚠️ No PN column found in Direct/Indirect section for {target_pn}")
                            return False

                        # Check values below that PN column
                        for k in range(header_row + 1, sheet2.max_row + 1):
                            next_a_val = sheet2.cell(row=k, column=1).value
                            if next_a_val and str(next_a_val).strip().upper() == "PN":
                                # new PN block starts → stop
                                break

                            pn_val = sheet2.cell(row=k, column=pn_col_idx).value
                            if pn_val and normalize_PN(pn_val) == normalize_PN(target_pn):
                                print(f"✅ Found {target_pn} under Direct/Indirect PN column (row {k})")
                                return True

                        # if reached here, not found
                        print(f"❌ {target_pn} not found under Direct/Indirect PN column.")
                        return False

    return False
