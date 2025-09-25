def check_wy_dsi(self):
    from openpyxl.styles import PatternFill

    pn_to_wy = {}
    last_pn = None
    pending_pn = None

    # --- Step 1: Map PN to WY value from Sheet2 ---
    for row in self.sheet2.iter_rows(min_row=1):
        a_val = str(row[0].value).strip() if row[0].value else None

        if not a_val:
            continue

        # "JK" means next row has PN
        if a_val.upper() == "JK":
            pending_pn = True
            continue

        if pending_pn:
            last_pn = a_val
            pending_pn = False
            continue

        # Search the row for "WY" (not only col A!)
        for cell in row:
            if str(cell.value).strip().upper() == "WY" and last_pn:
                wy_cell = self.sheet2.cell(row=cell.row + 1, column=cell.column)  # value below WY
                wy_val = str(wy_cell.value).strip() if wy_cell.value else None
                if wy_val:
                    try:
                        wy_val = int(wy_val)
                    except:
                        wy_val = 0
                    pn_to_wy[last_pn] = wy_val
                    print(f"Mapped PN={last_pn} -> WY={wy_val} (from row {cell.row})")
                last_pn = None

    # --- Step 2: Walk Sheet3 and apply WY+DSI rules ---
    last_part_no = None
    for row in self.sheet3.iter_rows(min_row=2):
        if row[9].value:  # Part number in col J
            last_part_no = str(row[9].value).strip()

        part_no = last_part_no
        if not part_no or part_no not in pn_to_wy:
            continue

        wy_val = pn_to_wy[part_no]

        # DSI col = M (index 12)
        dsi_val = str(row[12].value).strip() if row[12].value else None
        # T col = col T (index 19)
        t_val = str(row[19].value).strip() if row[19].value else None

        if dsi_val == "MU":
            if wy_val == 1:
                for cell in row:
                    cell.fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
                print(f"[GREEN] PN={part_no}, WY=1, DSI=MU (row {row[0].row})")
            else:
                if t_val == "D":
                    for cell in row:
                        cell.fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
                    print(f"[GREEN] PN={part_no}, WY={wy_val}, DSI=MU, T=D (row {row[0].row})")
                else:
                    for cell in row:
                        cell.fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
                    print(f"[RED] PN={part_no}, WY={wy_val}, DSI=MU, T={t_val} (row {row[0].row})")
