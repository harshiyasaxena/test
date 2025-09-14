def validate_and_color(self):
    wb = load_workbook(self.excel_file)
    ws_data = wb["DATA"]
    ws_inputs = wb["inputs"]

    # Step 1: Extract JK → Description mapping
    jk_to_desc = {}
    current_jk = None
    for row in ws_data.iter_rows(min_row=2, values_only=True):
        if row[0]:  # Column A has value
            current_jk = str(row[0]).strip()

        if current_jk and str(row[0]).strip() == "KJ":
            # Collect description text from cols D–L (index 3–11)
            desc_parts = [str(cell).strip() for cell in row[3:12] if cell]
            desc_text = " ".join(desc_parts)
            jk_to_desc[current_jk] = desc_text
            print(f"[DATA] JK={current_jk}, Description={desc_text}")

    # Step 2: Loop over INPUTS sheet
    for row in ws_inputs.iter_rows(min_row=2):
        fish_ko = str(row[9].value).strip() if row[9].value else None   # col J
        hehe = str(row[11].value).strip() if row[11].value else None   # col L
        usp = str(row[12].value).strip() if row[12].value else None    # col M

        print(f"[INPUTS] Fish Ko={fish_ko}, HEHE={hehe}, USP={usp}")

        if usp == "KJ" and fish_ko in jk_to_desc:
            desc = jk_to_desc[fish_ko]
            print(f"   Comparing HEHE='{hehe}' with DESC='{desc}'")
            if hehe == desc:
                row[11].fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
                print("   → GREEN")
            else:
                row[11].fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
                print("   → RED")

    wb.save(self.excel_file)
