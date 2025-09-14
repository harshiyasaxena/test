def validate_and_color(self):
    wb = load_workbook(self.excel_file)
    ws_data = wb["DATA"]
    ws_inputs = wb["inputs"]

    # get all JK → Karna Description mappings from DATA sheet
    jk_to_desc = {}
    for row in ws_data.iter_rows(min_row=2, values_only=True):
        jk, _, _, desc = row[0], row[1], row[2], row[3] if len(row) > 3 else None
        if jk:
            jk_to_desc[jk] = desc
            print(f"[DATA] JK={jk}, Karna Description={desc}")

    # now check in inputs sheet
    for row in ws_inputs.iter_rows(min_row=2):
        fish_ko = row[7].value   # column H
        hehe = row[9].value      # column J
        usp = row[10].value      # column K

        print(f"[INPUTS] Fish Ko={fish_ko}, HEHE={hehe}, USP={usp}")

        if usp == "KJ" and fish_ko in jk_to_desc:
            desc = jk_to_desc[fish_ko]
            print(f"   Matching Fish Ko found, desc={desc}")
            if desc == hehe:
                row[9].fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
                print("   → Colored GREEN")
            else:
                row[9].fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
                print("   → Colored RED")

    wb.save(self.excel_file)
