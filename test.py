def validate_MD_per_block(sheet3, MD_map_steps):
    red = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    green = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")

    blocks = build_MD_blocks_from_fcr(sheet3)

    for blk in blocks:
        pn = blk["pn"]
        parent_row = blk["rows"][0]

        # -------------------------------------------------
        # Skip entire block if parent row is deleted
        # -------------------------------------------------
        t_parent = sheet3.cell(row=parent_row, column=4).value
        if t_parent and str(t_parent).strip().upper() == "D":
            continue

        steps_md_list = MD_map_steps.get(pn, [])

        # -------------------------------------------------
        # Collect MD rows (ignore deleted rows)
        # -------------------------------------------------
        md_rows = []
        for r in blk["rows"]:
            dsi_val = sheet3.cell(row=r, column=13).value
            if dsi_val and str(dsi_val).strip().upper() == "MD":

                t_val = sheet3.cell(row=r, column=4).value
                if t_val and str(t_val).strip().upper() == "D":
                    continue

                md_rows.append(r)

        steps_count = len(steps_md_list)
        fcr_count = len(md_rows)

        # -------------------------------------------------
        # CASE 1 → MD in FCR but NOT in STEPS
        # -------------------------------------------------
        if fcr_count > 0 and steps_count == 0:
            sheet3.cell(row=parent_row, column=10).fill = red
            add_error_comment_to_cell(
                sheet3, parent_row, 10,
                "MD present in FCR but not in STEP DATA"
            )
            for r in md_rows:
                sheet3.cell(row=r, column=12).fill = red
                sheet3.cell(row=r, column=13).fill = red
            continue

        # -------------------------------------------------
        # CASE 2 → MD in STEPS but NOT in FCR
        # -------------------------------------------------
        if steps_count > 0 and fcr_count == 0:
            sheet3.cell(row=parent_row, column=10).fill = red
            add_error_comment_to_cell(
                sheet3, parent_row, 10,
                "Missing MD in FCR"
            )
            continue

        # -------------------------------------------------
        # CASE 3 → Count mismatch
        # -------------------------------------------------
        if steps_count != fcr_count:
            sheet3.cell(row=parent_row, column=10).fill = red
            add_error_comment_to_cell(
                sheet3, parent_row, 10,
                f"MD count mismatch: expected {steps_count}, got {fcr_count}"
            )

            from collections import Counter
            expected_counter = Counter([d.lower() for d in steps_md_list])
            used_counter = Counter()

            for r in md_rows:
                nomen = sheet3.cell(row=r, column=12).value
                norm = (nomen or "").strip().lower()

                if expected_counter[norm] > used_counter[norm]:
                    sheet3.cell(row=r, column=12).fill = green
                    sheet3.cell(row=r, column=13).fill = green
                    used_counter[norm] += 1
                else:
                    sheet3.cell(row=r, column=12).fill = red
                    sheet3.cell(row=r, column=13).fill = red
            continue   # ✅ FIXED position

        # -------------------------------------------------
        # CASE 4 → Description matching (unordered, duplicates safe)
        # -------------------------------------------------
        from collections import Counter
        expected_counter = Counter([d.lower() for d in steps_md_list])
        used_counter = Counter()

        mismatch_found = False

        for r in md_rows:
            row_obj = sheet3[r]
            nomen = row_obj[11].value.strip().lower() if row_obj[11].value else ""

            if expected_counter[nomen] > used_counter[nomen]:
                row_obj[11].fill = green
                row_obj[12].fill = green
                used_counter[nomen] += 1
            else:
                row_obj[11].fill = red
                row_obj[12].fill = red
                mismatch_found = True

        if mismatch_found:
            sheet3.cell(row=parent_row, column=10).fill = red
            add_error_comment_to_cell(
                sheet3, parent_row, 10,
                "MD description unmatched"
            )
