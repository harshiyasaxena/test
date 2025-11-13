desc_list = PN_to_desc3[part_no3]
    qd_index = sum(
        1
        for r in range(2, row[0].row)
        if self.sheet3.cell(row=r, column=13).value
        and str(self.sheet3.cell(row=r, column=13).value).strip().upper() == "QD"
        and normalize_PN(self.sheet3.cell(row=r, column=10).value)
        == normalize_PN(part_no3)
    )

    if qd_index < len(desc_list):
        desc = desc_list[qd_index]
    else:
        desc = desc_list[-1]
        
