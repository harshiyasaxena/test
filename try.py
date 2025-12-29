def read_FCR_parents(ws_report, IN_COL, PART_COL):
    """
    Extract parents ONLY from FCR using IN hierarchy.
    Rule:
    - IN == 1 AND Part No present => parent
    """
    parents = []
    last_at_level = {}

    for r_idx in range(2, ws_report.max_row + 1):
        in_val = ws_report.cell(row=r_idx, column=IN_COL).value
        part = ws_report.cell(row=r_idx, column=PART_COL).value

        if not in_val or not part:
            continue

        try:
            level = int(in_val)
        except:
            continue

        part_u = normalize_spaces(str(part)).upper()

        # Track last seen at this level
        last_at_level[level] = part_u

        # Clear deeper levels if IN goes backward
        for l in list(last_at_level.keys()):
            if l > level:
                del last_at_level[l]

        # IN == 1 => DEFINITE parent
        if level == 1:
            parents.append(part_u)
            print(f"[FCR PARENT] row={r_idx} parent='{part_u}'")

        # Debug hierarchy walk (very useful)
        print(
            f"[FCR ITER] row={r_idx} IN={level} part={part_u} "
            f"parent={last_at_level.get(1)} "
            f"lvl2={last_at_level.get(2)} "
            f"lvl3={last_at_level.get(3)}"
        )

    print("\n[FINAL FCR PARENTS]")
    for i, p in enumerate(parents, 1):
        print(f"{i}. {p}")
    print()

    return parents
