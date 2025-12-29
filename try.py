def read_FCR_parents(ws_report, IN_COL, PART_COL):
    """
    Parents are derived ONLY from FCR based on IN hierarchy.
    IN=1 with Part No => parent.
    """
    parents = []
    last_at_level = {}

    for r in range(2, ws_report.max_row + 1):
        in_val = ws_report.cell(row=r, column=IN_COL).value
        part = ws_report.cell(row=r, column=PART_COL).value

        if not in_val or not part:
            continue

        try:
            level = int(in_val)
        except:
            continue

        part_u = normalize_spaces(str(part)).upper()

        # store last seen at this level
        last_at_level[level] = part_u

        # clear deeper levels if IN goes backward
        for l in list(last_at_level.keys()):
            if l > level:
                del last_at_level[l]

        # IN == 1 => definite parent
        if level == 1:
            parents.append(part_u)
            print(f"[FCR PARENT FOUND] row={r} parent={part_u}")

        # optional debug for hierarchy tracking
        print(
            f"[FCR ITER] row={r} IN={level} part={part_u} "
            f"parent={last_at_level.get(1)} "
            f"lvl2={last_at_level.get(2)} "
            f"lvl3={last_at_level.get(3)}"
        )

    # ✅ FINAL SUMMARY PRINT
    print("\n========== FINAL PARENTS FROM FCR ==========")
    for i, p in enumerate(parents, start=1):
        print(f"{i}. {p}")
    print("===========================================\n")

    return parents
