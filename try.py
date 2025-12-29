def read_FCR_parents(ws_report, IN_COL, PART_COL, debug=True):
    """
    Extract parents ONLY from FCR based on IN hierarchy rules.
    """

    last_at_level = {}   # level -> most recent valid part
    parents = set()

    prev_in = None

    for r_idx, row in enumerate(ws_report.iter_rows(min_row=2), start=2):
        in_val = row[IN_COL - 1].value
        part = row[PART_COL - 1].value

        if part:
            part = str(part).strip().upper()
        else:
            continue

        try:
            in_val = int(in_val)
        except Exception:
            continue

        # BAC rule (global)
        if part.startswith("BAC"):
            if debug:
                print(f"[SKIP BAC] row={r_idx} IN={in_val} part={part}")
            continue

        if debug:
            print(f"[ROW] row={r_idx} IN={in_val} part={part}")

        # Rule 1: IN == 1 → parent
        if in_val == 1:
            parents.add(part)
            if debug:
                print(f"  → PARENT (IN=1): {part}")

        # Rule 2: IN drop → new parent
        if prev_in is not None and in_val < prev_in:
            parents.add(part)
            if debug:
                print(f"  → PARENT (IN DROP {prev_in}->{in_val}): {part}")

        # Rule 3: attach to IN-1 → promote that parent
        parent_level = in_val - 1
        if parent_level in last_at_level:
            parent_candidate = last_at_level[parent_level]
            parents.add(parent_candidate)
            if debug:
                print(
                    f"  → CHILD of {parent_candidate} (IN={parent_level})"
                )

        # update last seen at this level
        last_at_level[in_val] = part
        prev_in = in_val

    if debug:
        print("\n[FINAL PARENTS]")
        for p in parents:
            print(" ", p)

    return parents
