# pip install pdfplumber pandas openpyxl
import pdfplumber
import pandas as pd
import re

HEADER_NAMES = ["Added","Revised","T","A","AP","PID","Keyword","Nomenclature","DSI","UPA","LC"]

def _group_words_into_lines(words, y_tol=3.0):
    """Cluster words into lines by Y position."""
    words = sorted(words, key=lambda w: (w["top"], w["x0"]))
    lines, cur, cur_top = [], [], None
    for w in words:
        if cur_top is None or abs(w["top"] - cur_top) <= y_tol:
            cur.append(w)
            cur_top = w["top"] if cur_top is None else cur_top
        else:
            lines.append(sorted(cur, key=lambda x: x["x0"]))
            cur, cur_top = [w], w["top"]
    if cur:
        lines.append(sorted(cur, key=lambda x: x["x0"]))
    return lines

def _find_header_positions(words):
    """
    Find the header line and compute column x-boundaries
    from the header word boxes.
    Returns: (header_y, columns[ {name, x0, x1} ... ])
    """
    # Pick words that exactly match header tokens (case-insensitive)
    header_tokens = {h.lower(): h for h in HEADER_NAMES}
    # group words into lines then look for a line containing all headers (order not required)
    for line in _group_words_into_lines(words):
        texts = [w["text"].strip() for w in line]
        found = {}
        for w in line:
            t = w["text"].strip().lower()
            if t in header_tokens:
                found[header_tokens[t]] = w
        if all(h in found for h in HEADER_NAMES):
            # sort by x0 to establish column order
            ordered = [found[h] for h in HEADER_NAMES]
            ordered = sorted(ordered, key=lambda w: w["x0"])
            # build x boundaries halfway between neighbors
            columns = []
            for i, w in enumerate(ordered):
                if i == 0:
                    left = 0
                else:
                    left = (ordered[i-1]["x1"] + w["x0"]) / 2.0
                if i == len(ordered) - 1:
                    right = float("inf")
                else:
                    right = (w["x1"] + ordered[i+1]["x0"]) / 2.0
                columns.append({"name": HEADER_NAMES[i], "x0": left, "x1": right})
            header_y = min(w["top"] for w in ordered)
            return header_y, columns
    return None, None

def _assign_line_to_row(line_words, columns):
    """Assign words in one PDF line into the known columns by x position."""
    buckets = {c["name"]: [] for c in columns}
    for w in line_words:
        xmid = (w["x0"] + w["x1"]) / 2.0
        for c in columns:
            if c["x0"] <= xmid < c["x1"]:
                buckets[c["name"]].append(w["text"])
                break
    # join tokens by space
    return {k: " ".join(v).strip() for k, v in buckets.items()}

def _is_code_like(nom: str) -> bool:
    """
    Treat as code if it has NO spaces and at least two dashes, e.g. 25-23-00-01F.
    (You can tweak if your data needs stricter/looser rules.)
    """
    s = (nom or "").strip()
    if not s:
        return False
    return (" " not in s) and (s.count("-") >= 2) and bool(re.match(r"^[0-9A-Z\-]+$", s))

def pdf_to_excel(pdf_file="sample.pdf", output_file="output.xlsx"):
    with pdfplumber.open(pdf_file) as pdf:
        # Collect top text (exact) and words (for geometry-based table parsing)
        all_top_lines = []
        table_rows = []
        for page in pdf.pages:
            page_text = page.extract_text() or ""
            text_lines = [ln for ln in page_text.split("\n") if ln.strip()]
            words = page.extract_words(
                keep_blank_chars=False,
                use_text_flow=True,
                x_tolerance=2,
                y_tolerance=3,
            )

            header_y, columns = _find_header_positions(words)
            if not columns:
                # No table on this page; all lines are part of the "top" section
                all_top_lines.extend(text_lines)
                continue

            # Split top lines (above header) vs the rest, preserving EXACT top lines
            # We can safely keep only the lines above the first header we find.
            for ln in text_lines:
                all_top_lines.append(ln)
                if all(h in ln for h in ["Added", "Revised", "Nomenclature", "DSI"]):
                    # stop adding further lines from this page as top info
                    break

            # Build rows from lines below header_y
            line_groups = _group_words_into_lines([w for w in words if w["top"] > header_y + 0.1])
            for line in line_groups:
                row = _assign_line_to_row(line, columns)
                # Skip empty lines
                if any(v for v in row.values()):
                    table_rows.append(row)

    if not table_rows:
        raise RuntimeError("Couldn't parse any table rows from the PDF.")

    df = pd.DataFrame(table_rows)

    # Guarantee all expected columns exist
    for col in HEADER_NAMES:
        if col not in df.columns:
            df[col] = ""

    # Normalize/clean
    df["DSI"] = df["DSI"].astype(str).str.strip().str.upper()
    df["Nomenclature"] = (
        df["Nomenclature"]
        .astype(str)
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
    )

    # Merge NOMENCLATURE sentences for same DSI (only when it's "words", not code-like)
    merged = []
    first_word_row_idx_for_dsi = {}  # DSI -> index in 'merged' list for word-type rows

    for _, row in df.iterrows():
        dsi = row["DSI"]
        nom = row["Nomenclature"]
        if nom and not _is_code_like(nom):
            # words/sentence → merge into first row of this DSI
            if dsi in first_word_row_idx_for_dsi:
                i = first_word_row_idx_for_dsi[dsi]
                if nom:
                    # add a space before appending to build a single sentence
                    merged[i]["Nomenclature"] = (merged[i]["Nomenclature"] + " " + nom).strip()
            else:
                merged.append(row.to_dict())
                first_word_row_idx_for_dsi[dsi] = len(merged) - 1
        else:
            # code-like or empty nomenclature → keep as its own row (no merging)
            merged.append(row.to_dict())

    final_df = pd.DataFrame(merged)

    # Write to Excel: top info first (as-is), then a blank row, then headers + table
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        # Top lines (exactly as found)
        pd.DataFrame(all_top_lines, columns=["Report Info"]).to_excel(
            writer, sheet_name="Sheet1", index=False, startrow=0
        )
        start = len(all_top_lines) + 2
        # Reorder columns to the expected order
        final_df = final_df[HEADER_NAMES]
        final_df.to_excel(writer, sheet_name="Sheet1", index=False, startrow=start)

    print(f"✅ Excel saved to {output_file}")

if __name__ == "__main__":
    pdf_to_excel("sample.pdf", "output.xlsx")
