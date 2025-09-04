import os.path
import pdfplumber
import pandas as pd
import re
import string
import tkinter as tk
from tkinter import Tk, filedialog

# ------------------- Headers per Report -------------------
# Edit this to match your exact headers for each report type.
# If the PDF uses "Necessory Report" spelling, keep that key as is.
REPORT_HEADERS = {
    "Mandatory Report": [
        "Dates", "Verified", "Code", "D", "A", "LK", "ML", "MID", "JK",
        "Fish Ko", "Movie", "Nomenclature", "DSI", "LK", "BN", "OP"
    ],
    # You can use either spelling in the PDF; detection handles both.
    "Necessary Report": [
        # <<< Replace with the exact headers for the Necessary report >>>
        "Date", "Clock", "Different Col1", "Different Col2", "Nomenclature", "DSI"
    ],
    "Necessory Report": [
        # If your PDF literally says "Necessory Report", set headers here too (or reuse the same)
        "Date", "Clock", "Different Col1", "Different Col2", "Nomenclature", "DSI"
    ]
}

# Normalize report title variants in PDFs to a consistent sheet/report name
REPORT_NAME_ALIASES = {
    "mandatory report": "Mandatory Report",
    "necessary report": "Necessary Report",
    "necessory report": "Necessory Report",
}

# ------------------- Helpers -------------------

def _group_words_into_lines(words, y_tol=3.0):
    words = sorted(words, key=lambda w: (w.get("top", 0), w.get("x0", 0)))
    lines, cur, cur_top = [], [], None
    for w in words:
        top = w.get("top", 0)
        if cur_top is None or abs(top - cur_top) <= y_tol:
            cur.append(w)
            cur_top = top if cur_top is None else cur_top
        else:
            lines.append(sorted(cur, key=lambda x: x.get("x0", 0)))
            cur, cur_top = [w], top
    if cur:
        lines.append(sorted(cur, key=lambda x: x.get("x0", 0)))
    return lines

def _normalize(t):
    if t is None:
        return ""
    s = " ".join(str(t).split())
    return s.strip().strip(string.punctuation).lower()

def _find_header_positions(words, headers):
    """
    Find header words' x-positions and compute column bands.
    """
    if not headers:
        return None, None

    header_map = {h.lower(): h for h in headers}
    # Use a few likely anchors to shortlist potential header lines
    anchors = {h.lower() for h in headers[:4]} if len(headers) >= 4 else {h.lower() for h in headers}

    lines = _group_words_into_lines(words)
    for line in lines:
        toks = []
        for w in line:
            norm = _normalize(w.get("text", ""))
            toks.append({"box": w, "norm": norm})

        combined = " ".join(t["norm"] for t in toks if t["norm"])
        if not any(a in combined for a in anchors):
            continue

        found = {}
        # Exact single-token header matches
        for t in toks:
            if t["norm"] in header_map:
                found[header_map[t["norm"]]] = t["box"]

        # Two-token header matches (e.g., "Fish Ko")
        for i in range(len(toks) - 1):
            joined = toks[i]["norm"] + " " + toks[i + 1]["norm"]
            if joined in header_map:
                w0 = toks[i]["box"]
                w1 = toks[i + 1]["box"]
                synthetic = {
                    "text": header_map[joined],
                    "x0": w0.get("x0", 0),
                    "x1": w1.get("x1", w1.get("x0", 0)),
                    "top": min(w0.get("top", 0), w1.get("top", 0)),
                }
                found[header_map[joined]] = synthetic

        if not found:
            continue

        ordered_words = sorted(found.values(), key=lambda w: w.get("x0", 0))
        columns = []
        for i, w in enumerate(ordered_words):
            x0 = w.get("x0", 0)
            x1 = w.get("x1", x0 + 1)
            left = 0.0 if i == 0 else (ordered_words[i - 1].get("x1", 0) + x0) / 2.0
            right = float("inf") if i == len(ordered_words) - 1 else (x1 + ordered_words[i + 1].get("x0", 0)) / 2.0
            name = header_map.get(_normalize(w.get("text", "")), w.get("text", "").strip())
            columns.append({"name": name, "x0": left, "x1": right})

        header_y = min(w.get("top", 0) for w in ordered_words)
        return header_y, columns

    return None, None

def _assign_line_to_row(line_words, columns):
    """Assign words in a PDF line to known columns by x position."""
    buckets = {c["name"]: [] for c in columns}
    for w in line_words:
        wx0 = w.get("x0", 0)
        wx1 = w.get("x1", 0)

        # Prefer left boundary containment
        left_col = None
        for c in columns:
            if c["x0"] <= wx0 < c["x1"]:
                left_col = c
                break
        if left_col:
            buckets[left_col["name"]].append(w.get("text", ""))
            continue

        # Otherwise, choose the column with maximum overlap
        best_col = None
        best_overlap = 0.0
        for c in columns:
            cx0, cx1 = c["x0"], c["x1"]
            overlap = max(0.0, min(wx1, cx1) - max(wx0, cx0))
            if overlap > best_overlap:
                best_overlap = overlap
                best_col = c
        if best_col and best_overlap > 0:
            buckets[best_col["name"]].append(w.get("text", ""))
            continue

        # Fallback: midpoint containment
        xmid = (wx0 + wx1) / 2.0
        for c in columns:
            if c["x0"] <= xmid < c["x1"]:
                buckets[c["name"]].append(w.get("text", ""))
                break

    row = {k: " ".join(v).strip() for k, v in buckets.items()}
    return row

valid_dsi_codes = {"AB", "CD", "EF"}

def _fix_spill_into_dsi(row: dict) -> dict:
    """
    If DSI is empty but last token of Nomenclature looks like a DSI code, move it.
    If DSI exists but is invalid, append to Nomenclature and clear DSI.
    """
    dsi = (row.get("DSI") or "").strip().upper()
    nom = (row.get("Nomenclature") or "").strip()

    if not nom and not dsi:
        return row

    if not dsi and nom:
        parts = nom.split()
        if parts and parts[-1].upper() in valid_dsi_codes:
            row["DSI"] = parts[-1].upper()
            row["Nomenclature"] = " ".join(parts[:-1]).strip()
        return row

    if dsi and dsi not in valid_dsi_codes and nom:
        row["Nomenclature"] = (nom + " " + dsi).strip()
        row["DSI"] = ""
    return row

def merge_consecutive_rows(df: pd.DataFrame) -> pd.DataFrame:
    """
    Merge rows when DSI is FI/FP and appear consecutively.
    Only runs if both 'DSI' and 'Nomenclature' exist.
    """
    if "DSI" not in df.columns or "Nomenclature" not in df.columns:
        return df

    merged_rows = []
    i = 0
    while i < len(df):
        row = df.iloc[i].copy()
        dsi = str(row["DSI"]).strip().upper()
        if dsi in ("FI", "FP"):
            j = i + 1
            merged_nom = str(row["Nomenclature"])
            while j < len(df) and str(df.iloc[j]["DSI"]).strip().upper() == dsi:
                merged_nom += " " + str(df.iloc[j]["Nomenclature"])
                j += 1
            row["Nomenclature"] = merged_nom.strip()
            merged_rows.append(row)
            i = j
        else:
            merged_rows.append(row)
            i += 1
    return pd.DataFrame(merged_rows, columns=df.columns)

def _row_is_header_like(r: dict, headers: list) -> bool:
    """
    Detect rows that are actually header repeats (to drop them).
    """
    non_empty = [str(v).strip().lower() for v in r.values() if v and str(v).strip()]
    if not non_empty:
        return False
    header_set = {h.lower() for h in headers}
    header_hits = sum(1 for v in non_empty if v in header_set)
    # Consider it header-like if all non-empty tokens are from the header set
    return header_hits == len(non_empty)

def _normalize_report_name(line: str) -> str | None:
    s = (line or "").strip().lower()
    for k, v in REPORT_NAME_ALIASES.items():
        if k in s:
            return v  # Canonical name for sheet & headers
    return None

# ------------------- Main -------------------

def pdf_to_excel(pdf_file="sample.pdf", output_file="output.xlsx"):
    # We'll accumulate rows per report across the whole PDF
    reports_data = {}  # {report_name: {"rows": [], "top_lines": []}}
    current_report = None
    current_rows = []
    current_top_lines = []

    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text() or ""
            text_lines = [ln for ln in (page_text.split("\n")) if ln and ln.strip()]
            words = page.extract_words(
                keep_blank_chars=False,
                use_text_flow=True,
                x_tolerance=2,
                y_tolerance=3,
            )

            # Detect report transitions
            detected = None
            for ln in text_lines:
                rep = _normalize_report_name(ln)
                if rep:
                    detected = rep
                    break

            if detected:
                # Flush previous report data
                if current_report is not None:
                    if current_rows:
                        bucket = reports_data.setdefault(current_report, {"rows": [], "top_lines": []})
                        bucket["rows"].extend(current_rows)
                        bucket["top_lines"].extend(current_top_lines)
                # Start new report
                current_report = detected
                current_rows = []
                current_top_lines = []

            # If we haven't seen any report header yet, skip parsing
            if not current_report:
                continue

            headers = REPORT_HEADERS.get(current_report, [])
            if not headers:
                # If headers for this report aren't defined, skip parsing its table
                # but keep collecting top lines (for context)
                current_top_lines.extend(text_lines)
                continue

            header_y, columns = _find_header_positions(words, headers)
            if header_y is None or not columns:
                # Still in the heading/context zone; stash lines
                current_top_lines.extend(text_lines)
                continue

            # Capture lines on the page for the sheet header context until a probable header line
            for line_text in text_lines:
                current_top_lines.append(line_text)
                # crude early stop once we hit something that looks like headers
                if all(h.lower() in line_text.lower() for h in headers[:2]):
                    break

            # Collect table rows below header line
            line_groups = _group_words_into_lines([w for w in words if w.get("top", 0) > header_y + 0.1])
            for line in line_groups:
                row = _assign_line_to_row(line, columns)
                row = _fix_spill_into_dsi(row)
                if any((str(v).strip() if v is not None else "") for v in row.values()):
                    current_rows.append(row)

        # Flush the last report after finishing all pages
        if current_report is not None and current_rows:
            bucket = reports_data.setdefault(current_report, {"rows": [], "top_lines": []})
            bucket["rows"].extend(current_rows)
            bucket["top_lines"].extend(current_top_lines)

    if not reports_data:
        print("No tables found")
        return

    # Define desired sheet order: Mandatory first, then Necessary/Necessory, then others
    preferred_order = ["Mandatory Report", "Necessary Report", "Necessory Report"]
    ordered_report_names = [r for r in preferred_order if r in reports_data] + \
                           [r for r in reports_data.keys() if r not in preferred_order]

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        for report_name in ordered_report_names:
            data = reports_data[report_name]
            headers = REPORT_HEADERS.get(report_name, [])
            rows = data["rows"]
            top_lines = data["top_lines"]

            # Drop header-like rows before DataFrame build (safer to do after dict creation)
            filtered_rows = [r for r in rows if not _row_is_header_like(r, headers)]
            df = pd.DataFrame(filtered_rows)

            # Ensure all expected headers exist as columns
            for col in headers:
                if col not in df.columns:
                    df[col] = ""

            # Reindex to enforce column order and fill empties
            df = df.reindex(columns=headers, fill_value="")

            # Normalize common columns if present
            if "DSI" in df.columns:
                df["DSI"] = df["DSI"].astype(str).str.strip().str.upper()
            if "Nomenclature" in df.columns:
                df["Nomenclature"] = (
                    df["Nomenclature"].astype(str).str.replace(r"\s+", " ", regex=True).str.strip()
                )

            # Merge FI/FP blocks if applicable
            final_df = merge_consecutive_rows(df)

            # Clean: turn '---' etc into empty
            final_df = final_df.replace(r'^\-+$', '', regex=True)

            # Write the table
            if final_df.empty:
                pd.DataFrame(columns=headers).to_excel(writer, sheet_name=report_name, index=False, startrow=0)
                start_row = 0
            else:
                final_df.to_excel(writer, sheet_name=report_name, index=False, startrow=0)
                start_row = len(final_df)

            # Append report info below the table (deduped)
            if top_lines:
                seen = set()
                dedup_top_lines = []
                for ln in top_lines:
                    key = " ".join(ln.split()).strip()
                    if key and key not in seen:
                        dedup_top_lines.append(ln)
                        seen.add(key)
                if dedup_top_lines:
                    info_df = pd.DataFrame(dedup_top_lines, columns=["Report Info"])
                    info_df.to_excel(writer, sheet_name=report_name, index=False, startrow=start_row + 3)

    print(f"Excel saved to {output_file}")

# ------------------- Run -------------------

if __name__ == "__main__":
    Tk().withdraw()
    pdf_path = filedialog.askopenfilename(title="Select PDF file", filetypes=[("PDF files", "*.pdf")])
    if not pdf_path:
        print("No PDF selected")
    else:
        out = pdf_path.rsplit(".", 1)[0] + ".xlsx"
        pdf_to_excel(pdf_path, out)
        print(f"Saved Excel to: {out}")
