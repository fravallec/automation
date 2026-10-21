#!/usr/bin/env python3
"""
internal_linker.py
------------------
Automatically add internal links to a Word (.docx) document from a simple
Excel map.

Excel format (first sheet):
    Column A = the word / anchor text to link
    Column B = the URL to link it to
A header row (e.g. "Keyword" | "URL") is detected and skipped automatically.

Rules:
    * Each word is linked ONLY on its first match, and never more than once.
    * Matching is whole-word and case-insensitive by default.
    * Text that is already a hyperlink is left untouched (never double-linked).
    * Original run formatting (bold, size, etc.) is preserved.
    * When several words could match at the same spot, the longest one wins
      (so "data analysis" is linked rather than just "data").

Usage:
    python internal_linker.py --excel map.xlsx --docx article.docx
    python internal_linker.py -e map.xlsx -d article.docx -o article_linked.docx
    python internal_linker.py -e map.xlsx -d article.docx --match-case --dry-run

Requirements:
    pip install python-docx openpyxl

Author: Francesco Vallecoccia  —  shared for the SEO community. MIT licence.
"""

import argparse
import copy
import re
import sys
from pathlib import Path

try:
    import openpyxl
except ImportError:
    sys.exit("Missing dependency: pip install openpyxl")

try:
    from docx import Document
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.opc.constants import RELATIONSHIP_TYPE as RT
except ImportError:
    sys.exit("Missing dependency: pip install python-docx")


HEADER_WORDS = {"keyword", "word", "anchor", "term", "phrase"}
HEADER_URLS = {"url", "link", "href", "destination", "target"}


# --------------------------------------------------------------------------- #
# Load the Excel word -> url map
# --------------------------------------------------------------------------- #
def load_mapping(path, match_case=False):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    rows = [r for r in ws.iter_rows(values_only=True) if r and any(r)]
    if not rows:
        sys.exit(f"No data found in {path}")

    # Detect and drop a header row.
    first_a = str(rows[0][0]).strip().lower() if rows[0][0] else ""
    first_b = str(rows[0][1]).strip().lower() if len(rows[0]) > 1 and rows[0][1] else ""
    if first_a in HEADER_WORDS or first_b in HEADER_URLS:
        rows = rows[1:]

    entries = []
    seen = set()
    for i, row in enumerate(rows, start=1):
        word = (str(row[0]).strip() if row[0] is not None else "")
        url = (str(row[1]).strip() if len(row) > 1 and row[1] is not None else "")
        if not word or not url:
            print(f"  ! Skipping row {i}: missing word or URL ({word!r} -> {url!r})")
            continue
        key = word if match_case else word.lower()
        if key in seen:
            print(f"  ! Duplicate word '{word}' ignored (first one kept).")
            continue
        seen.add(key)
        flags = 0 if match_case else re.IGNORECASE
        pattern = re.compile(r"\b" + re.escape(word) + r"\b", flags | re.UNICODE)
        entries.append({"word": word, "url": url, "key": key, "pattern": pattern})
    if not entries:
        sys.exit("No valid word/URL pairs found.")
    return entries


# --------------------------------------------------------------------------- #
# Build a <w:hyperlink> element and split a run around the matched text
# --------------------------------------------------------------------------- #
def _make_hyperlink(paragraph, url, text, source_rpr, color, underline):
    part = paragraph.part
    r_id = part.relate_to(url, RT.HYPERLINK, is_external=True)

    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)

    new_run = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    if source_rpr is not None:
        # Copy existing formatting, but drop any colour/underline we will reset.
        for child in source_rpr:
            if child.tag in (qn("w:color"), qn("w:u")):
                continue
            rpr.append(copy.deepcopy(child))
    if color:
        c = OxmlElement("w:color")
        c.set(qn("w:val"), color)
        rpr.append(c)
    if underline:
        u = OxmlElement("w:u")
        u.set(qn("w:val"), "single")
        rpr.append(u)
    new_run.append(rpr)

    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = text
    new_run.append(t)
    hyperlink.append(new_run)
    return hyperlink


def _split_and_link(paragraph, run, start, end, url, color, underline):
    r = run._r
    source_rpr = r.find(qn("w:rPr"))
    full = run.text
    before, match_text, after = full[:start], full[start:end], full[end:]

    run.text = before  # keep original run (and formatting) as the 'before' chunk

    hyperlink = _make_hyperlink(paragraph, url, match_text, source_rpr, color, underline)
    r.addnext(hyperlink)

    if after:
        after_r = OxmlElement("w:r")
        if source_rpr is not None:
            after_r.append(copy.deepcopy(source_rpr))
        t = OxmlElement("w:t")
        t.set(qn("xml:space"), "preserve")
        t.text = after
        after_r.append(t)
        hyperlink.addnext(after_r)


# --------------------------------------------------------------------------- #
# Walk every paragraph (body + tables) in document order
# --------------------------------------------------------------------------- #
def iter_paragraphs(doc):
    for p in doc.paragraphs:
        yield p
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    yield p


def link_paragraph(paragraph, entries, used, color, underline):
    """Link the earliest unused keyword(s) inside one paragraph; longest wins on ties."""
    count = 0
    while True:
        best = None  # (run_index, start, end, entry)
        runs = paragraph.runs  # re-read each pass so new 'after' runs are scanned
        for ri, run in enumerate(runs):
            text = run.text
            if not text:
                continue
            for entry in entries:
                if entry["key"] in used:
                    continue
                m = entry["pattern"].search(text)
                if not m:
                    continue
                cand = (ri, m.start(), m.end(), entry)
                if best is None:
                    best = cand
                else:
                    # prefer smaller run index, then earlier start, then longer match
                    better = (cand[0], cand[1], -(cand[2] - cand[1]))
                    current = (best[0], best[1], -(best[2] - best[1]))
                    if better < current:
                        best = cand
        if best is None:
            break
        ri, start, end, entry = best
        _split_and_link(paragraph, paragraph.runs[ri], start, end, entry["url"], color, underline)
        used.add(entry["key"])
        count += 1
    return count


# --------------------------------------------------------------------------- #
# Main
# --------------------------------------------------------------------------- #
def main():
    ap = argparse.ArgumentParser(description="Add internal links to a .docx from an Excel word/URL map.")
    ap.add_argument("-e", "--excel", required=True, help="Path to the .xlsx map (col A = word, col B = URL).")
    ap.add_argument("-d", "--docx", required=True, help="Path to the Word .docx to process.")
    ap.add_argument("-o", "--output", help="Output path (default: <name>_linked.docx).")
    ap.add_argument("--match-case", action="store_true", help="Case-sensitive matching.")
    ap.add_argument("--color", default="0563C1", help="Link colour as hex, or 'none' to keep original. Default 0563C1.")
    ap.add_argument("--no-underline", action="store_true", help="Do not underline the links.")
    ap.add_argument("--dry-run", action="store_true", help="Report what would be linked without saving.")
    args = ap.parse_args()

    excel_path = Path(args.excel)
    docx_path = Path(args.docx)
    if not excel_path.exists():
        sys.exit(f"Excel file not found: {excel_path}")
    if not docx_path.exists():
        sys.exit(f"Word file not found: {docx_path}")

    color = None if args.color.lower() == "none" else args.color.lstrip("#")
    underline = not args.no_underline

    print(f"Loading map from {excel_path} ...")
    entries = load_mapping(excel_path, match_case=args.match_case)
    print(f"  {len(entries)} word/URL pairs loaded.\n")

    print(f"Processing {docx_path} ...")
    doc = Document(str(docx_path))
    used = set()
    total = 0
    for paragraph in iter_paragraphs(doc):
        if len(used) == len(entries):
            break  # every word already linked once
        total += link_paragraph(paragraph, entries, used, color, underline)

    # Report
    linked = [e["word"] for e in entries if e["key"] in used]
    missed = [e["word"] for e in entries if e["key"] not in used]
    print(f"\n  Linked {total} word(s):")
    for w in linked:
        print(f"    + {w}")
    if missed:
        print(f"\n  Not found in the document ({len(missed)}):")
        for w in missed:
            print(f"    - {w}")

    if args.dry_run:
        print("\nDry run — no file written.")
        return

    out_path = Path(args.output) if args.output else docx_path.with_name(docx_path.stem + "_linked.docx")
    doc.save(str(out_path))
    print(f"\nSaved: {out_path}")


if __name__ == "__main__":
    main()
