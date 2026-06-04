#!/usr/bin/env python3
"""
Extract TRREB Market Watch "Summary of Existing Home Transactions" tables to CSV.

FORMAT: TRREB's 2025+ report layout (BI-generated, with sparkline 'Abc' cells and
an overprinted ghost of the previous row fused into each row). For the older
pre-2025 layout, use the prior row-text version of this script.

  page 3 -> ALL TRREB AREAS (regions + sub-municipalities)
  page 4 -> CITY OF TORONTO municipal breakdown (W/C/E districts)

Approach (char-level):
  * Cluster characters into visual rows by exact vertical position.
  * Split name vs the 11 value columns by x-position bins.
  * Read each cell's chars in PDF stream order, which places the ghost copy of the
    previous row first and the real value second; de-merge by stripping the
    previous row's value as a prefix.
  * Strip 'Abc' sparkline letters from numeric cells.

Usage:  python3 extract_trreb.py [in.pdf] [out_dir] [-v]
Install: pip install pdfplumber
"""

import csv
import re
import sys
from collections import defaultdict
from pathlib import Path

import pdfplumber

PAGES = [3, 4]
KEEP_RAW = False

COLUMNS = [
    "Area", "Sales", "Dollar Volume", "Average Price", "Median Price",
    "New Listings", "SNLR Trend %", "Active Listings", "Months Inv Trend",
    "Avg SP/LP %", "Avg LDOM", "Avg PDOM",
]
N_NUM = len(COLUMNS) - 1
PAGE_LABELS = {3: "all_trreb_areas", 4: "toronto_breakdown"}

# x-bin left edges for the 11 numeric columns; name is everything left of bin 0
BIN_EDGES = [92, 130, 208, 270, 335, 400, 455, 520, 600, 660, 725, 800]
TOP_MIN, TOP_MAX = 92, 605      # data band (skip title/footer)
ROW_TOL = 0.5                    # vertical px tolerance to group chars into a row

NUM_RE = re.compile(r"^\$?-?[\d,]+(?:\.\d+)?%?$")
LETTERS = re.compile(r"[A-Za-z]")


def col_of(xc: float):
    if xc < BIN_EDGES[0]:
        return -1                # name zone
    for i in range(N_NUM):
        if BIN_EDGES[i] <= xc < BIN_EDGES[i + 1]:
            return i
    return N_NUM - 1


def cluster_rows(chars):
    """Return ordered list of row-clusters; each is a dict col -> stream-ordered text."""
    tops = sorted({round(c["top"], 2) for c in chars})
    groups, cur = [], []
    for t in tops:
        if cur and t - cur[0] > ROW_TOL:
            groups.append(cur)
            cur = []
        cur.append(t)
    if cur:
        groups.append(cur)
    top2cid = {t: i for i, g in enumerate(groups) for t in g}

    cells = defaultdict(lambda: defaultdict(list))   # cid -> col -> [text] (stream order)
    for c in chars:                                  # pg.chars is in stream order
        cid = top2cid[round(c["top"], 2)]
        cells[cid][col_of((c["x0"] + c["x1"]) / 2)].append(c["text"])
    return [cells[i] for i in range(len(groups))]


def cell_text(cell, col):
    return "".join(cell.get(col, []))


def norm_name(raw):
    return re.sub(r"\s+", " ", raw.replace("\t", " ")).strip()


def to_number(s):
    if s in ("", "-"):
        return None
    core = s.replace("$", "").replace(",", "").rstrip("%")
    try:
        return float(core) if "." in core else int(core)
    except ValueError:
        return s


def parse_page(pdf, page_no, keep_raw, verbose=False):
    pg = pdf.pages[page_no - 1]
    chars = [c for c in pg.chars if TOP_MIN <= c["top"] <= TOP_MAX]
    raw_rows = []
    for cell in cluster_rows(chars):
        name = norm_name(cell_text(cell, -1))
        if not name:
            continue                       # nameless ghost / Abc-only line
        vals = [LETTERS.sub("", cell_text(cell, c)) for c in range(N_NUM)]
        raw_rows.append((name, vals))

    out, skipped = [], []
    prev_name, prev_fmt = None, None
    for name, vals in raw_rows:
        if prev_name and name == prev_name:           # offset ghost duplicate line
            skipped.append((name, "dup"))
            continue
        merged = bool(prev_name and name.startswith(prev_name + " "))
        real_name = name[len(prev_name):].strip() if merged else name

        fmt = []
        for i, v in enumerate(vals):
            if merged and prev_fmt and prev_fmt[i] and v.startswith(prev_fmt[i]):
                v = v[len(prev_fmt[i]):]
            elif merged and prev_fmt and prev_fmt[i] and len(v) > len(prev_fmt[i]):
                v = v[len(prev_fmt[i]):]              # length fallback
            fmt.append(v)

        if not (all(fmt) and all(NUM_RE.match(x) for x in fmt)):
            if any(fmt):
                skipped.append((real_name, "unparsed:" + "|".join(fmt)))
            # advance prev only when this looked like a real (de-mergeable) data row
            if any(fmt):
                prev_name, prev_fmt = real_name, fmt
            continue

        values = fmt if keep_raw else [to_number(x) for x in fmt]
        out.append([real_name] + values)
        prev_name, prev_fmt = real_name, fmt

    if verbose:
        for nm, why in skipped:
            print(f"    [skip {why[:40]}] {nm!r}")
    return out


def main():
    pdf_path = Path(sys.argv[1]) if len(sys.argv) > 1 and not sys.argv[1].startswith("-") else Path("/mnt/user-data/uploads/mw2605.pdf")
    out_dir = Path(sys.argv[2]) if len(sys.argv) > 2 and not sys.argv[2].startswith("-") else Path("/mnt/user-data/outputs")
    out_dir.mkdir(parents=True, exist_ok=True)
    verbose = "-v" in sys.argv
    stem = pdf_path.stem

    written = []
    with pdfplumber.open(pdf_path) as pdf:
        for pg in PAGES:
            rows = parse_page(pdf, pg, KEEP_RAW, verbose)
            out_csv = out_dir / f"{stem}_p{pg}_{PAGE_LABELS.get(pg, f'page{pg}')}.csv"
            with open(out_csv, "w", newline="", encoding="utf-8") as f:
                w = csv.writer(f)
                w.writerow(COLUMNS)
                w.writerows(rows)
            written.append((pg, out_csv, len(rows)))
    for pg, path, n in written:
        print(f"page {pg}: {n} rows -> {path}")


if __name__ == "__main__":
    main()
