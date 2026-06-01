# -*- coding: utf-8 -*-
"""Profile every CSV in data/: encoding, first raw lines, column count, row count (small files)."""
import os, sys, io

DATA = r"d:\data_intelligence_2\data"
files = [
    "table_map_ver.csv", "table_value_tree.csv", "table_str_tree.csv",
    "column_value_tree.csv", "table_map.csv", "table_summary_map.csv",
    "column_map.csv", "column_summary_map.csv",
    "combined_pair_map.csv", "table_pair_map.csv",
    "distinct_value.csv", "column_pair_map.csv",
]

ENCODINGS = ["utf-8-sig", "utf-8", "cp949", "euc-kr", "latin-1"]

def detect_encoding(path, nbytes=200000):
    with open(path, "rb") as f:
        raw = f.read(nbytes)
    for enc in ENCODINGS:
        try:
            raw.decode(enc)
            return enc
        except UnicodeDecodeError:
            continue
    return "latin-1"

def head_lines(path, enc, n=4):
    out = []
    with io.open(path, "r", encoding=enc, errors="replace", newline="") as f:
        for i, line in enumerate(f):
            if i >= n:
                break
            out.append(line.rstrip("\r\n"))
    return out

SMALL = 50 * 1024 * 1024  # 50MB -> count rows

for fn in files:
    path = os.path.join(DATA, fn)
    size = os.path.getsize(path)
    enc = detect_encoding(path)
    print("=" * 90)
    print(f"FILE: {fn}   size={size/1e6:.1f} MB   enc={enc}")
    lines = head_lines(path, enc, 4)
    # column count from first line: try tab then comma
    first = lines[0] if lines else ""
    ncols_tab = first.count("\t") + 1
    ncols_comma = first.count(",") + 1
    print(f"  cols(by tab)={ncols_tab}  cols(by comma)={ncols_comma}")
    for i, ln in enumerate(lines):
        disp = ln if len(ln) <= 400 else ln[:400] + " ...[truncated]"
        print(f"  L{i}: {disp}")
    if size < SMALL:
        with open(path, "rb") as f:
            rows = sum(1 for _ in f)
        print(f"  ROW COUNT: {rows}")
    sys.stdout.flush()
