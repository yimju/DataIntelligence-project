# -*- coding: utf-8 -*-
"""Dump FULL headers (all column names) of every file to a UTF-8 report."""
import io, os, csv

DATA = r"d:\data_intelligence_2\data"
OUT  = r"d:\data_intelligence_2\analysis\headers.md"
os.makedirs(os.path.dirname(OUT), exist_ok=True)

files = [
    "table_map.csv", "column_map.csv", "table_summary_map.csv",
    "column_summary_map.csv", "combined_pair_map.csv", "table_pair_map.csv",
    "column_pair_map.csv", "distinct_value.csv",
    "column_value_tree.csv", "table_value_tree.csv", "table_str_tree.csv",
]

with io.open(OUT, "w", encoding="utf-8") as out:
    for fn in files:
        path = os.path.join(DATA, fn)
        with io.open(path, "r", encoding="utf-8-sig", newline="") as f:
            header = f.readline().rstrip("\r\n")
            # parse with csv to respect quotes
            cols = next(csv.reader([header]))
            # also grab one data row
            datarow = f.readline().rstrip("\r\n")
            dvals = next(csv.reader([datarow])) if datarow else []
        out.write(f"\n## {fn}  ({len(cols)} columns)\n\n")
        for i, c in enumerate(cols):
            sample = dvals[i] if i < len(dvals) else ""
            if len(sample) > 60:
                sample = sample[:60] + "…"
            out.write(f"  {i:3d}. {c:<32} | e.g. {sample}\n")

print("WROTE", OUT)
