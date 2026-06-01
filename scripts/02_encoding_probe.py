# -*- coding: utf-8 -*-
"""Find the true encoding by decoding a Korean-bearing line under several codecs."""
import io

path = r"d:\data_intelligence_2\data\table_map.csv"

# read raw bytes of the 2nd line (first data row, contains Korean)
with open(path, "rb") as f:
    f.readline()           # header
    raw = f.readline()     # first data row

print("raw bytes (first 80):", raw[:80])
print()
for enc in ["utf-8-sig", "utf-8", "cp949", "euc-kr"]:
    try:
        s = raw.decode(enc)
        print(f"[{enc:10}] OK : {s[:120]!r}")
    except Exception as e:
        print(f"[{enc:10}] ERR: {e}")

# Also test on column_map values
print("\n--- distinct_value.csv line 2 ---")
with open(r"d:\data_intelligence_2\data\distinct_value.csv", "rb") as f:
    f.readline()
    raw2 = f.readline()
for enc in ["utf-8", "cp949"]:
    try:
        print(f"[{enc}] {raw2.decode(enc)[:120]!r}")
    except Exception as e:
        print(f"[{enc}] ERR {e}")
