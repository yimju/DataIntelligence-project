# -*- coding: utf-8 -*-
"""(a) line counts of big files, (b) similarity-tree cluster sizes,
   (c) SAMPLED cross-domain shared-value scan on distinct_value.
   Cross-domain findings here are LOCAL (sample) -> to be globally verified later (rule #4)."""
import os, io, csv, collections, numpy as np, pandas as pd
DATA = r"d:\data_intelligence_2\data"
OUT  = r"d:\data_intelligence_2\analysis\scale_crossdomain.md"
buf=[]; w=lambda s="":buf.append(str(s))

def count_lines(path, bufsize=1<<22):
    n=0
    with open(path,"rb") as f:
        while True:
            b=f.read(bufsize)
            if not b: break
            n+=b.count(b"\n")
    return n

# ---- (a) scale ----
w("# 규모 & 교차도메인 스캔\n")
w("## (a) 대용량 파일 행 수 (헤더 포함 줄 수)")
w("| file | lines |")
w("|---|---:|")
for fn in ["combined_pair_map.csv","table_pair_map.csv","column_pair_map.csv","distinct_value.csv"]:
    lc=count_lines(os.path.join(DATA,fn))
    w(f"| {fn} | {lc:,} |")
w()

# ---- (b) trees ----
w("## (b) 유사도 클러스터 트리 (root별 군집 크기)")
for fn in ["table_value_tree.csv","table_str_tree.csv","column_value_tree.csv"]:
    df=pd.read_csv(os.path.join(DATA,fn),encoding="utf-8-sig",dtype=str)
    sizes=df.groupby("root")["node"].nunique().sort_values(ascending=False)
    w(f"### {fn}: edges={len(df)}, #roots(군집)={df['root'].nunique()}, "
      f"군집크기 max={sizes.max()}, median={int(sizes.median())}")
    w("상위 군집(root: #nodes): " + ", ".join(f"{r}:{c}" for r,c in sizes.head(8).items()))
    w()

# ---- (c) cross-domain shared values (SAMPLE) ----
w("## (c) 도메인 교차 공유 값 — 샘플 기반(국소). 규칙#4에 따라 추후 전체검증 필요\n")
SAMPLE_ROWS = 3_000_000
# row in distinct_value = one (col_id, value) -> #rows per value == #columns holding it
val_rows = collections.defaultdict(int)        # value -> #columns
val_mask = collections.defaultdict(int)        # value -> domain bitmask
val_tbls = collections.defaultdict(set)        # value -> set(table_id) (bounded by sampling)
DOMAIN_BIT = {"공공행정":1,"교육":2,"보건":4}
path=os.path.join(DATA,"distinct_value.csv")
seen=0
with io.open(path,"r",encoding="utf-8-sig",newline="") as f:
    r=csv.reader(f)
    header=next(r)
    idx={c:i for i,c in enumerate(header)}
    ci_val=idx["value"]; ci_c1=idx["cate1"]; ci_tbl=idx["table_id"]
    for row in r:
        seen+=1
        if seen>SAMPLE_ROWS: break
        v=row[ci_val]
        if v=="" or len(v)>40: continue
        val_rows[v]+=1
        val_mask[v]|=DOMAIN_BIT.get(row[ci_c1],0)
        if len(val_tbls[v])<50: val_tbls[v].add(row[ci_tbl])
w(f"샘플 행수={seen-1:,}, distinct value 수(샘플)={len(val_rows):,}")
ndomain = lambda m:(m&1>0)+(m&2>0)+(m&4>0)
multi=[(v,val_rows[v],ndomain(val_mask[v]),len(val_tbls[v])) for v in val_rows if ndomain(val_mask[v])>=2]
w(f"2개 이상 도메인에 걸친 값 종류 수(샘플)={len(multi):,}")
w(f"3개 도메인 전부에 걸친 값 종류 수(샘플)={sum(1 for x in multi if x[2]==3):,}")
w()
w("### 3개 도메인 전부 공유 + 최다 컬럼 등장 값 (상위 40)")
w("| value | #컬럼등장 | #도메인 | #테이블(≤50) |")
w("|---|---:|---:|---:|")
multi.sort(key=lambda x:(-x[2],-x[1]))
for v,nc,nd,nt in multi[:40]:
    w(f"| {v} | {nc} | {nd} | {nt} |")
w()
with io.open(OUT,"w",encoding="utf-8") as f: f.write("\n".join(buf))
print("WROTE",OUT)
