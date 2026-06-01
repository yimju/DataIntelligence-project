# -*- coding: utf-8 -*-
"""Basic descriptive analysis + AOI rollups + 1D discretization preview + governance quick-scan.
Outputs a UTF-8 markdown report. NO LLM reasoning -- pure pandas/sklearn aggregation."""
import os, io, numpy as np, pandas as pd
from sklearn.cluster import KMeans

DATA = r"d:\data_intelligence_2\data"
OUT  = r"d:\data_intelligence_2\analysis\basic_analysis.md"
os.makedirs(os.path.dirname(OUT), exist_ok=True)
buf = []
def w(s=""): buf.append(str(s))

def md_counts(series, top=None):
    vc = series.value_counts(dropna=False)
    if top: vc = vc.head(top)
    w("| value | count | pct |")
    w("|---|---:|---:|")
    n = len(series)
    for k, v in vc.items():
        w(f"| {k} | {v} | {100*v/n:.1f}% |")
    w()

def num_summary(df, cols):
    w("| field | n | mean | std | min | p25 | p50 | p75 | p90 | p99 | max | %zero | %null |")
    w("|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|")
    for c in cols:
        s = pd.to_numeric(df[c], errors="coerce")
        n = s.notna().sum()
        if n == 0:
            w(f"| {c} | 0 | - | - | - | - | - | - | - | - | - | - | 100% |"); continue
        q = s.quantile([.25,.5,.75,.9,.99])
        pz = 100*(s==0).sum()/len(s)
        pn = 100*s.isna().sum()/len(s)
        w(f"| {c} | {n} | {s.mean():.2f} | {s.std():.2f} | {s.min():.2f} | {q[.25]:.2f} | {q[.5]:.2f} | {q[.75]:.2f} | {q[.9]:.2f} | {q[.99]:.2f} | {s.max():.2f} | {pz:.1f}% | {pn:.1f}% |")
    w()

# ================= TABLE MAP =================
tm = pd.read_csv(os.path.join(DATA,"table_map.csv"), encoding="utf-8-sig", dtype=str)
w("# 기초 분석 리포트 (basic_analysis)\n")
w(f"## 1. table_map — 데이터셋 {len(tm)}건\n")
w("### 1.1 cate1 (대분류) 분포"); md_counts(tm["cate1"])
w("### 1.2 cate2 (중분류) 분포"); md_counts(tm["cate2"])
w("### 1.3 table_source / src_platform / language")
md_counts(tm["table_source"]); md_counts(tm["table_src_platform"]); md_counts(tm["language"])

w("### 1.4 cate1 × cate2 교차표")
ct = pd.crosstab(tm["cate1"], tm["cate2"])
w("| cate1 \\ cate2 | " + " | ".join(ct.columns) + " | 합 |")
w("|" + "---|"*(len(ct.columns)+2))
for idx, row in ct.iterrows():
    w(f"| {idx} | " + " | ".join(str(x) for x in row.values) + f" | {row.sum()} |")
w(f"| **합** | " + " | ".join(str(x) for x in ct.sum().values) + f" | {ct.values.sum()} |")
w()

w("### 1.5 테이블 수치 통계")
num_summary(tm, ["column_cnt","categorical_cnt","numerical_cnt","rec_cnt","table_qty_index"])

# ================= COLUMN MAP =================
cm = pd.read_csv(os.path.join(DATA,"column_map.csv"), encoding="utf-8-sig", dtype=str)
w(f"\n## 2. column_map — 컬럼 {len(cm)}건\n")
w("### 2.1 num_cat_flag 분포 (1=범주형 / 0=수치형 추정)"); md_counts(cm["num_cat_flag"])
# cross check: per table categorical_cnt vs numerical_cnt sum
w("### 2.2 num_cat_flag × cate1")
ct2 = pd.crosstab(cm["cate1"], cm["num_cat_flag"])
w("| cate1 \\ flag | " + " | ".join(str(c) for c in ct2.columns) + " |")
w("|" + "---|"*(len(ct2.columns)+1))
for idx,row in ct2.iterrows():
    w(f"| {idx} | " + " | ".join(str(x) for x in row.values) + " |")
w()
w("### 2.3 컬럼 프로파일 수치 통계")
num_summary(cm, ["distinct_cnt","null_ratio","same_ratio","pk_ratio","rec_cnt","cum_search_num"])

# ================= 1D DISCRETIZATION (explainable) =================
w("\n## 3. 1D 이산화/클러스터링 미리보기 (설명가능)\n")
def discretize_report(df, col, k=4, log=False):
    s = pd.to_numeric(df[col], errors="coerce").dropna()
    if log: s_t = np.log1p(s)
    else:   s_t = s
    X = s_t.values.reshape(-1,1)
    km = KMeans(n_clusters=min(k,len(np.unique(X))), n_init=10, random_state=0).fit(X)
    lab = km.labels_
    # order clusters by center
    order = np.argsort(km.cluster_centers_.ravel())
    remap = {old:i for i,old in enumerate(order)}
    binid = np.array([remap[l] for l in lab])
    w(f"**{col}** (KMeans k={km.n_clusters}{' on log1p' if log else ''}, n={len(s)}):")
    w("| bin | n | 범위(원척도) | mean |")
    w("|---:|---:|---|---:|")
    sv = s.values
    for b in range(km.n_clusters):
        m = binid==b
        w(f"| {b} | {m.sum()} | [{sv[m].min():.3g}, {sv[m].max():.3g}] | {sv[m].mean():.3g} |")
    w()
discretize_report(cm, "null_ratio", k=4)
discretize_report(cm, "pk_ratio", k=4)
discretize_report(cm, "distinct_cnt", k=5, log=True)
discretize_report(tm, "rec_cnt", k=5, log=True)
discretize_report(tm, "column_cnt", k=4, log=True)

# ================= GOVERNANCE QUICK SCAN =================
w("\n## 4. 데이터 거버넌스 퀵스캔\n")
# 4.1 duplicate table names
dup_tbl = tm["table_nm"].value_counts()
dup_tbl = dup_tbl[dup_tbl>1]
w(f"### 4.1 중복 table_nm: {len(dup_tbl)}건 (동일 파일명 여러 table_id)")
for k,v in dup_tbl.head(15).items(): w(f"- ({v}) {k}")
w()
# 4.2 duplicate column names (global)
dup_col = cm["col_nm"].value_counts()
w(f"### 4.2 최빈 col_nm (여러 테이블 공유 = 표준/중복 후보) — 상위 25")
w("| col_nm | #테이블등장 |")
w("|---|---:|")
for k,v in dup_col.head(25).items(): w(f"| {k} | {v} |")
w()
# 4.3 quality flags
cm_null = pd.to_numeric(cm["null_ratio"], errors="coerce")
cm_pk   = pd.to_numeric(cm["pk_ratio"], errors="coerce")
cm_dist = pd.to_numeric(cm["distinct_cnt"], errors="coerce")
cm_same = pd.to_numeric(cm["same_ratio"], errors="coerce")
w("### 4.3 품질 플래그")
w(f"- null_ratio > 0.5 인 컬럼: {(cm_null>0.5).sum()} ({100*(cm_null>0.5).mean():.1f}%)")
w(f"- null_ratio = 1.0 (전부 결측) 컬럼: {(cm_null>=1.0).sum()}")
w(f"- distinct_cnt = 1 (상수 컬럼): {(cm_dist==1).sum()}")
w(f"- pk_ratio >= 0.99 (거의 유일=키 후보): {(cm_pk>=0.99).sum()}")
w(f"- pk_ratio = 0 컬럼: {(cm_pk==0).sum()}")
w(f"- null_ratio 결측(빈값) 컬럼: {cm_null.isna().sum()}")
w()
# 4.4 language anomalies / empty categories
w("### 4.4 분류/언어 이상치")
w(f"- table_map cate1 결측: {tm['cate1'].isna().sum()}, cate2 결측: {tm['cate2'].isna().sum()}")
w(f"- language != Korean (table): {(tm['language']!='Korean').sum()}")
w(f"- rec_cnt = 0 또는 결측 테이블: {(pd.to_numeric(tm['rec_cnt'],errors='coerce').fillna(0)==0).sum()}")
w(f"- column_cnt vs (cat+num) 불일치 테이블: " +
  str((pd.to_numeric(tm['column_cnt'],errors='coerce') !=
       (pd.to_numeric(tm['categorical_cnt'],errors='coerce')+pd.to_numeric(tm['numerical_cnt'],errors='coerce'))).sum()))
w()

with io.open(OUT,"w",encoding="utf-8") as f:
    f.write("\n".join(buf))
print("WROTE", OUT, "lines", len(buf))
