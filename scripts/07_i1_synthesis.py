# -*- coding: utf-8 -*-
"""I1 synthesis evidence: (1) provenance-loss breakdown, (2) maximal column families,
   (3) traceability (which tables/sources back the headline rules), (4) deduped themed rules."""
import os, io, re, numpy as np, pandas as pd
from mlxtend.frequent_patterns import fpgrowth
from mlxtend.preprocessing import TransactionEncoder

DATA=r"d:\data_intelligence_2\data"; OUT=r"d:\data_intelligence_2\analysis\I1_synthesis.md"
buf=[]; w=lambda s="":buf.append(str(s))
tm=pd.read_csv(os.path.join(DATA,"table_map.csv"),encoding="utf-8-sig",dtype=str).dropna(subset=["table_id"])
cm=pd.read_csv(os.path.join(DATA,"column_map.csv"),encoding="utf-8-sig",dtype=str)

# ---- (1) provenance loss breakdown ----
w("# I1 보강 근거\n\n## (1) 출처 결손(table_source='소스') 분해")
ph=tm["table_source"]=="소스"
w(f"- 전체 placeholder 데이터셋: {ph.sum()} / {len(tm)}")
w("\n**cate2별 placeholder 비율** (어느 분류에서 출처가 통째로 사라졌나)")
w("| cate2 | placeholder | 총 | 비율 |"); w("|---|---:|---:|---:|")
for c2,sub in tm.groupby("cate2"):
    p=(sub["table_source"]=="소스").sum()
    w(f"| {c2} | {p} | {len(sub)} | {100*p/len(sub):.0f}% |")
w()

# ---- (2) maximal column families ----
def norm_col(x):
    if not isinstance(x,str): return None
    y=re.sub(r"\(.*?\)","",x.strip().lower()); y=re.sub(r"[\s　]+","",y); y=re.sub(r"^\d+[_\.]","",y)
    return y or None
cm["cc"]=cm["col_nm"].map(norm_col)
basket=cm.dropna(subset=["cc"]).groupby("table_id")["cc"].apply(lambda s:sorted(set(s)))
tx=list(basket.values)
te=TransactionEncoder(); arr=te.fit(tx).transform(tx); df=pd.DataFrame(arr,columns=te.columns_)
fi=fpgrowth(df,min_support=8/len(tx),use_colnames=True,max_len=8)
fi["len"]=fi["itemsets"].apply(len)
fi=fi[fi["len"]>=3].sort_values(["len","support"],ascending=[False,False])
# maximal: not a subset of another frequent itemset
sets=[frozenset(s) for s in fi["itemsets"]]
maximal=[]
for i,(s,sup) in enumerate(zip(sets,fi["support"])):
    if not any((s< t) for t in sets):
        maximal.append((s,sup))
maximal=sorted(maximal,key=lambda x:(-len(x[0]),-x[1]))
# attach dominant cate1/source for each family
tid2c1=dict(zip(tm["table_id"],tm["cate1"])); tid2src=dict(zip(tm["table_id"],tm["table_source"]))
w("## (2) 최대빈발 컬럼군 = '스키마 패밀리' (동시출현 컬럼 묶음, ≥8 테이블)")
w("| 컬럼군 | 크기 | #테이블 | 지배 도메인 | 대표 출처 |"); w("|---|---:|---:|---|---|")
for s,sup in maximal[:20]:
    members=[t for t in basket.index if set(s)<=set(basket[t])]
    c1=pd.Series([tid2c1.get(t) for t in members]).value_counts()
    src=pd.Series([tid2src.get(t) for t in members]).value_counts()
    w(f"| {{{', '.join(sorted(s))}}} | {len(s)} | {len(members)} | {c1.index[0]}({c1.iloc[0]}) | {src.index[0]}({src.iloc[0]}) |")
w()

# ---- (3) traceability for headline rules ----
def tables_with(cols):
    cols=[norm_col(c) for c in cols]
    return [t for t in basket.index if all(c in set(basket[t]) for c in cols)]
w("## (3) 헤드라인 규칙 추적성 (규칙을 뒷받침하는 실제 테이블)")
for label,cols in [("경도+위도",["경도","위도"]),("의료기관전화번호",["의료기관전화번호"]),
                   ("연령대+성별",["연령대","성별"]),("교과목코드+분반",["교과목코드","분반"])]:
    ts=tables_with(cols);
    nm=[tm.loc[tm["table_id"]==t,"table_nm"].iloc[0] for t in ts[:3] if (tm["table_id"]==t).any()]
    w(f"- **{label}** : {len(ts)}개 테이블. 예) " + " / ".join(n[:45] for n in nm))
w()
with io.open(OUT,"w",encoding="utf-8") as f: f.write("\n".join(buf))
print("WROTE",OUT)
