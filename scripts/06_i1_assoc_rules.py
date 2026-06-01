# -*- coding: utf-8 -*-
"""I1 — Attribute association rules (evidence-based).
Transactions = tables (999). Items = (A) discretized/rolled-up table attributes,
                                       (B) column-concept basket (col_nm present in table).
Algorithms: 1D KMeans discretization, rule-based AOI rollup, FP-Growth, association_rules(lift).
Rule #4: mine per-domain (local) then re-verify every rule on full set (global).
NO LLM reasoning — all items/metrics from deterministic code.
"""
import os, io, re, numpy as np, pandas as pd
from sklearn.cluster import KMeans
from mlxtend.frequent_patterns import fpgrowth, association_rules
from mlxtend.preprocessing import TransactionEncoder

DATA = r"d:\data_intelligence_2\data"
OUTDIR = r"d:\data_intelligence_2\analysis"
buf=[]; w=lambda s="":buf.append(str(s))

tm = pd.read_csv(os.path.join(DATA,"table_map.csv"), encoding="utf-8-sig", dtype=str).dropna(subset=["table_id"])
cm = pd.read_csv(os.path.join(DATA,"column_map.csv"), encoding="utf-8-sig", dtype=str)
for c in ["column_cnt","categorical_cnt","numerical_cnt","rec_cnt"]:
    tm[c]=pd.to_numeric(tm[c],errors="coerce")
for c in ["distinct_cnt","pk_ratio","same_ratio","rec_cnt"]:
    cm[c]=pd.to_numeric(cm[c],errors="coerce")

# ---------- 1D discretization (explainable, KMeans on log1p) ----------
def kbins(s, k, labels, log=True):
    s=s.astype(float)
    X=(np.log1p(s) if log else s).values.reshape(-1,1)
    uniq=np.unique(X)
    k=min(k,len(uniq))
    km=KMeans(n_clusters=k,n_init=10,random_state=0).fit(X)
    order=np.argsort(km.cluster_centers_.ravel()); remap={o:i for i,o in enumerate(order)}
    b=np.array([remap[l] for l in km.labels_])
    lab=[labels[min(i,len(labels)-1)] for i in b]
    # build readable boundaries
    bounds={}
    sv=s.values
    for i in range(k):
        m=b==i; bounds[labels[min(i,len(labels)-1)]]=(sv[m].min(),sv[m].max(),m.sum())
    return pd.Series(lab,index=s.index), bounds

tm=tm.reset_index(drop=True)
rec_lab, rec_b = kbins(tm["rec_cnt"], 5, ["극소","소","중","대","극대"])
col_lab, col_b = kbins(tm["column_cnt"], 4, ["협소","표준","광폭","초광폭"])
numratio = (tm["numerical_cnt"]/tm["column_cnt"].replace(0,np.nan)).fillna(0)
nr_lab, nr_b = kbins(numratio, 3, ["범주우세","혼합","수치우세"], log=False)

# ---------- AOI rollup: table_source -> 기관유형 (rule-based suffix parsing) ----------
def src_type(s):
    if s is None or s=="" or s=="소스": return "미상(소스placeholder)"
    if re.search(r"(교육청|교육지원청|교육연수원|교육원|교육연구원|도서관|대학교|대학원|폴리텍|학교$)", s): return "교육기관"
    if re.search(r"(부$|청$|처$|위원회$|^교육부|^보건복지부|^행정안전부|^고용노동부|^국방부|^법무부)", s): return "중앙행정기관"
    if re.search(r"(공단|공사|진흥원|개발원|재단|연구원|시험원|평가원|관리공단|센터$|기금|공제회|진흥회|연수원|병원|암센터|진흥재단|장학재단|보험공단)", s): return "공공기관/공기업"
    if re.search(r"(특별시$|광역시$|특별자치시$|도$|특별자치도$)", s): return "광역지자체"
    if re.search(r"(시$|군$|구$)", s): return "기초지자체"
    return "기타"
tm["src_type"]=tm["table_source"].map(src_type)

# ---------- column-concept basket: normalize col_nm deterministically ----------
def norm_col(x):
    if not isinstance(x,str): return None
    y=x.strip().lower()
    y=re.sub(r"\(.*?\)","",y)          # drop parenthetical
    y=re.sub(r"[\s　]+","",y)      # drop whitespace
    y=re.sub(r"^\d+[_\.]","",y)        # drop leading numeric prefix like 5_
    return y if y else None
cm["col_concept"]=cm["col_nm"].map(norm_col)

# ============ build per-table item set (attributes) ============
# per-table key/const flags from column_map
g=cm.groupby("table_id")
has_key  = (g["pk_ratio"].max()>=0.99)
has_const= (g["distinct_cnt"].min()==1)
catemax  = g["same_ratio"].max()       # high multiplicity present
tm=tm.set_index("table_id")
tm["REC"]=rec_lab.values; tm["COLN"]=col_lab.values; tm["NUMR"]=nr_lab.values
attr_items={}
for tid,row in tm.iterrows():
    its=[f"cate1={row['cate1']}", f"cate2={row['cate2']}", f"기관={row['src_type']}",
         f"레코드={row['REC']}", f"컬럼폭={row['COLN']}", f"수치비={row['NUMR']}"]
    its.append("유일키=" + ("있음" if has_key.get(tid,False) else "없음"))
    its.append("상수컬럼=" + ("있음" if has_const.get(tid,False) else "없음"))
    attr_items[tid]=its

# ============ build per-table column basket ============
basket = cm.dropna(subset=["col_concept"]).groupby("table_id")["col_concept"].apply(lambda s:sorted(set(s)))

# ---------- helpers ----------
def mine(transactions, min_support, metric_min_lift=1.0, max_len=None):
    te=TransactionEncoder(); arr=te.fit(transactions).transform(transactions)
    df=pd.DataFrame(arr,columns=te.columns_)
    fi=fpgrowth(df,min_support=min_support,use_colnames=True,max_len=max_len)
    if fi.empty: return pd.DataFrame()
    rules=association_rules(fi,metric="lift",min_threshold=metric_min_lift)
    return rules

def fmt_set(s): return "{" + ", ".join(sorted(s)) + "}"

def rules_table(rules, n=25, sort="lift"):
    if rules.empty:
        w("(규칙 없음)"); return
    r=rules.sort_values(sort,ascending=False).head(n)
    w("| 선행(antecedent) | → | 후행(consequent) | sup | conf | lift | lev | conv |")
    w("|---|:-:|---|---:|---:|---:|---:|---:|")
    for _,x in r.iterrows():
        conv = x['conviction']; conv = "inf" if not np.isfinite(conv) else f"{conv:.2f}"
        w(f"| {fmt_set(x['antecedents'])} | → | {fmt_set(x['consequents'])} | "
          f"{x['support']:.3f} | {x['confidence']:.3f} | {x['lift']:.2f} | {x['leverage']:.3f} | {conv} |")
    w()

# ================= report =================
w("# I1 — 속성 연관규칙 (근거중심 리포트)\n")
w(f"- 트랜잭션 = 테이블 {len(tm)}건 · 아이템 = 이산화/롤업 속성 + 보유 컬럼개념")
w("- 알고리즘: 1D KMeans 이산화 → AOI 롤업 → FP-Growth → association_rules(lift)")
w("- 규칙#4: 도메인별 채굴 후 전체 검증. 규칙#1: 모든 아이템/지표는 결정적 코드 산출.\n")

w("## 0. 1D 이산화 경계 (설명가능)")
def dump_bounds(name,b):
    w(f"- **{name}**: " + " · ".join(f"{k}[{int(v[0])}~{int(v[1])}]n={v[2]}" for k,v in b.items()))
dump_bounds("레코드수(rec_cnt)",rec_b); dump_bounds("컬럼폭(column_cnt)",col_b)
w(f"- **수치비(numerical/column)**: " + " · ".join(f"{k}[{v[0]:.2f}~{v[1]:.2f}]n={v[2]}" for k,v in nr_b.items()))
w(f"- **기관유형(AOI)**: " + ", ".join(f"{k}:{v}" for k,v in tm['src_type'].value_counts().items()))
w()

# ---------- A. attribute rules (global) ----------
w("## A. 테이블 속성 연관규칙 (전체 999, min_support=0.05, lift>1.2)")
w("> 구조적 자명규칙(cate1↔cate2)은 별도 표기/제외.\n")
attr_tx=list(attr_items.values())
rA=mine(attr_tx, min_support=0.05, metric_min_lift=1.2)
# flag structural cate1<->cate2
def is_struct(a,c):
    al=list(a)+list(c)
    return all(x.startswith("cate1=") or x.startswith("cate2=") for x in al) and \
           any(x.startswith("cate1=") for x in al) and any(x.startswith("cate2=") for x in al)
if not rA.empty:
    rA["struct"]=rA.apply(lambda x:is_struct(x["antecedents"],x["consequents"]),axis=1)
    w("### A1. 비자명 규칙 상위 (lift順)")
    rules_table(rA[~rA["struct"]], n=30)
    w("### A2. (참고) 구조적 자명규칙 예시")
    rules_table(rA[rA["struct"]], n=6)
else:
    w("(규칙 없음)")

# ---------- B. column-concept basket rules (global) ----------
w("\n## B. 컬럼개념 장바구니 연관규칙 — '기저귀>맥주' 대응 (전체 999)")
b_tx=list(basket.values)
w(f"- 평균 컬럼개념/테이블 = {np.mean([len(t) for t in b_tx]):.1f}, 고유 컬럼개념 = {len(set(c for t in b_tx for c in t))}")
# support threshold ~ 8 tables
ms=8/len(b_tx)
rB=mine(b_tx, min_support=ms, metric_min_lift=2.0, max_len=3)
w(f"### B1. 컬럼 동시출현 규칙 상위 (min_support≈{ms:.3f}=8테이블, lift>2, lift順)")
rules_table(rB, n=35)
if not rB.empty:
    w("### B2. 신뢰도 높은 강한 함의 (confidence>0.8, support≥0.01, lift順)")
    rules_table(rB[(rB["confidence"]>0.8)&(rB["support"]>=0.01)], n=25, sort="lift")

# ---------- C. local (per-domain) mining + global verification ----------
w("\n## C. 국소(도메인별) 채굴 → 전체 검증 (규칙#4)")
# build global lookup for basket support/conf/lift
te=TransactionEncoder(); arrG=te.fit(b_tx).transform(b_tx); dfG=pd.DataFrame(arrG,columns=te.columns_)
N=len(dfG)
def supp(items):
    items=[i for i in items if i in dfG.columns]
    if not items: return 0.0
    return dfG[items].all(axis=1).mean()
def verify(a,c):
    sa=supp(list(a)); sac=supp(list(a)+list(c)); sc=supp(list(c))
    conf=sac/sa if sa>0 else 0; lift=conf/sc if sc>0 else 0
    return sac,conf,lift
tm_reset=tm.reset_index()
for dom in ["보건","교육","공공행정"]:
    tids=set(tm_reset.loc[tm_reset["cate1"]==dom,"table_id"])
    tx=[basket[t] for t in basket.index if t in tids]
    if len(tx)<20: continue
    ms_d=max(5/len(tx),0.02)
    rd=mine(tx,min_support=ms_d,metric_min_lift=2.0,max_len=2)
    w(f"\n### C-{dom} (n={len(tx)} 테이블, 국소 lift>2 상위 12 → 전체 재검증)")
    if rd.empty: w("(국소 규칙 없음)"); continue
    rd=rd.sort_values("lift",ascending=False).head(12)
    w("| 규칙 | 국소 sup | 국소 conf | 국소 lift | ‖ | 전체 sup | 전체 conf | 전체 lift |")
    w("|---|---:|---:|---:|:-:|---:|---:|---:|")
    for _,x in rd.iterrows():
        gs,gc,gl=verify(x["antecedents"],x["consequents"])
        w(f"| {fmt_set(x['antecedents'])}→{fmt_set(x['consequents'])} | {x['support']:.3f} | {x['confidence']:.3f} | {x['lift']:.2f} | ‖ | {gs:.3f} | {gc:.3f} | {gl:.2f} |")
w()

with io.open(os.path.join(OUTDIR,"I1_association_rules.md"),"w",encoding="utf-8") as f:
    f.write("\n".join(buf))
print("WROTE I1_association_rules.md ; rules A=%d B=%d"%(0 if rA is None else len(rA), len(rB)))
