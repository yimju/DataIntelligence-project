# -*- coding: utf-8 -*-
"""Produce (1) exact full metrics for issue rules, (2) appendix of additional patterns,
(3) charts (PNG). Exports JSON for the PPTX builder + a metrics markdown.
Re-mines the 999-table baskets/attributes (fast, deterministic)."""
import os, io, re, json, numpy as np, pandas as pd
from sklearn.cluster import KMeans
from mlxtend.frequent_patterns import fpgrowth, association_rules
from mlxtend.preprocessing import TransactionEncoder
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
plt.rcParams["font.family"]="Malgun Gothic"; plt.rcParams["axes.unicode_minus"]=False

DATA=r"d:\data_intelligence_2\data"; AN=r"d:\data_intelligence_2\analysis"
FIG=r"d:\data_intelligence_2\report\figs"; os.makedirs(FIG,exist_ok=True)
NAVY="#081f46"; GREEN="#00965f"; GRAY="#9aa3b2"; LNAVY="#3b5a8c"

tm=pd.read_csv(os.path.join(DATA,"table_map.csv"),encoding="utf-8-sig",dtype=str).dropna(subset=["table_id"]).reset_index(drop=True)
cm=pd.read_csv(os.path.join(DATA,"column_map.csv"),encoding="utf-8-sig",dtype=str)
for c in ["column_cnt","categorical_cnt","numerical_cnt","rec_cnt"]: tm[c]=pd.to_numeric(tm[c],errors="coerce")
for c in ["distinct_cnt","pk_ratio","rec_cnt"]: cm[c]=pd.to_numeric(cm[c],errors="coerce")

# ---- column-concept baskets ----
def norm_col(x):
    if not isinstance(x,str): return None
    y=re.sub(r"\(.*?\)","",x.strip().lower()); y=re.sub(r"[\s　]+","",y); y=re.sub(r"^\d+[_\.]","",y)
    return y or None
cm["cc"]=cm["col_nm"].map(norm_col)
basket=cm.dropna(subset=["cc"]).groupby("table_id")["cc"].apply(lambda s:sorted(set(s)))
b_tx=list(basket.values); N=len(b_tx)
te=TransactionEncoder(); df=pd.DataFrame(te.fit(b_tx).transform(b_tx),columns=te.columns_)

def metrics(ante,cons):
    a=[x for x in ante if x in df.columns]; c=[x for x in cons if x in df.columns]
    if len(a)<len(ante) or len(c)<len(cons): return None
    sA=df[a].all(axis=1); sC=df[c].all(axis=1); sAC=(sA&sC)
    pA,pC,pAC=sA.mean(),sC.mean(),sAC.mean()
    conf=pAC/pA if pA else 0; lift=conf/pC if pC else 0
    lev=pAC-pA*pC; conv=(1-pC)/(1-conf) if conf<1 else float("inf")
    return dict(supp=round(pAC,4),cnt=int(round(pAC*N)),conf=round(conf,3),
                lift=round(lift,2),lev=round(lev,4),conv=(None if conf>=1 else round(conv,2)),
                anteN=int(sA.sum()),consN=int(sC.sum()))

# ---- curated ISSUE rules (exact metrics) ----
issue_rules=[
 ("지리좌표쌍",["경도"],["위도"]),
 ("기간쌍",["교육시작일"],["교육종료일"]),
 ("행정구역계층",["시도명"],["시군구명"]),
 ("주소이중표기",["지번주소"],["도로명주소"]),
 ("소재지주소(공공행정)",["소재지지번주소"],["소재지도로명주소"]),
 ("인구통계",["연령대"],["성별"]),
 ("의료기관PII",["의료기관전화번호"],["의료기관명","의료기관주소"]),
 ("교육스키마",["교과목코드"],["분반"]),
 ("교육스키마",["학종"],["학교명"]),
 ("사학재정",["연구비"],["일반기부금","지정기부금"]),
]
issue_out=[]
for topic,a,c in issue_rules:
    m=metrics(a,c)
    if m: issue_out.append(dict(topic=topic,ante="+".join(a),cons="+".join(c),**m))

# ---- mine ALL basket rules for appendix ----
fi=fpgrowth(df,min_support=8/N,use_colnames=True,max_len=3)
R=association_rules(fi,metric="lift",min_threshold=2.0)
R["antecedents"]=R["antecedents"].apply(lambda s:frozenset(s)); R["consequents"]=R["consequents"].apply(lambda s:frozenset(s))
# dedup symmetric pairs (A->B vs B->A): keep higher lift
R["pairkey"]=R.apply(lambda x:frozenset([frozenset(x["antecedents"]),frozenset(x["consequents"])]),axis=1)
R["unionkey"]=R.apply(lambda x:frozenset(x["antecedents"]|x["consequents"]),axis=1)
R=R.sort_values("lift",ascending=False).drop_duplicates("pairkey")
curated_keys={frozenset([frozenset(a),frozenset(c)]) for _,a,c in issue_rules}
# vocab already shown in deck (issue rules + schema-family slide) -> exclude to surface NEW topics
shown_vocab=[set(a)|set(c) for _,a,c in issue_rules]+[
 {"순번","의료기관명","의료기관전화번호","의료기관주소"},{"경도","위도","도로명주소","전화번호"},
 {"연도","직종","회차"},{"소재지전화","업소명","업종명"},
 {"법인명","연구비","일반기부금","지정기부금","학교명","학종"}]
cand=R[~R["pairkey"].isin(curated_keys)].sort_values("lift",ascending=False)
appendix_rows_df=[]; shown=list(shown_vocab)
for _,x in cand.iterrows():
    u=set(x["antecedents"]|x["consequents"])
    if max((len(u&s) for s in shown), default=0)>=2:   # too similar to something already shown
        continue
    appendix_rows_df.append(x); shown.append(u)
    if len(appendix_rows_df)>=14: break
appendix=pd.DataFrame(appendix_rows_df)
def setstr(s): return "{"+", ".join(sorted(s))+"}"
appx_rows=[]
for _,x in appendix.head(18).iterrows():
    conv=x["conviction"]; conv="∞" if not np.isfinite(conv) else round(conv,1)
    appx_rows.append(dict(rule=f"{setstr(x['antecedents'])}→{setstr(x['consequents'])}",
        supp=round(x["support"],4),cnt=int(round(x["support"]*N)),conf=round(x["confidence"],3),
        lift=round(x["lift"],2),lev=round(x["leverage"],4),conv=conv))

# ---- attribute association rules (non-structural) for appendix ----
def kbins(s,k,labels,log=True):
    X=(np.log1p(s.astype(float)) if log else s.astype(float)).values.reshape(-1,1)
    k=min(k,len(np.unique(X))); km=KMeans(n_clusters=k,n_init=10,random_state=0).fit(X)
    order=np.argsort(km.cluster_centers_.ravel()); remap={o:i for i,o in enumerate(order)}
    return pd.Series([labels[remap[l]] for l in km.labels_],index=s.index)
tm["REC"]=kbins(tm["rec_cnt"],5,["극소","소","중","대","극대"]).values
tm["COLN"]=kbins(tm["column_cnt"],4,["협소","표준","광폭","초광폭"]).values
nr=(tm["numerical_cnt"]/tm["column_cnt"].replace(0,np.nan)).fillna(0)
tm["NUMR"]=kbins(nr,3,["범주우세","혼합","수치우세"],log=False).values
g=cm.groupby("table_id"); has_key=(g["pk_ratio"].max()>=0.99); has_const=(g["distinct_cnt"].min()==1)
tmi=tm.set_index("table_id")
attr_tx=[]
for tid,row in tmi.iterrows():
    its=[f"cate1={row['cate1']}",f"cate2={row['cate2']}",f"레코드={row['REC']}",f"컬럼폭={row['COLN']}",
         f"수치비={row['NUMR']}","유일키="+("있음" if has_key.get(tid,False) else "없음"),
         "상수컬럼="+("있음" if has_const.get(tid,False) else "없음")]
    attr_tx.append(its)
te2=TransactionEncoder(); dfa=pd.DataFrame(te2.fit(attr_tx).transform(attr_tx),columns=te2.columns_)
fia=fpgrowth(dfa,min_support=0.05,use_colnames=True); Ra=association_rules(fia,metric="lift",min_threshold=1.3)
def isstruct(a,c):
    al=list(a)+list(c)
    return all(x.startswith(("cate1=","cate2=")) for x in al) and any(x.startswith("cate1=") for x in al) and any(x.startswith("cate2=") for x in al)
Ra["struct"]=Ra.apply(lambda x:isstruct(x["antecedents"],x["consequents"]),axis=1)
Ra["unionkey"]=Ra.apply(lambda x:frozenset(x["antecedents"]|x["consequents"]),axis=1)
Ra=Ra[(~Ra["struct"])&(Ra["confidence"]>=0.6)].sort_values("lift",ascending=False).drop_duplicates("unionkey")
attr_rows=[]
for _,x in Ra.head(12).iterrows():
    conv=x["conviction"]; conv="∞" if not np.isfinite(conv) else round(conv,1)
    attr_rows.append(dict(rule=f"{setstr(x['antecedents'])}→{setstr(x['consequents'])}",
        supp=round(x["support"],3),conf=round(x["confidence"],3),lift=round(x["lift"],2),conv=conv))

# ================= EXPORT JSON =================
export=dict(N=N, issue_rules=issue_out, appendix_basket=appx_rows, appendix_attr=attr_rows)
with io.open(os.path.join(AN,"I1_metrics.json"),"w",encoding="utf-8") as f:
    json.dump(export,f,ensure_ascii=False,indent=1)

# ================= METRICS MARKDOWN =================
buf=["# I1 상세 지표 & Appendix\n",f"전체 트랜잭션 N={N}. 모든 지표는 전체 기준 재계산.\n",
     "## A. 이슈 패턴 상세 지표",
     "| 주제 | 규칙 | #테이블 | support | confidence | lift | leverage | conviction |",
     "|---|---|---:|---:|---:|---:|---:|---:|"]
for r in issue_out:
    buf.append(f"| {r['topic']} | {r['ante']}→{r['cons']} | {r['cnt']} | {r['supp']} | {r['conf']} | {r['lift']} | {r['lev']} | {r['conv'] if r['conv'] is not None else '∞'} |")
buf+=["\n## B. Appendix — 추가 컬럼 연관규칙 (이슈 미포함, lift順·대칭중복 제거)",
      "| 규칙 | #테이블 | support | confidence | lift | leverage | conviction |","|---|---:|---:|---:|---:|---:|---:|"]
for r in appx_rows:
    buf.append(f"| {r['rule']} | {r['cnt']} | {r['supp']} | {r['conf']} | {r['lift']} | {r['lev']} | {r['conv']} |")
buf+=["\n## C. Appendix — 테이블 속성 연관규칙 (비자명, conf≥0.6)",
      "| 규칙 | support | confidence | lift | conviction |","|---|---:|---:|---:|---:|"]
for r in attr_rows:
    buf.append(f"| {r['rule']} | {r['supp']} | {r['conf']} | {r['lift']} | {r['conv']} |")
with io.open(os.path.join(AN,"I1_metrics_detail.md"),"w",encoding="utf-8") as f: f.write("\n".join(buf))

# ================= FIGURES =================
def style(ax):
    for s in ["top","right"]: ax.spines[s].set_visible(False)
    ax.tick_params(colors="#333"); ax.title.set_color(NAVY)

# fig1: domain distribution
fig,ax=plt.subplots(figsize=(4.2,3.0),dpi=150)
d=tm["cate1"].value_counts()
ax.bar(d.index,d.values,color=[NAVY,GREEN,LNAVY]); style(ax)
for i,v in enumerate(d.values): ax.text(i,v+5,str(v),ha="center",fontsize=10,color="#333")
ax.set_title("도메인(cate1) 분포",fontsize=12,fontweight="bold"); ax.set_ylim(0,460)
fig.tight_layout(); fig.savefig(os.path.join(FIG,"fig_domain.png")); plt.close(fig)

# fig2: rec_cnt discretization (log hist + boundaries)
fig,ax=plt.subplots(figsize=(6.4,3.0),dpi=150)
lr=np.log10(tm["rec_cnt"].clip(lower=1)); ax.hist(lr,bins=40,color=LNAVY,alpha=.85)
for b in [22,130,829,8019]: ax.axvline(np.log10(b),color=GREEN,ls="--",lw=1.4)
ax.set_title("레코드수 1D 이산화 (log10 스케일, 경계=초록선)",fontsize=12,fontweight="bold")
ax.set_xlabel("log10(rec_cnt)"); ax.set_ylabel("테이블 수"); style(ax)
for x,lab in zip([0.8,1.6,2.6,3.6,4.7],["극소","소","중","대","극대"]): ax.text(x,ax.get_ylim()[1]*.86,lab,ha="center",color=NAVY,fontsize=9,fontweight="bold")
fig.tight_layout(); fig.savefig(os.path.join(FIG,"fig_disc_rec.png")); plt.close(fig)

# fig3: AOI institution type
fig,ax=plt.subplots(figsize=(6.0,3.0),dpi=150)
order=["기초지자체","공공기관","미상(소스)","광역지자체","교육기관","중앙행정","기타"]
def src_type(s):
    if s in (None,"","소스"): return "미상(소스)"
    if re.search(r"(교육청|교육지원청|교육연수원|교육원|교육연구원|도서관|대학교|대학원|폴리텍|학교$)",s): return "교육기관"
    if re.search(r"(부$|청$|처$|위원회$)",s): return "중앙행정"
    if re.search(r"(공단|공사|진흥원|개발원|재단|연구원|시험원|평가원|관리공단|센터$|기금|공제회|진흥회|연수원|병원|암센터|보험공단)",s): return "공공기관"
    if re.search(r"(특별시$|광역시$|특별자치시$|도$|특별자치도$)",s): return "광역지자체"
    if re.search(r"(시$|군$|구$)",s): return "기초지자체"
    return "기타"
vc=tm["table_source"].map(src_type).value_counts().reindex(order).fillna(0)
cols=[GREEN if k=="미상(소스)" else NAVY for k in order]
ax.barh(order[::-1],vc.values[::-1],color=cols[::-1]); style(ax)
for i,v in enumerate(vc.values[::-1]): ax.text(v+4,i,str(int(v)),va="center",fontsize=9,color="#333")
ax.set_title("AOI 롤업: 기관유형 분포",fontsize=12,fontweight="bold"); ax.set_xlim(0,440)
fig.tight_layout(); fig.savefig(os.path.join(FIG,"fig_aoi.png")); plt.close(fig)

# fig4: issue rules lift bar
fig,ax=plt.subplots(figsize=(6.6,3.6),dpi=150)
lab=[f"{r['ante']}→{r['cons']}" for r in issue_out]; lift=[r['lift'] for r in issue_out]
ord_=np.argsort(lift); lab=[lab[i] for i in ord_]; lift=[lift[i] for i in ord_]
ax.barh(lab,lift,color=GREEN); style(ax)
for i,v in enumerate(lift): ax.text(v+1,i,f"{v}",va="center",fontsize=9,color="#333")
ax.set_title("이슈 연관규칙 lift (우연 대비 배수)",fontsize=12,fontweight="bold"); ax.set_xlim(0,135)
ax.tick_params(axis="y",labelsize=8.5)
fig.tight_layout(); fig.savefig(os.path.join(FIG,"fig_rules_lift.png")); plt.close(fig)

# fig5: rule#4 local vs global (소재지 도로명->지번 asymmetry)
fig,ax=plt.subplots(figsize=(5.4,3.2),dpi=150)
labels=["도로명→지번","지번→도로명","경도→위도"]; loc=[1.00,1.00,1.00]; glo=[0.44,0.92,1.00]
x=np.arange(len(labels)); wd=0.38
ax.bar(x-wd/2,loc,wd,label="국소(공공행정)",color=LNAVY); ax.bar(x+wd/2,glo,wd,label="전체(999)",color=GREEN)
ax.set_xticks(x); ax.set_xticklabels(labels,fontsize=9); ax.set_ylabel("confidence"); ax.set_ylim(0,1.15)
for i,v in enumerate(loc): ax.text(i-wd/2,v+.02,f"{v:.2f}",ha="center",fontsize=8)
for i,v in enumerate(glo): ax.text(i+wd/2,v+.02,f"{v:.2f}",ha="center",fontsize=8)
ax.set_title("규칙#4: 국소→전체 confidence 변화",fontsize=12,fontweight="bold"); ax.legend(fontsize=8,frameon=False); style(ax)
fig.tight_layout(); fig.savefig(os.path.join(FIG,"fig_rule4.png")); plt.close(fig)

# fig6: support vs lift scatter (all basket rules)
fig,ax=plt.subplots(figsize=(5.4,3.2),dpi=150)
ax.scatter(R["support"],R["lift"],s=14,color=LNAVY,alpha=.5,edgecolor="none")
iss=pd.DataFrame(issue_out); ax.scatter(iss["supp"],iss["lift"],s=40,color=GREEN,label="이슈 규칙",zorder=3)
ax.set_xlabel("support"); ax.set_ylabel("lift"); ax.set_title("연관규칙 분포 (support–lift)",fontsize=12,fontweight="bold")
ax.legend(fontsize=8,frameon=False); style(ax); ax.set_yscale("log")
fig.tight_layout(); fig.savefig(os.path.join(FIG,"fig_scatter.png")); plt.close(fig)

print("OK issues=%d appendix_basket=%d appendix_attr=%d figs=6"%(len(issue_out),len(appx_rows),len(attr_rows)))
