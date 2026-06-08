# -*- coding: utf-8 -*-
"""Compute deterministic examples for the G3 (page 11) rewrite. Output UTF-8 md."""
import pandas as pd, io, os, csv, collections
DATA=r"d:\data_intelligence_2\data"; OUT=r"d:\data_intelligence_2\analysis\_g3_examples.md"
cm=pd.read_csv(os.path.join(DATA,"column_map.csv"),encoding="utf-8-sig",dtype=str)
tm=pd.read_csv(os.path.join(DATA,"table_map.csv"),encoding="utf-8-sig",dtype=str).dropna(subset=["table_id"])
for c in ["distinct_cnt","pk_ratio"]: cm[c]=pd.to_numeric(cm[c],errors="coerce")
for c in ["column_cnt","categorical_cnt","numerical_cnt"]: tm[c]=pd.to_numeric(tm[c],errors="coerce")
B=[]; w=B.append

# 1) distinct_cnt meaning + constant count
w("# G3 예시 데이터 (자동계산)\n")
w("## 1) distinct_cnt 정의·분포")
w(f"- 전체 컬럼 수: {len(cm)}")
w(f"- distinct_cnt=1 (상수 컬럼): {(cm['distinct_cnt']==1).sum()}  ({100*(cm['distinct_cnt']==1).mean():.1f}%)")
w(f"- distinct_cnt 분위: min={cm['distinct_cnt'].min():.0f}, p50={cm['distinct_cnt'].median():.0f}, max={cm['distinct_cnt'].max():.0f}")

# 2/3) col_nm exact frequency = count of columns(=tables) bearing that name
w("\n## 2~3) col_nm 빈도 집계(기준=col_nm 문자열, 값=그 이름을 가진 컬럼 수)")
vc=cm["col_nm"].value_counts()
w("상위 12 col_nm:")
for k,v in vc.head(12).items(): w(f"  {k} : {v}")

def fam(names):
    rows=[]
    for n in names:
        sub=cm[cm["col_nm"]==n]
        if len(sub)==0: rows.append((n,0,None,None,None)); continue
        rows.append((n,len(sub),sub["distinct_cnt"].min(),
                     sub["pk_ratio"].median(), sub["pk_ratio"].quantile(.25)))
    return rows
w("\n일련번호 계열 (이름·#컬럼·min(distinct)·pk중앙·pk25%):")
for r in fam(["연번","순번","번호","일련번호"]): w(f"  {r}")
w("기준일 계열:")
for r in fam(["데이터기준일자","데이터기준일","기준일자","기준일","등록일"]): w(f"  {r}")

# 4) value fingerprint detail for the serial family (min_val, pk_ratio stats)
w("\n## 4) 동일 개념 비교 = '값 지문'(min_val·pk_ratio·distinct) 비교")
for n in ["연번","순번","번호"]:
    sub=cm[cm["col_nm"]==n]
    mins=sub["min_val"].dropna().astype(str)
    minmode=mins.mode().iloc[0] if len(mins) else "?"
    w(f"- {n}: #컬럼={len(sub)}, min_val 최빈='{minmode}', pk_ratio 중앙={sub['pk_ratio'].median():.3f}, "
      f"pk_ratio 범위=[{sub['pk_ratio'].min():.2f}~{sub['pk_ratio'].max():.2f}]")

# 5) column_cnt vs categorical+numerical mismatch
w("\n## 5) column_cnt vs (categorical+numerical) 정합성 검사")
tm["sum_cn"]=tm["categorical_cnt"]+tm["numerical_cnt"]
mis=tm[tm["column_cnt"]!=tm["sum_cn"]]
w(f"- 불일치 테이블 수: {len(mis)}")
w("- 예시 (table_nm | column_cnt | cat | num | 합):")
for _,r in mis.head(6).iterrows():
    nm=str(r['table_nm'])[:46]
    w(f"  {nm} | {r['column_cnt']:.0f} | {r['categorical_cnt']:.0f} | {r['numerical_cnt']:.0f} | {r['sum_cn']:.0f}")

# 6) constant-column most frequent names + an actual constant value example (join distinct_value)
w("\n## 6) 의미분석 근거 숫자")
const=cm[cm["distinct_cnt"]==1]
w(f"- 상수 컬럼 1,372개 중 최빈 col_nm:")
for k,v in const["col_nm"].value_counts().head(8).items(): w(f"  {k} : {v}")
w(f"- 일련번호 3이름 합계: 연번+순번+번호 = {vc.get('연번',0)+vc.get('순번',0)+vc.get('번호',0)} 테이블")

# actual constant VALUE examples from distinct_value for a few constant col_ids
const_ids=set(const["col_id"].astype(str).head(4000))
examples=[]
path=os.path.join(DATA,"distinct_value.csv")
with io.open(path,"r",encoding="utf-8-sig",newline="") as f:
    r=csv.reader(f); hdr=next(r); ci=hdr.index("col_id"); vi=hdr.index("value")
    seen=set()
    for row in r:
        cid=row[ci]
        if cid in const_ids and cid not in seen:
            val=row[vi]
            if val and val.strip() and not val.strip().isdigit() and len(val)<=12:
                nm=cm[cm["col_id"]==cid]["col_nm"]
                if len(nm): examples.append((nm.iloc[0],val)); seen.add(cid)
        if len(examples)>=12: break
w("- 상수 컬럼 실제 단일값 예시 (col_nm = 값):")
for nm,val in examples[:10]: w(f"  {nm} = '{val}'")

io.open(OUT,"w",encoding="utf-8").write("\n".join(map(str,B)))
print("WROTE",OUT)
