# -*- coding: utf-8 -*-
"""V6 deck:
(A) 8페이지(트랜잭션·패턴·검증)에 '테이블 1건 = 1 트랜잭션(장바구니)' 예시 도표 추가.
    - 실제 예시 테이블(table_id=112, 전북특별자치도 무주군_행정기관)을 데이터에서 계산:
      6페이지 1D 이산화 결과(레코드=중·컬럼폭=표준·수치비=혼합)와 7페이지 AOI 롤업 결과(기관=기초지자체)를
      '속성 아이템'으로, 보유 컬럼개념 7종을 '장바구니 아이템'으로 담은 트랜잭션을 시각화.
    - 좌측 bullet은 폭 축소, 우측에 [원본 테이블] ▼ [트랜잭션=장바구니(칩)] ▼ [FP-Growth→연관규칙] 흐름.
(B) 전체 슬라이드의 '텍스트 박스 한글 잘림 방지': 모든 <p:txBody>의 <a:bodyPr>에 wrap="square" 보장 +
    autofit을 spAutoFit(있으면 유지)/normAutofit(없거나 noAutofit이면 부여)로 설정 → 텍스트가 잘리지 않음
    (spAutoFit=상자가 글자에 맞게 커짐, normAutofit=글자가 상자에 맞게 줄어듦; 둘 다 잘림 없음).

Base=V5. 안전방식(문자열 편집·python-pptx 검증 도형구조 복제·좌표식별). slide8 외 슬라이드는
(B) 한글잘림 방지 bodyPr 변경만 적용.
"""
import zipfile, os, re, csv, io

SRC=r"d:\DataInt_2\report\데이터지도_분석보고서_V5.pptx"
DST=r"d:\DataInt_2\report\데이터지도_분석보고서_V6.pptx"
TM =r"d:\DataInt_2\data\table_map.csv"
CM =r"d:\DataInt_2\data\column_map.csv"
SLIDE8="ppt/slides/slide8.xml"
EX_TID="112"

def EMU(i): return int(round(i*914400))
NAVY="081F46"; GREEN="00965F"; DGRAY="505050"; BODYTX="282D37"; WHITE="FFFFFF"
LIGHT="EEF2F8"; LGREEN="E8F5F0"; LINE="D0D8E4"; LGLINE="C7DED3"; CHIPLN="9DB0CC"; FONT="맑은 고딕"

# ================= (A-0) compute the example transaction from data =================
def fnum(x):
    try: return float(x)
    except: return None
def b(v,edges,labs):
    for i,e in enumerate(edges):
        if v<=e: return labs[i]
    return labs[-1]
def norm_col(x):
    y=re.sub(r"\(.*?\)","",(x or "").strip().lower()); y=re.sub(r"[\s　]+","",y); y=re.sub(r"^\d+[_\.]","",y); return y
trow=next(r for r in csv.DictReader(io.open(TM,encoding="utf-8-sig")) if r["table_id"]==EX_TID)
ccols=[r for r in csv.DictReader(io.open(CM,encoding="utf-8-sig")) if r.get("table_id")==EX_TID]
rec=fnum(trow["rec_cnt"]); col=fnum(trow["column_cnt"]); num=fnum(trow["numerical_cnt"]); cat=fnum(trow["categorical_cnt"])
REC=b(rec,[22,130,829,8300],["극소","소","중","대","극대"])
COLN=b(col,[5,11,44],["협소","표준","광폭","초광폭"])
NUMR=b((num/col) if col else 0,[0.23,0.59],["범주우세","혼합","수치우세"])
src=(trow.get("table_source") or "").strip()
def src_type(s):
    if s is None or s=="" or s=="소스": return "미상(소스)"
    if re.search(r"(교육청|교육지원청|교육연수원|교육원|교육연구원|도서관|대학교|대학원|폴리텍|학교$)",s): return "교육기관"
    if re.search(r"(부$|청$|처$|위원회$|^교육부|^보건복지부|^행정안전부|^고용노동부|^국방부|^법무부)",s): return "중앙행정"
    if re.search(r"(공단|공사|진흥원|개발원|재단|연구원|시험원|평가원|관리공단|센터$|기금|공제회|진흥회|연수원|병원|암센터|진흥재단|장학재단|보험공단)",s): return "공공기관"
    if re.search(r"(특별시$|광역시$|특별자치시$|도$|특별자치도$)",s): return "광역지자체"
    if re.search(r"(시$|군$|구$)",s): return "기초지자체"
    return "기타"
GIGWAN=src_type(src)
concepts=[norm_col(r.get("col_nm")) for r in ccols]
has_key=any((fnum(r.get("pk_ratio")) or 0)>=0.99 for r in ccols)
has_const=any(int(fnum(r.get("distinct_cnt")) or 0)==1 for r in ccols)
TBL_NM=trow["table_nm"]; CATE1=trow["cate1"]; CATE2=trow["cate2"]

# ================= (A-1) PowerPoint-safe shape builders =================
_id=[400]
def nid(): _id[0]+=1; return _id[0]
STYLE=('<p:style><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef>'
       '<a:fillRef idx="3"><a:schemeClr val="accent1"/></a:fillRef>'
       '<a:effectRef idx="2"><a:schemeClr val="accent1"/></a:effectRef>'
       '<a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style>')
EMPTY_TX='<p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:endParaRPr/></a:p></p:txBody>'
def roundrect(x,y,w,h,fill,linecol,name="rr"):
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{nid()}" name="{name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
            f'<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>'
            f'<a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>'
            f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>'
            f'<a:ln w="9525"><a:solidFill><a:srgbClr val="{linecol}"/></a:solidFill></a:ln><a:effectLst/></p:spPr>'
            f'{STYLE}{EMPTY_TX}</p:sp>')
def _run(text,sz,b_,color):
    return (f'<a:r><a:rPr sz="{sz}" b="{1 if b_ else 0}" i="0"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
            f'<a:latin typeface="{FONT}"/><a:ea typeface="{FONT}"/></a:rPr><a:t>{text}</a:t></a:r>')
def _para(runs,algn="l",lnspc=120000,spcaft=0):
    rs="".join(_run(t,sz,bb,c) for (t,sz,bb,c) in runs)
    return (f'<a:p><a:pPr algn="{algn}"><a:lnSpc><a:spcPct val="{lnspc}"/></a:lnSpc>'
            f'<a:spcBef><a:spcPts val="0"/></a:spcBef><a:spcAft><a:spcPts val="{spcaft}"/></a:spcAft></a:pPr>{rs}</a:p>')
def tbox(x,y,w,h,paras,anchor="t",lins=12700,name="txt"):
    body="".join(paras)
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{nid()}" name="{name}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
            f'<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
            f'<p:txBody><a:bodyPr wrap="square" lIns="{lins}" tIns="6350" rIns="{lins}" bIns="6350" anchor="{anchor}"><a:normAutofit/></a:bodyPr>'
            f'<a:lstStyle/>{body}</p:txBody></p:sp>')
HEADER_H=EMU(0.32)
def card(x,y,w,h,hdr,bg,line,title,body_paras):
    return [roundrect(x,y,w,h,bg,line,name="card-bg"),
            roundrect(x,y,w,HEADER_H,hdr,hdr,name="card-hdr"),
            tbox(x,y,w,HEADER_H,[_para([(title,1000,True,WHITE)],algn="ctr")],anchor="ctr",name="card-title"),
            tbox(x+EMU(0.12),y+HEADER_H+EMU(0.04),w-EMU(0.22),h-HEADER_H-EMU(0.08),body_paras,anchor="t",name="card-body")]
def chip(x,y,w,h,text):
    return [roundrect(x,y,w,h,LIGHT,CHIPLN,name="chip-bg"),
            tbox(x,y,w,h,[_para([(text,850,False,NAVY)],algn="ctr")],anchor="ctr",lins=6350,name="chip-tx")]

# ================= (A-2) build slide8 diagram =================
RX=EMU(4.9); RW=EMU(7.78)
T_Y=EMU(2.0); T_H=EMU(0.94)
A1_Y=EMU(3.00); A1_H=EMU(0.30)
BK_Y=EMU(3.34); BK_H=EMU(2.12)
A2_Y=EMU(5.52); A2_H=EMU(0.30)
LG_Y=EMU(5.88); LG_H=EMU(0.28)
shapes=[]
# [1] source table
shapes+=card(RX,T_Y,RW,T_H,NAVY,LIGHT,LINE,"예시 테이블 1건",
    [_para([(TBL_NM,900,True,NAVY)]),
     _para([(f"rec_cnt={int(rec)} · column_cnt={int(col)} (범주 {int(cat)}·수치 {int(num)}) · 출처=‘{src}’",830,False,BODYTX)])])
# arrow 1
shapes.append(tbox(RX,A1_Y,RW,A1_H,[_para([("▼  ① 1D 이산화(6p) · ② AOI 롤업(7p) 적용 → 아이템 생성",950,True,GREEN)],algn="ctr")],anchor="ctr",name="arrow1"))
# [2] basket (transaction)
shapes.append(roundrect(RX,BK_Y,RW,BK_H,LGREEN,LGLINE,name="basket-bg"))
shapes.append(roundrect(RX,BK_Y,RW,HEADER_H,GREEN,GREEN,name="basket-hdr"))
shapes.append(tbox(RX,BK_Y,RW,HEADER_H,[_para([("트랜잭션(장바구니) = 테이블 1건 = 1 거래",1000,True,WHITE)],algn="ctr")],anchor="ctr",name="basket-title"))
BKX=RX+EMU(0.16); BKW=RW-EMU(0.32)
by=BK_Y+HEADER_H+EMU(0.05)
shapes.append(tbox(BKX,by,BKW,EMU(0.22),[_para([("① 속성 아이템 (A)",850,True,NAVY)])],anchor="t",name="subA"))
shapes.append(tbox(BKX,by+EMU(0.21),BKW,EMU(0.46),[
    _para([(f"cate1={CATE1} · cate2={CATE2} · ",880,False,BODYTX),
           (f"레코드={REC} · 컬럼폭={COLN} · 수치비={NUMR}",880,True,GREEN)]),
    _para([(f"기관={GIGWAN}",880,True,GREEN),
           (f"  ·  유일키={'있음' if has_key else '없음'} · 상수컬럼={'있음' if has_const else '없음'}",880,False,BODYTX)])],
    anchor="t",name="attrA"))
shapes.append(tbox(BKX,by+EMU(0.70),BKW,EMU(0.22),[_para([(f"② 컬럼개념 장바구니 (B) — 보유 컬럼개념 {len(concepts)}종",850,True,NAVY)])],anchor="t",name="subB"))
# concept chips (2 rows)
cy0=by+EMU(0.94); slot=BKW//4; chw=int(slot*0.92); chh=EMU(0.28)
for k,cc in enumerate(concepts):
    row,coli=divmod(k,4)
    cx=BKX+coli*slot+(slot-chw)//2
    shapes+=chip(cx,cy0+row*EMU(0.31),chw,chh,cc)
# arrow 2
shapes.append(tbox(RX,A2_Y,RW,A2_H,[_para([("▼  999개 테이블 = 999개 트랜잭션  →  FP-Growth  →  연관규칙(lift)",950,True,GREEN)],algn="ctr")],anchor="ctr",name="arrow2"))
# legend
shapes.append(tbox(RX,LG_Y,RW,LG_H,[_para([("※ 초록 = 6·7페이지의 1D 이산화·AOI 롤업 결과 항목",820,False,DGRAY)],algn="ctr")],anchor="ctr",name="legend"))

# ================= (B) global no-clip bodyPr fix =================
def fix_clipping(xml):
    def repl(m):
        bp=m.group(1)
        bp2=re.sub(r'wrap="[^"]*"','wrap="square"',bp,count=1) if 'wrap="' in bp else bp.replace('<a:bodyPr','<a:bodyPr wrap="square"',1)
        if '<a:spAutoFit/>' in bp2:
            pass
        elif '<a:normAutofit' in bp2:
            bp2=re.sub(r'<a:normAutofit[^>]*/>','<a:normAutofit/>',bp2)
        elif '<a:noAutofit/>' in bp2:
            bp2=bp2.replace('<a:noAutofit/>','<a:normAutofit/>')
        else:
            if bp2.endswith('/>'):
                bp2=bp2[:-2]+'><a:normAutofit/></a:bodyPr>'
            elif '<a:prstTxWarp' in bp2:
                bp2=re.sub(r'(</a:prstTxWarp>)',r'\1<a:normAutofit/>',bp2,count=1)
            else:
                bp2=re.sub(r'(<a:bodyPr[^>]*>)',r'\1<a:normAutofit/>',bp2,count=1)
        return '<p:txBody>'+bp2
    return re.sub(r'<p:txBody>(<a:bodyPr[^>]*/>|<a:bodyPr[^>]*>.*?</a:bodyPr>)',repl,xml,flags=re.DOTALL)

# ================= main =================
def main():
    with zipfile.ZipFile(SRC) as z:
        data={i.filename:z.read(i.filename) for i in z.infolist()}; order=z.infolist()
    # (A) slide8: narrow bullets + inject diagram
    s8=data[SLIDE8].decode("utf-8")
    assert '<a:ext cx="11057839" cy="4526280"/>' in s8, "bullets ext not found"
    s8=s8.replace('<a:ext cx="11057839" cy="4526280"/>','<a:ext cx="3840480" cy="4526280"/>')  # narrow bullets to 4.2in
    assert s8.count("</p:spTree>")==1
    s8=s8.replace("</p:spTree>","".join(shapes)+"</p:spTree>")
    data[SLIDE8]=s8.encode("utf-8")
    # (B) no-clip on all slides
    n_slides=0; n_bodypr=0
    for name in list(data):
        if re.match(r"ppt/slides/slide\d+\.xml$",name):
            xml=data[name].decode("utf-8")
            n_bodypr+=len(re.findall(r'<p:txBody><a:bodyPr',xml))
            data[name]=fix_clipping(xml).encode("utf-8")
            n_slides+=1
    if os.path.exists(DST): os.remove(DST)
    with zipfile.ZipFile(DST,"w",zipfile.ZIP_DEFLATED) as zout:
        for i in order: zout.writestr(i, data[i.filename])
    print("example tid",EX_TID,"=",TBL_NM)
    print("  bins: 레코드=%s 컬럼폭=%s 수치비=%s | 기관=%s | 유일키=%s 상수컬럼=%s"%(REC,COLN,NUMR,GIGWAN,has_key,has_const))
    print("  concepts:",concepts)
    print("slide8 shapes added:",len(shapes))
    print("no-clip applied: slides=%d, p:txBody bodyPr=%d"%(n_slides,n_bodypr))
    print("WROTE",DST)

if __name__=="__main__":
    main()
