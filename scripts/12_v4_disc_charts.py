# -*- coding: utf-8 -*-
"""V4 deck (rebuilt from clean V2, string-based edits — PowerPoint-safe).

목적(6페이지/1D 이산화):
 - 1D KMeans 이산화는 ① 레코드수(5구간,log) ② 컬럼폭(4구간,log) ③ 수치비(3구간,비율→log미적용)
   '3개 수치'에만 적용됨(scripts/06·09, I1_REPORT §0.2). 기존 슬라이드는 레코드수 1종만 시각화.
 - 컬럼폭·수치비를 함께 보이도록 '구간별 테이블 수' 막대차트 3종을 PowerPoint 네이티브 도형으로 작도.
 - 핵심메시지·bullet을 쉬운 용어로(멱법칙→쏠림 등), '멱법칙' 설명 콜아웃 재추가, 수치비 설명 보강.

설계 결정(중요):
 - 이 환경은 numpy/matplotlib C-확장이 깨져 raster 차트 생성 불가 → 막대=사각형 도형으로 직접 작도.
 - 직전 V3는 ET 전체 재직렬화 + 손수 만든(="p:style" 누락, 텍스트를 채운 도형 안에 둠) 콜아웃 때문에
   PowerPoint가 '복구'하며 콜아웃/캡션을 제거함. → V4는 (a) 문자열 편집으로 기존 바이트 보존,
   (b) python-pptx가 실제로 만든 도형 구조(<p:style> 포함, 채운 도형은 빈 txBody + 텍스트는 별도 박스)를
   그대로 복제하여 PowerPoint 거부를 방지.
 - 구간 경계/개수는 보고서의 KMeans 결과를 pure-python 비닝으로 재현(§0.2와 정확히 일치 확인).

Base=V2(헤드 18pt, 깨끗). slide6.xml만 수정(레코드수 raster pic 제거 후 3종 막대차트+콜아웃 삽입),
다른 슬라이드는 V2와 byte-동일.
"""
import zipfile, os, re, csv, io

SRC = r"d:\DataInt_2\report\데이터지도_분석보고서_V2.pptx"
DST = r"d:\DataInt_2\report\데이터지도_분석보고서_V4.pptx"
TM  = r"d:\DataInt_2\data\table_map.csv"
SLIDE = "ppt/slides/slide6.xml"

def EMU(inch): return int(round(inch*914400))
def PT(pt):   return int(round(pt*12700))
NAVY="081F46"; GREEN="00965F"; DGRAY="505050"; LINE="C7CFDA"; LNAVY="3B5A8C"
LGREEN="E8F5F0"; LGLINE="C7DED3"; FONT="맑은 고딕"
CORE_Y = str(EMU(1.28))   # 1170432

# ============== 1) discretization bins (pure python == KMeans result, validated) =======
def fnum(x):
    try: return float(x)
    except: return None
rows=list(csv.DictReader(io.open(TM,encoding="utf-8-sig")))
rec=[];col=[];nr=[]
for r in rows:
    rc=fnum(r["rec_cnt"]); cc=fnum(r["column_cnt"]); nc=fnum(r["numerical_cnt"])
    if rc is not None: rec.append(rc)
    if cc is not None: col.append(cc)
    nr.append((nc/cc) if (cc and cc>0 and nc is not None) else 0.0)
def binit(vals,edges,labels):
    b={l:[] for l in labels}
    for v in vals:
        for i,e in enumerate(edges):
            if v<=e: b[labels[i]].append(v); break
        else: b[labels[-1]].append(v)
    return b
recB=binit(rec,[22,130,829,8300],["극소","소","중","대","극대"])
colB=binit(col,[5,11,44],["협소","표준","광폭","초광폭"])
nrB =binit(nr,[0.23,0.59],["범주우세","혼합","수치우세"])
def ri(vs,last=False):
    lo,hi=int(min(vs)),int(max(vs)); return (f"{lo:,}~" if last else f"{lo:,}–{hi:,}")
def rr(vs): return f"{min(vs):.2f}–{max(vs):.2f}"
def pdata(title,sub,B,ratio=False):
    items=[]; L=list(B.keys())
    for i,l in enumerate(L):
        vs=B[l]; rng=(rr(vs) if ratio else ri(vs,last=(i==len(L)-1))) if vs else ""
        items.append((l,len(vs),rng))
    return dict(title=title,sub=sub,items=items,mx=max(c for _,c,_ in items))
PANELS=[pdata("레코드수 · 5구간","log 적용",recB),
        pdata("컬럼폭 · 4구간","log 적용",colB),
        pdata("수치비 · 3구간","비율→log 미적용",nrB,ratio=True)]

# ============== 2) PowerPoint-safe shape builders (mimic python-pptx) =================
_id=[200]
def nid(): _id[0]+=1; return _id[0]
STYLE=('<p:style><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef>'
       '<a:fillRef idx="3"><a:schemeClr val="accent1"/></a:fillRef>'
       '<a:effectRef idx="2"><a:schemeClr val="accent1"/></a:effectRef>'
       '<a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style>')
EMPTY_TX='<p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:endParaRPr/></a:p></p:txBody>'

def rect(x,y,w,h,fill,name="bar"):
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{nid()}" name="{name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
            f'<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
            f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill><a:ln><a:noFill/></a:ln><a:effectLst/></p:spPr>'
            f'{STYLE}{EMPTY_TX}</p:sp>')

def roundrect(x,y,w,h,fill,linecol,name="card"):
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{nid()}" name="{name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
            f'<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>'
            f'<a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>'
            f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>'
            f'<a:ln w="9525"><a:solidFill><a:srgbClr val="{linecol}"/></a:solidFill></a:ln><a:effectLst/></p:spPr>'
            f'{STYLE}{EMPTY_TX}</p:sp>')

def _para(text,sz,b,color,algn="ctr"):
    return (f'<a:p><a:pPr algn="{algn}"><a:lnSpc><a:spcPct val="102000"/></a:lnSpc>'
            f'<a:spcBef><a:spcPts val="0"/></a:spcBef><a:spcAft><a:spcPts val="0"/></a:spcAft></a:pPr>'
            f'<a:r><a:rPr sz="{sz}" b="{1 if b else 0}" i="0"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
            f'<a:latin typeface="{FONT}"/><a:ea typeface="{FONT}"/></a:rPr><a:t>{text}</a:t></a:r></a:p>')

def tbox(x,y,w,h,lines,anchor="t",algn="ctr",lins=12700,name="txt"):
    paras="".join(_para(t,sz,b,c,algn) for (t,sz,b,c) in lines)
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{nid()}" name="{name}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
            f'<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
            f'<p:txBody><a:bodyPr wrap="square" lIns="{lins}" tIns="6350" rIns="{lins}" bIns="6350" anchor="{anchor}"/>'
            f'<a:lstStyle/>{paras}</p:txBody></p:sp>')

# ============== 3) chart layout (right block) =========================================
X0=EMU(6.05); W=EMU(6.95); G=EMU(0.18); PANEL_W=(W-2*G)//3
Y_TITLE=EMU(2.04); H_TITLE=EMU(0.36)
Y_BASE=EMU(4.55); MAXBAR=EMU(1.78)
Y_LABEL=EMU(4.58); H_LABEL=EMU(0.54); PAD=EMU(0.12)
shapes=[]
# (chart header intentionally omitted — long 18pt core message spans the upper width;
#  panel titles + caption already label the block, avoiding overlap)
for pi,pn in enumerate(PANELS):
    pl=X0+pi*(PANEL_W+G)
    shapes.append(tbox(pl,Y_TITLE,PANEL_W,H_TITLE,
                  [(pn["title"],950,True,NAVY),(pn["sub"],750,False,DGRAY)],anchor="t",name=f"title{pi}"))
    shapes.append(rect(pl+PAD,Y_BASE,PANEL_W-2*PAD,PT(1.2),LINE,name=f"axis{pi}"))
    n=len(pn["items"]); usable=PANEL_W-2*PAD; slot=usable//n
    barw=min(int(slot*0.52),EMU(0.46))
    barcol=GREEN if pi==1 else (LNAVY if pi==2 else NAVY)
    for j,(name,cnt,rng) in enumerate(pn["items"]):
        h=int(MAXBAR*cnt/pn["mx"]) if pn["mx"] else 0
        h=max(h,EMU(0.025))
        bx=pl+PAD+slot*j+(slot-barw)//2; by=Y_BASE-h
        shapes.append(rect(bx,by,barw,h,barcol,name=f"bar{pi}_{j}"))
        shapes.append(tbox(pl+PAD+slot*j,by-EMU(0.19),slot,EMU(0.17),
                      [(f"{cnt}",800,True,DGRAY)],anchor="b",name=f"cnt{pi}_{j}"))
        shapes.append(tbox(pl+PAD+slot*j,Y_LABEL,slot,H_LABEL,
                      [(name,800,True,NAVY),(rng,650,False,DGRAY)],anchor="t",name=f"lab{pi}_{j}"))

# ============== 4) re-add the '멱법칙' callout (bg roundRect + text box) ===============
cx,cyy,cw,ch=EMU(0.62),EMU(4.95),EMU(5.3),EMU(1.5)
shapes.append(roundrect(cx,cyy,cw,ch,LGREEN,LGLINE,name="callout-bg"))
shapes.append(tbox(cx,cyy,cw,ch,
    [("▶ 쉬운 설명 — ‘멱법칙(쏠림 분포)’이란?",1150,True,GREEN),
     ("소수의 값만 아주 크고 대부분은 작은, 한쪽으로 치우친 분포예요. 예) 소수 데이터셋만 수십만 건인데 대다수는 수십~수백 건. log(로그)를 씌우면 큰 값이 눌려 고르게 펴지므로 공정하게 비교·구간화할 수 있습니다.",1000,False,"282D37")],
    anchor="t",algn="l",lins=82296,name="callout-text"))

# ============== 5) string edits on V2 slide6.xml =====================================
B1="• 문제: 레코드수 1~833,466로 극단 쏠림 → 그대로면 큰 값이 결과를 지배"
B2="• 해결: log(로그)로 쏠림을 편 뒤 1차원 KMeans로 자동 구간화"
B3="• 이산화한 수치는 3개뿐: 레코드수(5구간)·컬럼폭(4구간)·수치비(3구간)"
B4="• 수치비=수치형 컬럼 비율(0~1) — 이미 고른 값이라 로그 없이 바로 구간화"
B5="• 설명가능: 각 구간 = 원래값 [최소~최대] (우측 막대=구간별 테이블 수)"
CORE_NEW="한쪽으로 심하게 쏠린 수치(멱법칙)를 로그로 고르게 편 뒤 자동으로 구간을 나눔"
CAP_NEW="레코드수·컬럼폭·수치비를 1D KMeans로 이산화 — 막대=구간별 테이블 수(전체 999) · scripts/12·06"

OLD_BULLETS=[
 "• 문제: rec_cnt가 1~833,466으로 극단적 멱법칙 → 그대로면 큰 값이 지배",
 "• 절차: log1p 변환 → 1차원 KMeans(n_init=10) → 중심 정렬로 순서 라벨",
 "• 설명가능: 각 구간을 원척도 [min~max]로 환원(우측 초록선=경계)",
 "• 레코드수: 극소[1-22]·소[23-130]·중[133-829]·대[843-8019]·극대[8578~]",
 "• 컬럼폭: 협소[1-5]·표준[6-11]·광폭[12-44]·초광폭[49-338]",
]
NEW_BULLETS=[B1,B2,B3,B4,B5]
OLD_CAP="레코드수 분포와 KMeans 이산화 경계 — scripts/09"

def edit_core(xml):
    # core sp = the <p:sp> whose <a:off ... y="1170432">; set first <a:t> to CORE_NEW, empty rest
    def repl(m):
        block=m.group(0)
        if f'y="{CORE_Y}"' not in block: return block
        seen=[0]
        def t_repl(tm):
            seen[0]+=1
            return f"<a:t>{CORE_NEW}</a:t>" if seen[0]==1 else "<a:t></a:t>"
        return re.sub(r"<a:t>.*?</a:t>", t_repl, block, flags=re.DOTALL)
    return re.sub(r"<p:sp>.*?</p:sp>", repl, xml, flags=re.DOTALL)

def main():
    with zipfile.ZipFile(SRC) as z:
        data={i.filename:z.read(i.filename) for i in z.infolist()}
        order=z.infolist()
    xml=data[SLIDE].decode("utf-8")
    n_pic=xml.count("<p:pic>")
    # remove rec picture
    xml=re.sub(r"<p:pic>.*?</p:pic>","",xml,flags=re.DOTALL)
    # bullets
    for old,new in zip(OLD_BULLETS,NEW_BULLETS):
        assert f"<a:t>{old}</a:t>" in xml, f"bullet not found: {old[:20]}"
        xml=xml.replace(f"<a:t>{old}</a:t>",f"<a:t>{new}</a:t>")
    # caption
    assert f"<a:t>{OLD_CAP}</a:t>" in xml, "caption not found"
    xml=xml.replace(f"<a:t>{OLD_CAP}</a:t>",f"<a:t>{CAP_NEW}</a:t>")
    # core message
    xml=edit_core(xml)
    assert CORE_NEW in xml, "core rework failed"
    # insert new shapes before </p:spTree>
    inject="".join(shapes)
    assert xml.count("</p:spTree>")==1
    xml=xml.replace("</p:spTree>", inject+"</p:spTree>")
    data[SLIDE]=xml.encode("utf-8")

    if os.path.exists(DST): os.remove(DST)
    with zipfile.ZipFile(DST,"w",zipfile.ZIP_DEFLATED) as zout:
        for i in order: zout.writestr(i, data[i.filename])

    print("bins 레코드수:",[(l,len(v)) for l,v in recB.items()])
    print("bins 컬럼폭  :",[(l,len(v)) for l,v in colB.items()])
    print("bins 수치비  :",[(l,len(v)) for l,v in nrB.items()])
    print("pic removed:",n_pic,"| shapes added:",len(shapes))
    print("WROTE",DST)

if __name__=="__main__":
    main()
