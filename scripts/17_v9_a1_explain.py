# -*- coding: utf-8 -*-
"""V9 deck: 10페이지에 'A1 비자명 연관규칙이 구체적으로 어떤 아이템 간 트랜잭션인지' 명확히 설명하는 카드 추가.

배경: 현재 V8 slide10의 좌상단(y 2.0~3.79″)이 비어 있음(이전 용어 카드를 사용자가 삭제). 그 빈 영역에
'A1 = 어떤 아이템 간 트랜잭션?' 카드를 신규 추가(기존 도형 제거/이동 없음 = 가장 안전).

설명 내용(scripts/06 §④⑤ 근거):
 - 트랜잭션 = 테이블 1건(속성 트랜잭션 A).
 - 아이템 = '속성=값' 8종: cate1·cate2·기관(AOI 롤업)·레코드·컬럼폭·수치비·유일키·상수컬럼(1D 이산화/롤업 산물).
 - A1 = 이 아이템들 사이 연관규칙 중 cate1↔cate2 자명종속(A2)을 제외한 '진짜' 규칙.
 - 구체 규칙(도식): [선행 cate2=유아및초·중등교육] ▶ [후행 기관=미상(소스)]  conf 1.00 · lift 6.89.

방식(안전): Base=V8. slide10.xml에 카드 1개(도형 ~10개) 추가 + 한글 잘림 방지. 그 외 슬라이드 byte-동일.
"""
import zipfile, os, re

SRC=r"d:\DataInt_2\report\데이터지도_분석보고서_V8.pptx"
DST=r"d:\DataInt_2\report\데이터지도_분석보고서_V9.pptx"
SLIDE="ppt/slides/slide10.xml"
def EMU(i): return int(round(i*914400))
NAVY="081F46"; GREEN="00965F"; DGRAY="505050"; BODYTX="282D37"; WHITE="FFFFFF"
LIGHT="EEF2F8"; LGREEN="E8F5F0"; LINE="D0D8E4"; LRED="FCE8E6"; RED="C0392B"; FONT="맑은 고딕"
HEADER_H=EMU(0.34)

_id=[700]
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
def _para(runs,algn="l",lnspc=112000,spcaft=20):
    rs="".join(_run(t,sz,bb,c) for (t,sz,bb,c) in runs)
    return (f'<a:p><a:pPr algn="{algn}"><a:lnSpc><a:spcPct val="{lnspc}"/></a:lnSpc>'
            f'<a:spcBef><a:spcPts val="0"/></a:spcBef><a:spcAft><a:spcPts val="{spcaft}"/></a:spcAft></a:pPr>{rs}</a:p>')
def tbox(x,y,w,h,paras,anchor="t",lins=12700,name="txt"):
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{nid()}" name="{name}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
            f'<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
            f'<p:txBody><a:bodyPr wrap="square" lIns="{lins}" tIns="6350" rIns="{lins}" bIns="6350" anchor="{anchor}"><a:normAutofit/></a:bodyPr>'
            f'<a:lstStyle/>{"".join(paras)}</p:txBody></p:sp>')
def chip(x,y,w,h,text,fill,line,tcol):
    return [roundrect(x,y,w,h,fill,line,name="a1-chip"),
            tbox(x,y,w,h,[_para([(text,820,True,tcol)],algn="ctr",spcaft=0)],anchor="ctr",lins=6350,name="a1-chiptx")]

# ---------- A1 card (left-top empty area: y 2.0~3.70) ----------
LX=EMU(0.62); LW=EMU(5.78); CY=EMU(2.0); CH=EMU(1.70)
shapes=[]
shapes.append(roundrect(LX,CY,LW,CH,LIGHT,LINE,name="a1-bg"))
shapes.append(roundrect(LX,CY,LW,HEADER_H,NAVY,NAVY,name="a1-hdr"))
shapes.append(tbox(LX,CY,LW,HEADER_H,[_para([("A1 비자명 연관규칙 — 어떤 ‘아이템’ 간 트랜잭션?",990,True,WHITE)],algn="l",spcaft=0)],anchor="ctr",lins=EMU(0.12),name="a1-title"))
shapes.append(tbox(LX+EMU(0.12),CY+HEADER_H+EMU(0.03),LW-EMU(0.22),EMU(0.66),[
    _para([("· 트랜잭션 = ",850,True,NAVY),("테이블 1건(속성 트랜잭션 A)",850,False,BODYTX)],spcaft=10),
    _para([("· 아이템 = ",850,True,NAVY),("‘속성=값’ 8종 — cate1·cate2·기관(AOI)·레코드·컬럼폭·수치비·유일키·상수컬럼",850,False,BODYTX)],spcaft=10),
    _para([("· A1 = ",850,True,NAVY),("이 아이템들 사이 규칙 중 cate1↔cate2 자명종속(A2)을 뺀 ‘진짜’ 규칙",850,False,BODYTX)],spcaft=0)],
    anchor="t",name="a1-text"))
# rule diagram: [선행] ▶ [후행]
DY=CY+EMU(1.05)
shapes+=chip(LX+EMU(0.16),DY,EMU(3.00),EMU(0.32),"cate2 = 유아및초·중등교육",LGREEN,GREEN,GREEN)
shapes.append(tbox(LX+EMU(3.18),DY,EMU(0.34),EMU(0.32),[_para([("▶",1400,True,DGRAY)],algn="ctr",spcaft=0)],anchor="ctr",lins=0,name="a1-arrow"))
shapes+=chip(LX+EMU(3.54),DY,EMU(1.98),EMU(0.32),"기관 = 미상(소스)",LRED,RED,RED)
shapes.append(tbox(LX+EMU(0.16),DY+EMU(0.34),LW-EMU(0.30),EMU(0.26),[
    _para([("규칙: 선행→후행  ·  conf 1.00 · lift 6.89  =  이 중분류면 출처가 100% ‘미상’(배너의 100%)",800,False,DGRAY)],spcaft=0)],anchor="t",name="a1-rulelabel"))

def fix_clipping(xml):
    def repl(m):
        bp=m.group(1)
        bp2=re.sub(r'wrap="[^"]*"','wrap="square"',bp,count=1) if 'wrap="' in bp else bp.replace('<a:bodyPr','<a:bodyPr wrap="square"',1)
        if '<a:spAutoFit/>' in bp2: pass
        elif '<a:normAutofit' in bp2: bp2=re.sub(r'<a:normAutofit[^>]*/>','<a:normAutofit/>',bp2)
        elif '<a:noAutofit/>' in bp2: bp2=bp2.replace('<a:noAutofit/>','<a:normAutofit/>')
        else:
            if bp2.endswith('/>'): bp2=bp2[:-2]+'><a:normAutofit/></a:bodyPr>'
            elif '<a:prstTxWarp' in bp2: bp2=re.sub(r'(</a:prstTxWarp>)',r'\1<a:normAutofit/>',bp2,count=1)
            else: bp2=re.sub(r'(<a:bodyPr[^>]*>)',r'\1<a:normAutofit/>',bp2,count=1)
        return '<p:txBody>'+bp2
    return re.sub(r'<p:txBody>(<a:bodyPr[^>]*/>|<a:bodyPr[^>]*>.*?</a:bodyPr>)',repl,xml,flags=re.DOTALL)

def main():
    with zipfile.ZipFile(SRC) as z:
        data={i.filename:z.read(i.filename) for i in z.infolist()}; order=z.infolist()
    xml=data[SLIDE].decode("utf-8")
    assert xml.count("</p:spTree>")==1
    xml=xml.replace("</p:spTree>","".join(shapes)+"</p:spTree>")
    xml=fix_clipping(xml)
    data[SLIDE]=xml.encode("utf-8")
    if os.path.exists(DST): os.remove(DST)
    with zipfile.ZipFile(DST,"w",zipfile.ZIP_DEFLATED) as zout:
        for i in order: zout.writestr(i, data[i.filename])
    print("A1 card shapes added:",len(shapes))
    print("WROTE",DST)

if __name__=="__main__":
    main()
