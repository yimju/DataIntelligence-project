# -*- coding: utf-8 -*-
"""V14: 11페이지를 '값 지문' 두 대표 사례 중심으로 재구성.
요청: 연번/순번/번호(min=1·pk≈1.0)와 데이터기준일자/데이터기준일/기준일자(pk≈0.0)의 대비를
      분석의 중심으로. 부가 이슈(스키마 불일치 등) 제거, 상수(distinct=1)는 사례 B 근거로 흡수.

구성:
 - 사례 A · 행 일련번호 지문 (min=1·pk≈1.0): 연번168·순번79·번호34 → 행 식별자, 281테이블 통합
 - 사례 B · 데이터 기준일 지문 (pk≈0.0·상수): 데이터기준일자169·데이터기준일46·기준일자10 → 스냅샷 날짜
   (근거: 상수컬럼 최빈 1·2위 = 데이터기준일자166·데이터기준일44)
 - 하단 밴드: 판별 방법(col_nm 빈도→값 지문 비교) & 대책(표준어 사전·column_value_tree)
 - 배너: 번호 1,372개→281개, 패턴/설명 문구를 값 지문 주제로 갱신(겹침도 완화)

방식(안전·CLAUDE.md §3): Base=V13. slide11.xml만 편집. 본문 쿼드(y∈[1.8M,6.0M)) 제거→새 도형,
배너 텍스트는 도형별 안전 치환(metric-desc는 분할 run을 단일 run으로 병합·서식 보존), fix_clipping.
"""
import zipfile, os, re
import xml.etree.ElementTree as ET

SRC=r"d:\data_intelligence_2\report\데이터지도_분석보고서_V13.pptx"
DST=r"d:\data_intelligence_2\report\데이터지도_분석보고서_V14.pptx"
SLIDE="ppt/slides/slide11.xml"
def EMU(i): return int(round(i*914400))
NAVY="081F46"; GREEN="00965F"; DGRAY="505050"; BODYTX="282D37"; WHITE="FFFFFF"
LIGHT="EEF2F8"; LGREEN="E8F5F0"; LINE="D0D8E4"; RED="C0392B"; ACC="00965F"; FONT="맑은 고딕"
HEADER_H=EMU(0.36)
_id=[3000]
def nid(): _id[0]+=1; return _id[0]
STYLE=('<p:style><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef>'
       '<a:fillRef idx="3"><a:schemeClr val="accent1"/></a:fillRef>'
       '<a:effectRef idx="2"><a:schemeClr val="accent1"/></a:effectRef>'
       '<a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style>')
EMPTY_TX='<p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:endParaRPr/></a:p></p:txBody>'
def esc(s): return s.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")
def roundrect(x,y,w,h,fill,linecol,name="rr"):
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{nid()}" name="{name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
            f'<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>'
            f'<a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>'
            f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>'
            f'<a:ln w="9525"><a:solidFill><a:srgbClr val="{linecol}"/></a:solidFill></a:ln><a:effectLst/></p:spPr>'
            f'{STYLE}{EMPTY_TX}</p:sp>')
def _run(text,sz,b_,color):
    return (f'<a:r><a:rPr sz="{sz}" b="{1 if b_ else 0}" i="0"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
            f'<a:latin typeface="{FONT}"/><a:ea typeface="{FONT}"/></a:rPr><a:t>{esc(text)}</a:t></a:r>')
def _para(runs,algn="l",lnspc=114000,spcaft=14,bullet=False):
    pPr=f'<a:pPr algn="{algn}"><a:lnSpc><a:spcPct val="{lnspc}"/></a:lnSpc><a:spcBef><a:spcPts val="0"/></a:spcBef><a:spcAft><a:spcPts val="{spcaft}"/></a:spcAft>'
    pPr+= '<a:buFont typeface="Arial"/><a:buChar char="•"/>' if bullet else '<a:buNone/>'
    pPr+='</a:pPr>'
    return f'<a:p>{pPr}{"".join(_run(t,sz,bb,c) for (t,sz,bb,c) in runs)}</a:p>'
def tbox(x,y,w,h,paras,anchor="t",lins=12700,name="txt"):
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{nid()}" name="{name}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
            f'<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
            f'<p:txBody><a:bodyPr wrap="square" lIns="{lins}" tIns="6350" rIns="{lins}" bIns="6350" anchor="{anchor}"><a:normAutofit/></a:bodyPr>'
            f'<a:lstStyle/>{"".join(paras)}</p:txBody></p:sp>')
def frame(x,y,w,h,hdr,bg,line,title):
    return [roundrect(x,y,w,h,bg,line,name="g3-bg"),
            roundrect(x,y,w,HEADER_H,hdr,hdr,name="g3-hdr"),
            tbox(x,y,w,HEADER_H,[_para([(title,1100,True,WHITE)],algn="l",spcaft=0)],anchor="ctr",lins=EMU(0.13),name="g3-title")]

# ---------------- BODY: two case cards + method/대책 band ----------------
W=EMU(12.09); LX=EMU(0.62)
CW=EMU(5.93); RXc=LX+CW+EMU(0.23)
CY=EMU(2.0); CH=EMU(2.95)
shapes=[]

# 사례 A
shapes+=frame(LX,CY,CW,CH,NAVY,LIGHT,LINE,"사례 A · 행 일련번호 지문  (min=1 · pk≈1.0)")
shapes.append(tbox(LX+EMU(0.14),CY+HEADER_H+EMU(0.06),CW-EMU(0.26),CH-HEADER_H-EMU(0.12),[
 _para([("이름 (등장 테이블): ",1000,True,NAVY),("연번 168 · 순번 79 · 번호 34",1000,True,BODYTX)],spcaft=9),
 _para([("값 지문: ",1000,True,NAVY),("min_val = 1 · pk_ratio ≈ 1.0 · distinct ≈ 행 수",1000,False,BODYTX)],spcaft=9),
 _para([("읽기: ",1000,True,NAVY),("값이 1부터 행마다 고유 → ‘행 식별자(일련번호)’",1000,False,BODYTX)],spcaft=9),
 _para([("판정: ",1000,True,GREEN),("이름은 3개지만 같은 개념 → 표준 1개 컬럼으로 통합 ",1000,False,BODYTX),("(281개 테이블)",1000,True,RED)],spcaft=0)],
 anchor="t",name="g3-caseA"))

# 사례 B
shapes+=frame(RXc,CY,CW,CH,NAVY,LIGHT,LINE,"사례 B · 데이터 기준일 지문  (pk≈0.0 · 상수)")
shapes.append(tbox(RXc+EMU(0.14),CY+HEADER_H+EMU(0.06),CW-EMU(0.26),CH-HEADER_H-EMU(0.12),[
 _para([("이름 (등장 테이블): ",1000,True,NAVY),("데이터기준일자 169 · 데이터기준일 46 · 기준일자 10",1000,True,BODYTX)],spcaft=9),
 _para([("값 지문: ",1000,True,NAVY),("pk_ratio ≈ 0.0 · distinct_cnt 작음(상당수 1) · 값 = 날짜",1000,False,BODYTX)],spcaft=9),
 _para([("읽기: ",1000,True,NAVY),("전 행이 같은 날짜 → ‘데이터 스냅샷 기준일(거의 상수)’",1000,False,BODYTX)],spcaft=9),
 _para([("근거: ",1000,True,NAVY),("상수 컬럼 최빈 1·2위 = 데이터기준일자(166)·데이터기준일(44)",1000,False,BODYTX)],spcaft=9),
 _para([("판정: ",1000,True,GREEN),("이름 여러 개지만 같은 개념 → 표준 1개로 통합",1000,False,BODYTX)],spcaft=0)],
 anchor="t",name="g3-caseB"))

# 하단 밴드: 방법 & 대책
BY=EMU(5.10); BH=EMU(1.82)
shapes+=frame(LX,BY,W,BH,GREEN,LGREEN,LINE,"값 지문으로 ‘같은 개념’ 자동 판별 — 방법 & 대책")
shapes.append(tbox(LX+EMU(0.14),BY+HEADER_H+EMU(0.05),W-EMU(0.28),BH-HEADER_H-EMU(0.10),[
 _para([("방법: ",1000,True,GREEN),("① col_nm 빈도 집계(이름별 등장 테이블 수) → ② 값 지문(min_val·pk_ratio·distinct) 비교 → 지문이 같으면 동일 개념 (의미 아닌 ‘값’으로 자동 판별)",1000,False,BODYTX)],spcaft=8),
 _para([("· pk_ratio = ",950,True,NAVY),("고유값 수 ÷ 행 수  (1=행마다 고유=식별자,  0=전 행 동일=상수)",950,False,DGRAY)],spcaft=8),
 _para([("대책: ",1000,True,GREEN),("값 지문 + 문자유사도(편집거리·자모)로 표준어 사전 구축 · column_value_tree(값 유사도 군집)로 전수 자동 검증·통합",1000,False,BODYTX)],spcaft=0)],
 anchor="t",name="g3-band"))

# ---------------- banner text updates (safe per-shape) ----------------
def replace_desc_runs(xml, newtext):
    i=xml.index('name="metric-desc"'); start=xml.rindex('<p:sp>',0,i); end=xml.index('</p:sp>',i)+len('</p:sp>')
    sp=xml[start:end]
    rpr=re.search(r'<a:r>(<a:rPr[^>]*/>|<a:rPr.*?</a:rPr>)',sp,re.S)
    rprxml=rpr.group(1) if rpr else '<a:rPr><a:solidFill><a:srgbClr val="282D37"/></a:solidFill></a:rPr>'
    newrun=f'<a:r>{rprxml}<a:t>{esc(newtext)}</a:t></a:r>'
    fr=sp.index('<a:r>'); lr=sp.rindex('</a:r>')+len('</a:r>')
    sp2=sp[:fr]+newrun+sp[lr:]
    return xml[:start]+sp2+xml[end:]

def fix_clipping(xml):
    def repl(m):
        bp=m.group(1)
        bp2=re.sub(r'wrap="[^"]*"','wrap="square"',bp,count=1) if 'wrap="' in bp else bp.replace('<a:bodyPr','<a:bodyPr wrap="square"',1)
        if '<a:spAutoFit/>' in bp2: pass
        elif '<a:normAutofit' in bp2: bp2=re.sub(r'<a:normAutofit[^>]*/>','<a:normAutofit/>',bp2)
        elif '<a:noAutofit/>' in bp2: bp2=bp2.replace('<a:noAutofit/>','<a:normAutofit/>')
        else:
            if bp2.endswith('/>'): bp2=bp2[:-2]+'><a:normAutofit/></a:bodyPr>'
            else: bp2=re.sub(r'(<a:bodyPr[^>]*>)',r'\1<a:normAutofit/>',bp2,count=1)
        return '<p:txBody>'+bp2
    return re.sub(r'<p:txBody>(<a:bodyPr[^>]*/>|<a:bodyPr[^>]*>.*?</a:bodyPr>)',repl,xml,flags=re.DOTALL)

def main():
    with zipfile.ZipFile(SRC) as z:
        order=z.infolist(); data={i.filename:z.read(i.filename) for i in order}
    xml=data[SLIDE].decode("utf-8")

    # 1) banner number 1,372개 -> 281개
    assert xml.count("<a:t>1,372개</a:t>")==1, "number run not unique"
    xml=xml.replace("<a:t>1,372개</a:t>","<a:t>281개</a:t>")
    # 2) banner pattern detail
    xml=xml.replace("<a:t>distinct=1 스캔 · col_nm 빈도 집계</a:t>",
                    "<a:t>값 지문(min_val·pk_ratio)으로 동의어 컬럼 판별</a:t>")
    # 3) banner desc (merge fragmented runs)
    xml=replace_desc_runs(xml,"‘행 일련번호’와 ‘데이터 기준일’이 여러 이름으로 분산 — 값 지문으로 같은 개념을 자동 판별")

    # 4) remove body cards (y in [1.8M,6.0M))
    def rm(m):
        blk=m.group(0); mo=re.search(r'<a:off x="-?\d+" y="(-?\d+)"/>',blk)
        return '' if (mo and 1800000<=int(mo.group(1))<6000000) else blk
    before=xml.count("<p:sp>")
    xml=re.sub(r"<p:sp>.*?</p:sp>",rm,xml,flags=re.DOTALL)
    removed=before-xml.count("<p:sp>")
    assert xml.count("</p:spTree>")==1
    xml=xml.replace("</p:spTree>","".join(shapes)+"</p:spTree>")
    xml=fix_clipping(xml)
    ET.fromstring(xml)
    ids=re.findall(r'<p:cNvPr id="(\d+)"',xml); assert len(ids)==len(set(ids)),"dup id"
    data[SLIDE]=xml.encode("utf-8")
    if os.path.exists(DST): os.remove(DST)
    with zipfile.ZipFile(DST,"w",zipfile.ZIP_DEFLATED) as zout:
        for i in order: zout.writestr(i,data[i.filename])
    print("removed body sp:",removed,"| added:",len(shapes),"| banner updated | ids OK")
    print("WROTE",DST)

if __name__=="__main__":
    main()
