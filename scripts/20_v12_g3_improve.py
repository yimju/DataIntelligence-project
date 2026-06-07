# -*- coding: utf-8 -*-
"""V12 deck: 11페이지(거버넌스 G3 컬럼 표준·품질 = 파일 slide12.xml, V11에서 위치 11로 당겨짐) 개선.
요청 반영:
 1) '동일 개념'이 수치적으로 무슨 의미인지: 이름은 달라도 '값 지문'이 같으면 동일 개념.
    예) 연번(169)·순번(81)·번호(34) — 셋 다 min_val=1, pk_ratio 0.96~0.99 = 동일 일련번호 지문.
 2) distinct=1 스캔 / col_nm 빈도 집계 '절차'를 구체적으로(정규화→value_counts→distinct 스캔→정합검사).
 3) 예시 데이터로 구체화(변형 이름 빈도·상수컬럼·스키마 불일치 실제 예).
 4) 의미 분석을 구체 수치·예시로(표준화 비용·동일성 근거·상수컬럼·회계 불일치).

데이터 근거(직접 계산): 일련번호 변형 연번169/순번81/번호34(min=1·pk≈0.99); 기준일 데이터기준일자177/데이터기준일49/기준일자12;
 상수컬럼 distinct=1 → 1,372개(13.9%), 예 '삭제여부 디폴트 N'(전 행 N); 스키마 불일치 56개(예 상주시 투표정보 8≠1+3).

방식(안전): Base=V11. slide12.xml만 편집(2×2 쿼드 제거→새 카드 4종), 한글잘림 방지. 그 외 byte-동일.
"""
import zipfile, os, re

SRC=r"d:\DataInt_2\report\데이터지도_분석보고서_V11.pptx"
DST=r"d:\DataInt_2\report\데이터지도_분석보고서_V12.pptx"
SLIDE="ppt/slides/slide12.xml"   # = 11페이지(G3)
def EMU(i): return int(round(i*914400))
NAVY="081F46"; GREEN="00965F"; DGRAY="505050"; BODYTX="282D37"; WHITE="FFFFFF"
LIGHT="EEF2F8"; LGREEN="E8F5F0"; LINE="D0D8E4"; RED="C0392B"; FONT="맑은 고딕"
HEADER_H=EMU(0.34)

_id=[900]
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
def _para(runs,algn="l",lnspc=114000,spcaft=18):
    rs="".join(_run(t,sz,bb,c) for (t,sz,bb,c) in runs)
    return (f'<a:p><a:pPr algn="{algn}"><a:lnSpc><a:spcPct val="{lnspc}"/></a:lnSpc>'
            f'<a:spcBef><a:spcPts val="0"/></a:spcBef><a:spcAft><a:spcPts val="{spcaft}"/></a:spcAft></a:pPr>{rs}</a:p>')
def tbox(x,y,w,h,paras,anchor="t",lins=12700,name="txt"):
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{nid()}" name="{name}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
            f'<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
            f'<p:txBody><a:bodyPr wrap="square" lIns="{lins}" tIns="6350" rIns="{lins}" bIns="6350" anchor="{anchor}"><a:normAutofit/></a:bodyPr>'
            f'<a:lstStyle/>{"".join(paras)}</p:txBody></p:sp>')
def frame(x,y,w,h,hdr,bg,line,title):
    return [roundrect(x,y,w,h,bg,line,name="g3-bg"),
            roundrect(x,y,w,HEADER_H,hdr,hdr,name="g3-hdr"),
            tbox(x,y,w,HEADER_H,[_para([(title,980,True,WHITE)],algn="l",spcaft=0)],anchor="ctr",lins=EMU(0.12),name="g3-title")]

LX=EMU(0.62); LW=EMU(5.78); RX=EMU(6.62); RW=EMU(6.09)
shapes=[]

# A. 동일 개념 의미(수치) + 탐지 절차
AY=EMU(2.0); AH=EMU(2.42)
shapes+=frame(LX,AY,LW,AH,NAVY,LIGHT,LINE,"‘동일 개념·다른 이름’의 수치적 의미 · 탐지 절차")
shapes.append(tbox(LX+EMU(0.12),AY+HEADER_H+EMU(0.03),LW-EMU(0.22),AH-HEADER_H-EMU(0.08),[
 _para([("▸ 동일 개념 = ",845,True,NAVY),("이름은 달라도 ‘값 지문’(값 범위·유일성)이 같은 컬럼",845,False,BODYTX)],spcaft=8),
 _para([("예) 일련번호 = ",840,True,NAVY),("연번·순번·번호 — 셋 다 min_val=1, pk_ratio 0.96~0.99 (1부터의 거의-유일) = 같은 지문, 다른 이름",840,False,BODYTX)],spcaft=10),
 _para([("▸ 탐지 절차 :",845,True,NAVY)],spcaft=4),
 _para([("① col_nm 정규화(괄호·공백·숫자접두 제거) → ② value_counts로 이름별 등장 테이블 수 집계",835,False,BODYTX)],spcaft=4),
 _para([("③ distinct_cnt=1 스캔(상수=정보량 0) · ④ column_cnt vs (범주+수치) 정합 검사",835,False,BODYTX)],spcaft=0)],
 anchor="t",name="g3-terms"))

# B. 예시 데이터
BY=EMU(4.52); BH=EMU(2.38)
shapes+=frame(LX,BY,LW,BH,NAVY,LIGHT,LINE,"예시 데이터")
bb=BY+HEADER_H+EMU(0.04)
shapes.append(tbox(LX+EMU(0.12),bb,LW-EMU(0.22),BH-HEADER_H-EMU(0.08),[
 _para([("① 같은 개념, 다른 이름 ",840,True,NAVY),("(등장 테이블 수)",800,False,DGRAY)],spcaft=4),
 _para([("· 일련번호: ",840,True,NAVY),("연번 169 · 순번 81 · 번호 34  (모두 min=1·pk≈0.99)",840,False,BODYTX)],spcaft=3),
 _para([("· 기준일: ",840,True,NAVY),("데이터기준일자 177 · 데이터기준일 49 · 기준일자 12",840,False,BODYTX)],spcaft=8),
 _para([("② 상수 컬럼(distinct=1) — 1,372개(13.9%)",840,True,NAVY)],spcaft=3),
 _para([("· 예) ",840,False,BODYTX),("‘삭제여부 디폴트 N’ → 전 행 ‘N’",840,True,RED),(" (정보량 0)",820,False,DGRAY)],spcaft=8),
 _para([("③ 스키마 회계 불일치 — 56개 테이블",840,True,NAVY)],spcaft=3),
 _para([("· 예) ",840,False,BODYTX),("상주시 투표정보 column_cnt=8 ≠ 범주1+수치3",840,True,RED)],spcaft=0)],
 anchor="t",name="g3-ex"))

# C. 의미 분석 (구체)
CY=EMU(2.0); CH=EMU(2.90)
shapes+=frame(RX,CY,RW,CH,NAVY,LIGHT,LINE,"② 의미 분석 (구체)")
cb=CY+HEADER_H+EMU(0.05)
clines=[
 [("① 표준 부재 비용: ",865,True,NAVY),("일련번호 1개 개념이 3개 이름으로 분산(연번169+순번81+번호34=284테이블), 기준일은 4개 이름 → 표준화 시 각 1컬럼으로 통합",865,False,BODYTX)],
 [("② 동일성 근거: ",865,True,NAVY),("이름이 달라도 값 지문 동일(min=1·pk≈0.99) → 사람 판단 없이 자동 통합 가능 신호",865,False,BODYTX)],
 [("③ 상수 컬럼: ",865,True,NAVY),("distinct=1인 1,372개(13.9%)는 전 행이 같은 값(예 ‘N’)=정보량 0; 최빈 데이터기준일자(166)·데이터기준일(45)=스냅샷 날짜 단일",865,False,BODYTX)],
 [("④ 스키마 불일치: ",865,True,NAVY),("56개 테이블 column_cnt ≠ 범주+수치(예 8≠4) → 미분류 컬럼 존재 = 메타 회계 오류",865,False,BODYTX)],
]
shapes.append(tbox(RX+EMU(0.12),cb,RW-EMU(0.22),CH-HEADER_H-EMU(0.10),[_para(l,spcaft=36) for l in clines],anchor="t",name="g3-mean"))

# D. 관련 패턴 · 대책
DY=EMU(5.0); DH=EMU(1.90)
shapes+=frame(RX,DY,RW,DH,GREEN,LGREEN,LINE,"관련 패턴 · 대책")
db=DY+HEADER_H+EMU(0.05)
dlines=[
 [("발견 패턴: ",855,True,NAVY),("연관규칙 아닌 1D 빈도·분포 집계(col_nm value_counts·distinct 분포). 값 지문 일치는 column_value_tree(값 유사도 군집, 최대 189컬럼/군집)로 자동 확인 가능",855,False,BODYTX)],
 [("대책: ",855,True,GREEN),("문자유사도(편집거리·자모)+값지문으로 표준어 사전 구축 · 상수/중복 컬럼 정리 · 스키마 검증 룰(컬럼수 정합) 적용",855,False,BODYTX)],
]
shapes.append(tbox(RX+EMU(0.12),db,RW-EMU(0.22),DH-HEADER_H-EMU(0.10),[_para(l,spcaft=24) for l in dlines],anchor="t",name="g3-action"))

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
        order=z.infolist(); data={i.filename:z.read(i.filename) for i in order}
    xml=data[SLIDE].decode("utf-8")
    def rm(m):
        blk=m.group(0); mo=re.search(r'<a:off x="\d+" y="(\d+)"/>',blk)
        return '' if (mo and 1800000<=int(mo.group(1))<6000000) else blk
    before=xml.count("<p:sp>")
    xml=re.sub(r"<p:sp>.*?</p:sp>",rm,xml,flags=re.DOTALL)
    removed=before-xml.count("<p:sp>")
    assert xml.count("</p:spTree>")==1
    xml=xml.replace("</p:spTree>","".join(shapes)+"</p:spTree>")
    xml=fix_clipping(xml)
    data[SLIDE]=xml.encode("utf-8")
    if os.path.exists(DST): os.remove(DST)
    with zipfile.ZipFile(DST,"w",zipfile.ZIP_DEFLATED) as zout:
        for i in order: zout.writestr(i, data[i.filename])
    print("removed quad sp:",removed,"| added shapes:",len(shapes))
    print("WROTE",DST)

if __name__=="__main__":
    main()
