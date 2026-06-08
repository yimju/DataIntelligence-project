# -*- coding: utf-8 -*-
"""V13: 11페이지(컬럼 표준·품질 = on-disk slide11.xml) 본문 개선.
사용자 요청(6) 반영 — 친절한 과정 설명 + 예시:
 1) distinct_cnt 정의와 =1(상수=정보량 0) 의미
 2) col_nm 빈도 집계 기준 = '컬럼 이름(문자열)' value_counts → 그 이름을 가진 컬럼(테이블) 수
 3) 집계 결과 예시(이름별 등장 테이블 수 + 값 지문)
 4) 동일 개념 비교 = '값 지문'(min_val·pk_ratio·distinct) 비교(의미 아닌 값으로)
 5) column_cnt =? (categorical+numerical) 정합성 검사 방법 + 예시
 6) 의미분석에 나열된 컬럼·숫자의 뜻 해설

데이터 근거(scripts/_g3_examples.py 직접계산):
 일련번호 연번168·순번79·번호34(min=1·pk중앙1.0) / 기준일 데이터기준일자169·데이터기준일46·기준일자10(min=1·pk≈0.0)
 상수 distinct=1 → 1,372개(13.9%), 최빈 데이터기준일자166·데이터기준일44, 실제값 예 데이터기준일자='2024-11-30'·기관명='이천시의회'
 스키마 불일치 56개, 예 경상북도 상주시 투표정보 column_cnt 8 ≠ 범주1+수치3

방식(안전·CLAUDE.md §3): Base=on-disk V12. slide11.xml만 편집(본문 쿼드 y∈[1.8M,6.0M) 제거→새 카드 4종),
배너/제목/페이지번호 보존, 마지막에 fix_clipping. 그 외 파트 byte-동일.
"""
import zipfile, os, re

SRC=r"d:\data_intelligence_2\report\데이터지도_분석보고서_V12.pptx"
DST=r"d:\data_intelligence_2\report\데이터지도_분석보고서_V13.pptx"
SLIDE="ppt/slides/slide11.xml"   # = 11페이지 (on-disk 검증 완료)
def EMU(i): return int(round(i*914400))
NAVY="081F46"; GREEN="00965F"; DGRAY="505050"; BODYTX="282D37"; WHITE="FFFFFF"
LIGHT="EEF2F8"; LGREEN="E8F5F0"; LINE="D0D8E4"; RED="C0392B"; FONT="맑은 고딕"
HEADER_H=EMU(0.34)
_id=[2000]
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
def _para(runs,algn="l",lnspc=112000,spcaft=14):
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
            tbox(x,y,w,HEADER_H,[_para([(title,1000,True,WHITE)],algn="l",spcaft=0)],anchor="ctr",lins=EMU(0.12),name="g3-title")]

LX=EMU(0.62); LW=EMU(5.82); RX=EMU(6.66); RW=EMU(6.05)
shapes=[]

# ① 분석 과정 (LEFT-TOP)
AY=EMU(2.0); AH=EMU(2.36)
shapes+=frame(LX,AY,LW,AH,NAVY,LIGHT,LINE,"① 분석 과정 — 무엇을 어떻게 집계했나")
shapes.append(tbox(LX+EMU(0.13),AY+HEADER_H+EMU(0.04),LW-EMU(0.24),AH-HEADER_H-EMU(0.10),[
 _para([("distinct_cnt = ",1000,True,NAVY),("한 컬럼의 ‘서로 다른 값’ 개수. ",1000,False,BODYTX),("=1이면 전 행이 같은 값 → 정보량 0(상수)",1000,True,RED)],spcaft=7),
 _para([("col_nm 빈도 = ",1000,True,NAVY),("‘컬럼 이름(문자열)’ 기준 value_counts → 그 이름을 가진 컬럼(=테이블) 수",1000,False,BODYTX)],spcaft=7),
 _para([("동일 개념 비교 = ",1000,True,NAVY),("이름이 달라도 ‘값 지문’(min_val·pk_ratio·distinct)이 같으면 동일으로 판정(의미 아닌 값으로)",1000,False,BODYTX)],spcaft=7),
 _para([("정합성 검사 = ",1000,True,NAVY),("테이블별 column_cnt =? 범주+수치, 다르면 미분류 컬럼 존재",1000,False,BODYTX)],spcaft=0)],
 anchor="t",name="g3-proc"))

# ② 집계 결과 예시 (LEFT-BOTTOM)
BY=EMU(4.46); BH=EMU(2.46)
shapes+=frame(LX,BY,LW,BH,NAVY,LIGHT,LINE,"② 집계 결과 예시 — 등장 테이블 수 · 값 지문")
shapes.append(tbox(LX+EMU(0.13),BY+HEADER_H+EMU(0.04),LW-EMU(0.24),BH-HEADER_H-EMU(0.10),[
 _para([("▸ 일련번호 지문 ",1000,True,NAVY),("(min=1 · pk≈1.0 = 1부터 거의 유일)",900,False,DGRAY)],spcaft=2),
 _para([("   연번 168 · 순번 79 · 번호 34",1000,True,BODYTX),("  → 이름만 다름",1000,False,DGRAY)],spcaft=7),
 _para([("▸ 스냅샷날짜 지문 ",1000,True,NAVY),("(min=1 · pk≈0.0 = 거의 상수)",900,False,DGRAY)],spcaft=2),
 _para([("   데이터기준일자 169 · 데이터기준일 46 · 기준일자 10",1000,True,BODYTX)],spcaft=7),
 _para([("▸ 상수값 예: ",1000,True,NAVY),("데이터기준일자=‘2024-11-30’ · 기관명=‘이천시의회’",1000,False,BODYTX)],spcaft=5),
 _para([("▸ 스키마 불일치 예: ",1000,True,NAVY),("상주시 투표정보 column_cnt 8 ≠ 범주1+수치3",1000,True,RED)],spcaft=0)],
 anchor="t",name="g3-ex"))

# ③ 의미 분석 (RIGHT-TOP)
CY=EMU(2.0); CH=EMU(2.92)
shapes+=frame(RX,CY,RW,CH,NAVY,LIGHT,LINE,"③ 의미 분석 — 나열된 컬럼·숫자의 뜻")
shapes.append(tbox(RX+EMU(0.13),CY+HEADER_H+EMU(0.05),RW-EMU(0.24),CH-HEADER_H-EMU(0.10),[
 _para([("연번168+순번79+번호34 = 281테이블 : ",1000,True,NAVY),("한 개념(행 일련번호)이 3개 이름으로 분산 → 표준화 시 1개 표준 컬럼으로 통합",1000,False,BODYTX)],spcaft=9),
 _para([("pk≈1.0·min=1 (값 지문) : ",1000,True,NAVY),("세 이름 모두 ‘1부터 거의-유일 정수’ = 사람 판단 없이 자동 통합 가능 신호",1000,False,BODYTX)],spcaft=9),
 _para([("1,372개(13.9%) : ",1000,True,NAVY),("상수 컬럼 수. 최빈=데이터기준일자166·데이터기준일44 → 데이터셋 전체가 단일 스냅샷 날짜만 보유(정보량 0)",1000,False,BODYTX)],spcaft=9),
 _para([("56개 : ",1000,True,NAVY),("column_cnt ≠ 범주+수치인 테이블 = 메타데이터 회계 오류(미분류 컬럼 존재)",1000,False,BODYTX)],spcaft=0)],
 anchor="t",name="g3-mean"))

# ④ 대책 · 추가 분석 (RIGHT-BOTTOM)
DY=EMU(5.02); DH=EMU(1.90)
shapes+=frame(RX,DY,RW,DH,GREEN,LGREEN,LINE,"④ 대책 · 추가 분석")
shapes.append(tbox(RX+EMU(0.13),DY+HEADER_H+EMU(0.05),RW-EMU(0.24),DH-HEADER_H-EMU(0.10),[
 _para([("표준어 사전: ",1000,True,GREEN),("문자유사도(편집거리·자모) + 값 지문(min·pk)으로 동의어 컬럼 자동 군집",1000,False,BODYTX)],spcaft=8),
 _para([("정리·검증: ",1000,True,GREEN),("상수·중복 컬럼 정리 · 스키마 검증 룰(컬럼수 정합) 적용",1000,False,BODYTX)],spcaft=8),
 _para([("전수 확인: ",1000,True,GREEN),("값 지문 일치는 column_value_tree(값 유사도 군집)로 자동 검증",1000,False,BODYTX)],spcaft=0)],
 anchor="t",name="g3-action"))

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
        blk=m.group(0); mo=re.search(r'<a:off x="-?\d+" y="(-?\d+)"/>',blk)
        return '' if (mo and 1800000<=int(mo.group(1))<6000000) else blk
    before=xml.count("<p:sp>")
    xml=re.sub(r"<p:sp>.*?</p:sp>",rm,xml,flags=re.DOTALL)
    removed=before-xml.count("<p:sp>")
    assert xml.count("</p:spTree>")==1, "spTree anomaly"
    xml=xml.replace("</p:spTree>","".join(shapes)+"</p:spTree>")
    xml=fix_clipping(xml)
    # validate well-formed + unique ids
    import xml.etree.ElementTree as ET
    ET.fromstring(xml)
    ids=re.findall(r'<p:cNvPr id="(\d+)"',xml); assert len(ids)==len(set(ids)), "dup cNvPr id"
    data[SLIDE]=xml.encode("utf-8")
    if os.path.exists(DST): os.remove(DST)
    with zipfile.ZipFile(DST,"w",zipfile.ZIP_DEFLATED) as zout:
        for i in order: zout.writestr(i, data[i.filename])
    print("removed body sp:",removed,"| added shapes:",len(shapes),"| ids unique OK")
    print("WROTE",DST)

if __name__=="__main__":
    main()
