# -*- coding: utf-8 -*-
"""V10 deck: 11페이지(거버넌스 G2 죽은 메트릭) 내용 개선.
요청 반영:
 1) '컬럼 분포 통계'가 무엇인지 명기.
 2) 품질·활용 지표 3종이 무엇인지 상세(null_ratio·table_qty_index·cum_search_num).
 3) '미충전' 표현 교체(→ '값이 계산·적재되지 않음 / 미산출 / 전부 0').
 4) 데이터 예시로 이슈 상세 설명(컬럼 '학년도': 값 통계는 산출, 지표 3종은 0).
 5) 관련 패턴/연관규칙 설명(1D 분포 통계로 발견; 상수컬럼·same_ratio 연결).

데이터 근거(직접 계산): column_map 9,840개 컬럼에서 null_ratio·table_qty_index·cum_search_num = 전부 0
 (distinct_cnt 100%·max_val 96.5%·min_val 73%는 산출됨=선택적 미산출). 예시 컬럼 '학년도' distinct=8·2017~2024,
 지표 3종 0; same_ratio=16,146(평균중복도=네이밍 오류).

방식(안전): Base=V9. slide11.xml만 편집 — 2×2 쿼드(본문 y1.8M~6.0M sp) 제거 후 새 카드 4종 삽입,
배너 설명에서 '미충전' 문구 교체, 마지막에 한글 잘림 방지. 그 외 슬라이드 byte-동일.
"""
import zipfile, os, re

SRC=r"d:\DataInt_2\report\데이터지도_분석보고서_V9.pptx"
DST=r"d:\DataInt_2\report\데이터지도_분석보고서_V10.pptx"
SLIDE="ppt/slides/slide11.xml"
def EMU(i): return int(round(i*914400))
NAVY="081F46"; GREEN="00965F"; DGRAY="505050"; BODYTX="282D37"; WHITE="FFFFFF"
LIGHT="EEF2F8"; LGREEN="E8F5F0"; LINE="D0D8E4"; RED="C0392B"; FONT="맑은 고딕"
HEADER_H=EMU(0.34)

_id=[800]
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
def esc(s): return s.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")
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
    return [roundrect(x,y,w,h,bg,line,name="g2-bg"),
            roundrect(x,y,w,HEADER_H,hdr,hdr,name="g2-hdr"),
            tbox(x,y,w,HEADER_H,[_para([(title,990,True,WHITE)],algn="l",spcaft=0)],anchor="ctr",lins=EMU(0.12),name="g2-title")]

LX=EMU(0.62); LW=EMU(5.78); RX=EMU(6.62); RW=EMU(6.09)
shapes=[]

# A. 컬럼 분포 통계 & 지표 3종
AY=EMU(2.0); AH=EMU(2.42)
shapes+=frame(LX,AY,LW,AH,NAVY,LIGHT,LINE,"‘컬럼 분포 통계’ & 품질·활용 지표 3종")
shapes.append(tbox(LX+EMU(0.12),AY+HEADER_H+EMU(0.03),LW-EMU(0.22),AH-HEADER_H-EMU(0.08),[
 _para([("▸ 컬럼 분포 통계 : ",850,True,NAVY),("9,840개 컬럼의 값 분포를 집계하는 1D 프로파일 스캔(지표별 0·빈값 비율, 고유값 수) → 항상 0인 ‘죽은 지표’ 탐지",850,False,BODYTX)],spcaft=12),
 _para([("▸ 품질·활용 지표 3종 :",850,True,NAVY)],spcaft=6),
 _para([("· null_ratio ",840,True,NAVY),("— 결측률(품질): 값 중 빈(null)의 비율",840,False,BODYTX)],spcaft=4),
 _para([("· table_qty_index ",840,True,NAVY),("— 품질지수(품질): 종합 품질 점수",840,False,BODYTX)],spcaft=4),
 _para([("· cum_search_num ",840,True,NAVY),("— 누적 검색수(활용): 데이터셋 조회·이용 횟수",840,False,BODYTX)],spcaft=0)],
 anchor="t",name="g2-terms"))

# B. 예시 데이터
BY=EMU(4.52); BH=EMU(2.38)
shapes+=frame(LX,BY,LW,BH,NAVY,LIGHT,LINE,"예시 데이터 — 죽은 지표")
by=BY+HEADER_H+EMU(0.03)
shapes.append(tbox(LX+EMU(0.12),by,LW-EMU(0.24),EMU(0.24),[_para([("예시 컬럼: ",840,True,NAVY),("‘학년도’ (교육 데이터셋)",840,False,BODYTX)],spcaft=0)],anchor="t",name="g2-ex-sub"))
ROWS=[
 ("distinct_cnt (고유값)", [("8",870,True,BODYTX),("   ✓ 산출됨",820,True,GREEN)]),
 ("min ~ max", [("2017 ~ 2024",870,False,BODYTX),("   ✓ 산출됨",820,True,GREEN)]),
 ("null_ratio (결측률)", [("0",870,True,RED),("   ✗ 미산출",820,True,RED)]),
 ("table_qty_index (품질)", [("0",870,True,RED),("   ✗ 미산출",820,True,RED)]),
 ("cum_search_num (활용)", [("0",870,True,RED),("   ✗ 미산출",820,True,RED)]),
]
ry=by+EMU(0.28); RH=EMU(0.285); fw=EMU(2.45)
for i,(fld,val) in enumerate(ROWS):
    yy=ry+i*RH
    shapes.append(tbox(LX+EMU(0.16),yy,fw,RH,[_para([(fld,830,True,NAVY)],spcaft=0)],anchor="ctr",lins=6350,name=f"g2-f{i}"))
    shapes.append(tbox(LX+EMU(0.16)+fw,yy,LW-EMU(0.32)-fw,RH,[_para(val,spcaft=0)],anchor="ctr",lins=6350,name=f"g2-v{i}"))
shapes.append(tbox(LX+EMU(0.14),ry+5*RH+EMU(0.01),LW-EMU(0.28),EMU(0.40),[
 _para([("※ 같은 컬럼에서 값 통계(고유값·최소·최대)는 채워졌으나 품질·활용 지표만 0 → 선택적 미산출",800,False,DGRAY)],spcaft=0)],anchor="t",name="g2-ex-note"))

# C. 의미 분석 (구체)
CY=EMU(2.0); CH=EMU(2.90)
shapes+=frame(RX,CY,RW,CH,NAVY,LIGHT,LINE,"② 의미 분석 (구체)")
cb=CY+HEADER_H+EMU(0.05)
clines=[
 [("① 값: ",870,True,NAVY),("9,840개 컬럼 전부 null_ratio=0 · table_qty_index=0 · cum_search_num=0 (값이 계산·적재되지 않음)",870,False,BODYTX)],
 [("② 비현실성: ",870,True,NAVY),("공공데이터에 결측률 0%는 불가능 → 실제 계산 없이 0으로만 채워진 ‘죽은 지표’",870,False,BODYTX)],
 [("③ 선택적 미산출: ",870,True,NAVY),("같은 행에서 distinct·min·max는 산출됐는데 이 3종(+avg·사분위)만 0/빈값",870,False,BODYTX)],
 [("④ 영향: ",870,True,NAVY),("품질(완전성)·활용(이용도) 축으로 데이터셋 우선순위·정제 대상 선정 불가",870,False,BODYTX)],
]
shapes.append(tbox(RX+EMU(0.12),cb,RW-EMU(0.22),CH-HEADER_H-EMU(0.10),[_para(l,spcaft=40) for l in clines],anchor="t",name="g2-mean"))

# D. 관련 패턴 · 대책
DY=EMU(5.0); DH=EMU(1.90)
shapes+=frame(RX,DY,RW,DH,GREEN,LGREEN,LINE,"관련 패턴 · 대책")
db=DY+HEADER_H+EMU(0.05)
dlines=[
 [("발견 패턴: ",860,True,NAVY),("연관규칙(동시출현)이 아니라 1D 분포 통계(%0=100%)로 포착. 3종은 distinct=1(상수) → 상수컬럼 1,372개(13.9%)의 일부, 속성규칙 ‘상수컬럼=있음’과 연결",860,False,BODYTX)],
 [("유사 오류: ",860,True,GREEN),("same_ratio도 ‘비율’ 아님 → 평균중복도(rec/distinct, 예 학년도=16,146) 네이밍 오류(G6)",860,False,BODYTX)],
 [("대책: ",860,True,GREEN),("지표 산출 파이프라인 재적재 · 검색로그 연동(cum_search_num) · distinct_value로 null_ratio 실재 계산 · 정의-구현 정합성 감사",860,False,BODYTX)],
]
shapes.append(tbox(RX+EMU(0.12),db,RW-EMU(0.22),DH-HEADER_H-EMU(0.10),[_para(l,spcaft=26) for l in dlines],anchor="t",name="g2-action"))

# banner desc: replace '미충전' wording
OLD_DESC="카탈로그 품질·활용 지표가 전부 미충전되어 분석축으로 쓸 수 없음"
NEW_DESC="품질·활용 지표 3종이 모든 컬럼에서 값 0 — 계산·적재되지 않아 해당 축으로 분석 불가"

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
    # update banner desc (remove '미충전')
    assert f"<a:t>{OLD_DESC}</a:t>" in xml, "banner desc not found"
    xml=xml.replace(f"<a:t>{OLD_DESC}</a:t>",f"<a:t>{NEW_DESC}</a:t>")
    # remove 2x2 quad body
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
    print("removed quad sp:",removed,"| added shapes:",len(shapes),"| banner desc updated")
    print("WROTE",DST)

if __name__=="__main__":
    main()
