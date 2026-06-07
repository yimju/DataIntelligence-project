# -*- coding: utf-8 -*-
"""V7 deck: 10페이지(거버넌스 G1 출처 결손) 내용 개선.
요청 반영:
 1) placeholder 의미 설명 추가('소스'=출처 칸을 채운 더미값).
 2) A1 비자명 연관규칙 의미 설명 추가(cate1↔cate2 자명규칙을 뺀 진짜 패턴).
 3) ‘의미 분석’을 더 구체적으로(누락필드·계통성·단일배치·동형템플릿·영향/복구).
 4) 예시 데이터로 어떤 값이 누락됐는지 자세히(table_source='소스', url='url'; 파일명엔 기관명 잔존).

근거(데이터 계산): cate2='유아및초·중등교육' 100건이 모두 table_source='소스'·url='url'·regi_date=2026-03-30(단일 배치),
table_nm엔 기관명 95% 잔존(예: '광주광역시 남구_어린이집 현황_20230712.csv'). 보고서 §0.1·§2와 일치.

방식(안전): Base=V6(사용자가 8p 수정한 최신본). slide10.xml만 편집 — 기존 2×2 쿼드(본문 y 1.8M~6.0M sp) 제거 후
새 카드 삽입(문자열 편집 + python-pptx 검증 도형구조 <p:style>·빈 txBody·별도 텍스트박스). 마지막에 한글잘림 방지 적용.
다른 슬라이드는 V6와 byte-동일.
"""
import zipfile, os, re

SRC=r"d:\DataInt_2\report\데이터지도_분석보고서_V6.pptx"
DST=r"d:\DataInt_2\report\데이터지도_분석보고서_V7.pptx"
SLIDE="ppt/slides/slide10.xml"

def EMU(i): return int(round(i*914400))
NAVY="081F46"; GREEN="00965F"; DGRAY="505050"; BODYTX="282D37"; WHITE="FFFFFF"
LIGHT="EEF2F8"; LGREEN="E8F5F0"; LINE="D0D8E4"; RED="C0392B"; FONT="맑은 고딕"
HEADER_H=EMU(0.34)

# ---------- shape builders (proven structure) ----------
_id=[500]
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
def _para(runs,algn="l",lnspc=118000,spcaft=20):
    rs="".join(_run(t,sz,bb,c) for (t,sz,bb,c) in runs)
    return (f'<a:p><a:pPr algn="{algn}"><a:lnSpc><a:spcPct val="{lnspc}"/></a:lnSpc>'
            f'<a:spcBef><a:spcPts val="0"/></a:spcBef><a:spcAft><a:spcPts val="{spcaft}"/></a:spcAft></a:pPr>{rs}</a:p>')
def tbox(x,y,w,h,paras,anchor="t",lins=12700,name="txt"):
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{nid()}" name="{name}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
            f'<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
            f'<p:txBody><a:bodyPr wrap="square" lIns="{lins}" tIns="6350" rIns="{lins}" bIns="6350" anchor="{anchor}"><a:normAutofit/></a:bodyPr>'
            f'<a:lstStyle/>{"".join(paras)}</p:txBody></p:sp>')
def card_frame(x,y,w,h,hdr,bg,line,title):
    return [roundrect(x,y,w,h,bg,line,name="g1-bg"),
            roundrect(x,y,w,HEADER_H,hdr,hdr,name="g1-hdr"),
            tbox(x,y,w,HEADER_H,[_para([(title,1000,True,WHITE)],algn="l",spcaft=0)],anchor="ctr",lins=EMU(0.12),name="g1-title")]

# ---------- layout ----------
LX=EMU(0.62); LW=EMU(5.78)
RX=EMU(6.62); RW=EMU(6.09)
shapes=[]

# (1) 용어 설명 (left-top)
shapes+=card_frame(LX,EMU(2.0),LW,EMU(1.66),NAVY,LIGHT,LINE,"용어 설명")
shapes.append(tbox(LX+EMU(0.10),EMU(2.0)+HEADER_H+EMU(0.04),LW-EMU(0.20),EMU(1.66)-HEADER_H-EMU(0.08),[
    _para([("▸ placeholder('소스') : ",880,True,NAVY),
           ("출처(table_source) 칸에 실제 기관명 대신 채운 더미값. 빈칸(null)이 아니라 ‘소스’라는 글자로 채워 누락을 숨김 → 자동 결측검사에 안 걸림",880,False,BODYTX)]),
    _para([("▸ A1 비자명 연관규칙 : ",880,True,NAVY),
           ("속성(A) 규칙 중 cate1↔cate2의 당연한 계층종속(자명=A2)을 뺀 ‘진짜’ 규칙. 예) {유아및초·중등교육}→{기관=미상} lift 6.89·conf 1.0",880,False,BODYTX)])],
    anchor="t",name="g1-terms"))

# (2) 예시 데이터 — 무엇이 누락됐나 (left-bottom)
EY=EMU(3.78); EH=EMU(3.12)
shapes+=card_frame(LX,EY,LW,EH,NAVY,LIGHT,LINE,"예시 데이터 — 무엇이 누락됐나")
by=EY+HEADER_H+EMU(0.04)
shapes.append(tbox(LX+EMU(0.12),by,LW-EMU(0.24),EMU(0.40),[
    _para([("예시 1건:  ",850,True,NAVY),("광주광역시 남구_어린이집 현황_20230712.csv",850,False,BODYTX)],spcaft=0)],anchor="t",name="g1-ex-sub"))
# field | value rows
ROWS=[
 ("table_source (출처)", [("‘소스’",870,True,RED),("   ✗ placeholder (기관명 소실)",820,False,RED)]),
 ("url (원본 링크)",     [("‘url’",870,True,RED),("   ✗ placeholder",820,False,RED)]),
 ("table_nm (파일명)",  [("‘광주광역시 남구…’",870,False,BODYTX),("   ✓ 기관명 존재",820,True,GREEN)]),
 ("regi_date (등록)",   [("2026-03-30",870,False,BODYTX),("   (100건 전부 동일=단일 배치)",820,False,DGRAY)]),
 ("src_platform",      [("공공데이터포털",870,False,BODYTX),("   ✓ 정상",820,True,GREEN)]),
]
ry=by+EMU(0.44); RH=EMU(0.345); fcol_w=EMU(2.05)
for i,(fld,val) in enumerate(ROWS):
    yy=ry+i*RH
    shapes.append(tbox(LX+EMU(0.16),yy,fcol_w,RH,[_para([(fld,840,True,NAVY)],spcaft=0)],anchor="ctr",lins=6350,name=f"g1-f{i}"))
    shapes.append(tbox(LX+EMU(0.16)+fcol_w,yy,LW-EMU(0.32)-fcol_w,RH,[_para(val,spcaft=0)],anchor="ctr",lins=6350,name=f"g1-v{i}"))
shapes.append(tbox(LX+EMU(0.14),ry+5*RH+EMU(0.02),LW-EMU(0.28),EMU(0.42),[
    _para([("※ ‘소스’·‘url’ 두 메타필드만 더미로 덮임. 파일명엔 기관명 잔존(95%) → 출처 복구 가능",810,False,DGRAY)],spcaft=0)],anchor="t",name="g1-ex-note"))

# (3) ② 의미 분석 (right-top, 구체)
MY=EMU(2.0); MH=EMU(2.98)
shapes+=card_frame(RX,MY,RW,MH,NAVY,LIGHT,LINE,"② 의미 분석 (구체)")
mb=MY+HEADER_H+EMU(0.05)
mean_lines=[
 [("① 누락 값: ",880,True,NAVY),("table_source→‘소스’, url→‘url’ — 두 메타필드가 더미로 덮임",880,False,BODYTX)],
 [("② 계통성: ",880,True,NAVY),("placeholder 145건 중 100건이 전부 ‘유아및초·중등교육’(이 분류 100% 소실, 타 분류 1~15%)",880,False,BODYTX)],
 [("③ 단일 배치: ",880,True,NAVY),("100건 모두 regi_date=2026-03-30 동일 → 한 번의 적재에서 일괄 placeholder 처리(우연 아님)",880,False,BODYTX)],
 [("④ 동형 템플릿: ",880,True,NAVY),("유일키·상수컬럼·컬럼폭=표준 동반(conf 0.79~0.84) → 같은 양식으로 일괄 적재",880,False,BODYTX)],
 [("⑤ 영향: ",880,True,NAVY),("출처 기반 신뢰성·책임추적(provenance) 붕괴. 단, 파일명에 기관명 잔존 → 복구 가능",880,False,BODYTX)],
]
shapes.append(tbox(RX+EMU(0.12),mb,RW-EMU(0.22),MH-HEADER_H-EMU(0.10),[_para(l,spcaft=40) for l in mean_lines],anchor="t",name="g1-mean"))

# (4) ① 처리 · ③ 대책 · ④ 추가분석 (right-bottom, compact)
PY=EMU(5.08); PH=EMU(1.82)
shapes+=card_frame(RX,PY,RW,PH,GREEN,LGREEN,LINE,"① 처리 · ③ 대책 · ④ 추가 분석")
pb=PY+HEADER_H+EMU(0.05)
proc_lines=[
 [("처리: ",870,True,NAVY),("AOI로 ‘소스’=미상 식별 → cate2별 교차표 → A1 규칙으로 결합 포착(lift 6.89)",870,False,BODYTX)],
 [("대책: ",870,True,GREEN),("①파일명에서 출처 복구(1순위) ②적재 시 provenance 필수검증(placeholder 차단) ③출처 표준코드",870,False,BODYTX)],
 [("추가: ",870,True,GREEN),("regi_date·url로 배치 역추적 · table_str_tree 동형군집 확정 · 타 분류 결손율 모니터링",870,False,BODYTX)],
]
shapes.append(tbox(RX+EMU(0.12),pb,RW-EMU(0.22),PH-HEADER_H-EMU(0.10),[_para(l,spcaft=30) for l in proc_lines],anchor="t",name="g1-action"))

# ---------- no-clip (apply to whole slide at end) ----------
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
    # remove existing 2x2 quad (body sp with 1.8M<=y<=6.0M); keep header/page
    def rm(m):
        blk=m.group(0); mo=re.search(r'<a:off x="\d+" y="(\d+)"/>',blk)
        return '' if (mo and 1800000<=int(mo.group(1))<=6000000) else blk
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
