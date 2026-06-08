# -*- coding: utf-8 -*-
"""V16: 11페이지(=slide15.xml, 교차도메인 재식별 위험) 개선.
요청:
 1) 의료기관 패밀리가 어떻게 만들어졌는지 설명(장바구니→FP-Growth 최대빈발)
 2) PII 후보가 무엇인지 — 정의·의미·예시 추가
 3) 어떤 연관규칙인지 도표 추가 (10페이지 chip 도식 스타일 참고: [선행] ▶ [후행])

데이터 근거(I1_synthesis·I1_metrics):
 의료기관 패밀리 {순번·의료기관명·의료기관전화번호·의료기관주소} = 11개 보건 테이블 공통(FP-Growth 최대빈발 ≥8)
 규칙 의료기관전화번호 → {의료기관명, 의료기관주소}: support 15테이블·conf 0.882·lift 46.39·conviction 8.34
 PII 후보(교차도메인 다출현 식별자) 예: 전화번호140·주소85·위도44·경도44·의료기관명39·연락처27

방식(안전): Base=V15. slide15.xml만 편집(본문 2×2 제거→2카드+규칙 도표 밴드), 배너·제목·푸터 보존, fix_clipping.
"""
import zipfile, os, re
import xml.etree.ElementTree as ET

SRC=r"d:\data_intelligence_2\report\데이터지도_분석보고서_V15.pptx"
DST=r"d:\data_intelligence_2\report\데이터지도_분석보고서_V16.pptx"
SLIDE="ppt/slides/slide15.xml"
def EMU(i): return int(round(i*914400))
NAVY="081F46"; GREEN="00965F"; DGRAY="505050"; BODYTX="282D37"; WHITE="FFFFFF"
LIGHT="EEF2F8"; LGREEN="E8F5F0"; LINE="D0D8E4"; LRED="FCE8E6"; RED="C0392B"; FONT="맑은 고딕"
HEADER_H=EMU(0.36)
_id=[5000]
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
def _run(t,sz,b_,c):
    return (f'<a:r><a:rPr sz="{sz}" b="{1 if b_ else 0}" i="0"><a:solidFill><a:srgbClr val="{c}"/></a:solidFill>'
            f'<a:latin typeface="{FONT}"/><a:ea typeface="{FONT}"/></a:rPr><a:t>{esc(t)}</a:t></a:r>')
def _para(runs,algn="l",lnspc=114000,spcaft=14):
    rs="".join(_run(t,sz,bb,c) for (t,sz,bb,c) in runs)
    return (f'<a:p><a:pPr algn="{algn}"><a:lnSpc><a:spcPct val="{lnspc}"/></a:lnSpc>'
            f'<a:spcBef><a:spcPts val="0"/></a:spcBef><a:spcAft><a:spcPts val="{spcaft}"/></a:spcAft><a:buNone/></a:pPr>{rs}</a:p>')
def tbox(x,y,w,h,paras,anchor="t",lins=12700,name="txt"):
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{nid()}" name="{name}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
            f'<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
            f'<p:txBody><a:bodyPr wrap="square" lIns="{lins}" tIns="6350" rIns="{lins}" bIns="6350" anchor="{anchor}"><a:normAutofit/></a:bodyPr>'
            f'<a:lstStyle/>{"".join(paras)}</p:txBody></p:sp>')
def frame(x,y,w,h,hdr,bg,line,title):
    return [roundrect(x,y,w,h,bg,line,name="r-bg"),
            roundrect(x,y,w,HEADER_H,hdr,hdr,name="r-hdr"),
            tbox(x,y,w,HEADER_H,[_para([(title,1080,True,WHITE)],algn="l",spcaft=0)],anchor="ctr",lins=EMU(0.13),name="r-title")]
def chip(x,y,w,h,text,fill,line,tcol,sz=920):
    return [roundrect(x,y,w,h,fill,line,name="r-chip"),
            tbox(x,y,w,h,[_para([(text,sz,True,tcol)],algn="ctr",spcaft=0)],anchor="ctr",lins=6350,name="r-chiptx")]

LX=EMU(0.62); W=EMU(12.09)
shapes=[]
# ① 패밀리 생성  ② PII 후보  (row1)
RY=EMU(2.0); RH=EMU(2.00); LCW=EMU(5.85); RCX=LX+LCW+EMU(0.25); RCW=EMU(5.99)
shapes+=frame(LX,RY,LCW,RH,NAVY,LIGHT,LINE,"① 의료기관 패밀리는 어떻게 만들어졌나")
shapes.append(tbox(LX+EMU(0.13),RY+HEADER_H+EMU(0.05),LCW-EMU(0.26),RH-HEADER_H-EMU(0.10),[
 _para([("장바구니: ",980,True,NAVY),("각 테이블 → 보유 컬럼개념 집합(정규화 col_nm)",980,False,BODYTX)],spcaft=7),
 _para([("FP-Growth: ",980,True,NAVY),("최대빈발 항목집합(min_support ≥ 8 테이블)",980,False,BODYTX)],spcaft=7),
 _para([("결과: ",980,True,GREEN),("{순번·의료기관명·의료기관전화번호·의료기관주소} = 11개 보건 테이블 공통 → ‘의료기관 현황’ 데이터셋 원형",980,False,BODYTX)],spcaft=0)],
 name="r-fam"))
shapes+=frame(RCX,RY,RCW,RH,NAVY,LIGHT,LINE,"② PII 후보 (개인식별정보 후보)")
shapes.append(tbox(RCX+EMU(0.13),RY+HEADER_H+EMU(0.05),RCW-EMU(0.26),RH-HEADER_H-EMU(0.10),[
 _para([("정의: ",980,True,NAVY),("값이 개인·기관을 특정할 수 있는 컬럼 유형(연락처·주소·명칭·좌표)",980,False,BODYTX)],spcaft=7),
 _para([("식별: ",980,True,NAVY),("col_nm이 식별자 유형 + 교차도메인 다출현 → PII 후보로 태깅",980,False,BODYTX)],spcaft=7),
 _para([("예시(등장 테이블): ",980,True,NAVY),("전화번호 140 · 주소 85 · 위경도 44 · 의료기관명 39 · 연락처 27",980,True,RED)],spcaft=0)],
 name="r-pii"))

# ③ 연관규칙 도표 밴드 (row2, full width)
BY=EMU(4.15); BH=EMU(2.75)
shapes+=frame(LX,BY,W,BH,GREEN,LGREEN,LINE,"③ 발견된 연관규칙 (장바구니 동시출현) — 재식별 위험의 근거")
cy=BY+HEADER_H+EMU(0.10)
# chips: [선행] ▶ [후행]
shapes.append(tbox(LX+EMU(0.18),cy,EMU(0.95),EMU(0.42),[_para([("선행",820,True,DGRAY)],algn="l",spcaft=0)],anchor="ctr",lins=6350,name="r-lbl1"))
shapes+=chip(LX+EMU(1.00),cy,EMU(2.55),EMU(0.42),"의료기관전화번호",LGREEN,GREEN,GREEN,sz=1000)
shapes.append(tbox(LX+EMU(3.58),cy,EMU(0.40),EMU(0.42),[_para([("▶",1500,True,DGRAY)],algn="ctr",spcaft=0)],anchor="ctr",lins=0,name="r-arrow"))
shapes.append(tbox(LX+EMU(4.02),cy,EMU(0.95),EMU(0.42),[_para([("후행",820,True,DGRAY)],algn="l",spcaft=0)],anchor="ctr",lins=6350,name="r-lbl2"))
shapes+=chip(LX+EMU(4.78),cy,EMU(4.20),EMU(0.42),"의료기관명 · 의료기관주소",LRED,RED,RED,sz=1000)
shapes.append(tbox(LX+EMU(9.15),cy,EMU(2.80),EMU(0.42),
    [_para([("support 15 · conf 0.88 · lift 46.4",900,True,NAVY)],algn="l",spcaft=0)],anchor="ctr",name="r-metric"))
# label lines
ly=cy+EMU(0.58)
shapes.append(tbox(LX+EMU(0.18),ly,W-EMU(0.40),BH-HEADER_H-EMU(0.78),[
 _para([("읽기: ",980,True,NAVY),("의료기관전화번호가 있으면 88%가 ‘명·주소’도 함께 있음 → 우연 대비 46배, 예외 거의 없음(conviction 8.3)",980,False,BODYTX)],spcaft=8),
 _para([("의미: ",980,True,RED),("명+주소+전화가 한 데이터셋에 묶여 공개 = 직접 재식별 가능 · 도메인 횡단 공통 식별자로 결합 위험 증폭 (연계 기회 ↔ 프라이버시 위험)",980,False,BODYTX)],spcaft=8),
 _para([("대책: ",980,True,GREEN),("PII 마스킹·집계화 · 공개 전 재식별 영향평가 · 식별자 컬럼 자동 태깅 · combined_pair_map.linking_ratio로 결합가능성 정량화",980,False,BODYTX)],spcaft=0)],
 name="r-band"))

def fix_clipping(xml):
    def repl(m):
        bp=m.group(1)
        bp2=re.sub(r'wrap="[^"]*"','wrap="square"',bp,count=1) if 'wrap="' in bp else bp.replace('<a:bodyPr','<a:bodyPr wrap="square"',1)
        if '<a:spAutoFit/>' in bp2: pass
        elif '<a:normAutofit' in bp2: bp2=re.sub(r'<a:normAutofit[^>]*/>','<a:normAutofit/>',bp2)
        elif '<a:noAutofit/>' in bp2: bp2=bp2.replace('<a:noAutofit/>','<a:normAutofit/>')
        elif bp2.endswith('/>'): bp2=bp2[:-2]+'><a:normAutofit/></a:bodyPr>'
        else: bp2=re.sub(r'(<a:bodyPr[^>]*>)',r'\1<a:normAutofit/>',bp2,count=1)
        return '<p:txBody>'+bp2
    return re.sub(r'<p:txBody>(<a:bodyPr[^>]*/>|<a:bodyPr[^>]*>.*?</a:bodyPr>)',repl,xml,flags=re.DOTALL)

def main():
    with zipfile.ZipFile(SRC) as z:
        order=z.infolist(); data={i.filename:z.read(i.filename) for i in order}
    xml=data[SLIDE].decode("utf-8")
    def rm(m):
        b=m.group(0); mo=re.search(r'<a:off x="-?\d+" y="(-?\d+)"/>',b)
        return '' if (mo and 1800000<=int(mo.group(1))<6000000) else b
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
    print("removed body sp:",removed,"| added:",len(shapes),"| ids OK")
    print("WROTE",DST)

if __name__=="__main__":
    main()
