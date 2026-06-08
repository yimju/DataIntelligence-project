# -*- coding: utf-8 -*-
"""V17:
 1) 9페이지(발견 리스트) 표를 '이후 상세 페이지 순서(p.10→16)'대로 재배열 — 페이지당 1행.
 2) 11페이지(slide15)를 '교차도메인 재식별 위험' → '교차도메인 데이터 연계 기회'로 리프레임.
    근거: 의료기관·장소 정보는 개인정보(PII)가 아닌 공개 기관 정보 → 재식별이 아니라 '연계 자산'.
    PII 후보 → '연계 키 후보(공통 조인 키)'로 재해석, 의미·대책(활용)을 연계 관점으로 재작성.

방식(안전): Base=V16. slide9.xml(표·캡션 교체)·slide15.xml(제목·섹션·배너설명·본문)만 편집.
배너 도형/제목영역 좌표 보존, fix_clipping, ET 검증. 그 외 byte-동일.
"""
import zipfile, os, re
import xml.etree.ElementTree as ET

SRC=r"d:\data_intelligence_2\report\데이터지도_분석보고서_V16.pptx"
DST=r"d:\data_intelligence_2\report\데이터지도_분석보고서_V17.pptx"
def EMU(i): return int(round(i*914400))
NAVY="081F46"; GREEN="00965F"; DGRAY="505050"; BODYTX="282D37"; WHITE="FFFFFF"
LIGHT="EEF2F8"; LGREEN="E8F5F0"; LINE="D0D8E4"; LRED="FCE8E6"; RED="C0392B"; FONT="맑은 고딕"
HEADER_H=EMU(0.36)
_id=[6000]
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
    return [roundrect(x,y,w,h,bg,line,name="r-bg"),roundrect(x,y,w,HEADER_H,hdr,hdr,name="r-hdr"),
            tbox(x,y,w,HEADER_H,[_para([(title,1080,True,WHITE)],algn="l",spcaft=0)],anchor="ctr",lins=EMU(0.13),name="r-title")]
def chip(x,y,w,h,text,fill,line,tcol,sz=1000):
    return [roundrect(x,y,w,h,fill,line,name="r-chip"),
            tbox(x,y,w,h,[_para([(text,sz,True,tcol)],algn="ctr",spcaft=0)],anchor="ctr",lins=6350,name="r-chiptx")]
def _cell(text,sz,bold,color,fill,algn="l"):
    return (f'<a:tc><a:txBody><a:bodyPr wrap="square" lIns="50000" rIns="25000" tIns="10000" bIns="10000" anchor="ctr"/><a:lstStyle/>'
            f'<a:p><a:pPr algn="{algn}"/><a:r><a:rPr lang="ko-KR" sz="{sz}" b="{1 if bold else 0}">'
            f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill><a:latin typeface="{FONT}"/><a:ea typeface="{FONT}"/></a:rPr>'
            f'<a:t>{esc(text)}</a:t></a:r></a:p></a:txBody><a:tcPr><a:solidFill><a:srgbClr val="{fill}"/></a:solidFill></a:tcPr></a:tc>')
def make_table(x,y,w,colws,header,rows,rowh=EMU(0.55),hh=EMU(0.42),fs=1000,hfs=1050,headalgn=None):
    grid="".join(f'<a:gridCol w="{c}"/>' for c in colws)
    ha=headalgn or ["l","l","ctr","ctr"]
    trs=[f'<a:tr h="{hh}">'+"".join(_cell(h,hfs,True,WHITE,NAVY,ha[i] if i<len(ha) else "ctr") for i,h in enumerate(header))+'</a:tr>']
    for ri,row in enumerate(rows):
        fill=WHITE if ri%2 else LIGHT; cells=""
        for ci,val in enumerate(row):
            color=NAVY if ci==0 else (GREEN if ci==3 else BODYTX); bold=(ci==0 or ci==3); algn=ha[ci] if ci<len(ha) else "ctr"
            cells+=_cell(val,fs,bold,color,fill,algn)
        trs.append(f'<a:tr h="{rowh}">{cells}</a:tr>')
    cy=hh+rowh*len(rows)
    return (f'<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="{nid()}" name="list-tbl"/>'
            f'<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>'
            f'<p:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{cy}"/></p:xfrm>'
            f'<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">'
            f'<a:tbl><a:tblPr firstRow="1" bandRow="1"/><a:tblGrid>{grid}</a:tblGrid>{"".join(trs)}</a:tbl>'
            f'</a:graphicData></a:graphic></p:graphicFrame>')

def merge_runs(sp,newtext):
    rpr_m=re.search(r'<a:rPr.*?</a:rPr>|<a:rPr[^>]*/>',sp,re.S); rpr=rpr_m.group(0) if rpr_m else '<a:rPr/>'
    newrun=f'<a:r>{rpr}<a:t>{esc(newtext)}</a:t></a:r>'
    fr=sp.index('<a:r>'); lr=sp.rindex('</a:r>')+len('</a:r>')
    return sp[:fr]+newrun+sp[lr:]
def edit_sp_by_y(xml,ylo,yhi,newtext):
    for m in re.finditer(r'<p:sp>.*?</p:sp>',xml,re.S):
        sp=m.group(0); mo=re.search(r'<a:off x="-?\d+" y="(-?\d+)"/>',sp)
        if mo and ylo<=int(mo.group(1))<yhi and '<a:r>' in sp:
            return xml[:m.start()]+merge_runs(sp,newtext)+xml[m.end():]
    raise RuntimeError(f"no sp in y[{ylo},{yhi})")
def edit_sp_by_name(xml,name,newtext):
    m=re.search(r'<p:sp>(?:(?!</p:sp>).)*name="'+re.escape(name)+r'".*?</p:sp>',xml,re.S)
    if not m: raise RuntimeError("no shape "+name)
    return xml[:m.start()]+merge_runs(m.group(0),newtext)+xml[m.end():]
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

# =============== slide9: reorder list table (by detail page) ===============
def build_slide9(xml):
    # remove old table + caption
    xml=re.sub(r'<p:graphicFrame>.*?</p:graphicFrame>','',xml,flags=re.S)
    xml=re.sub(r'<p:sp>(?:(?!</p:sp>).)*name="list-cap".*?</p:sp>','',xml,flags=re.S)
    header=["구분","발견 내용","핵심 지표","상세"]
    rows=[  # ordered by following detail page p.10 → p.16 (페이지당 1행)
     ["거버넌스","출처 결손 — ‘유아및초·중등교육’ 100% placeholder","A1 conf 1.00 · lift 6.89","p.10"],
     ["연계","교차도메인 연계 — 의료기관전화 → {명·주소}","conf 0.88 · lift 46.4","p.11"],
     ["패턴","스키마 패밀리(의료기관·POI·사학재정·국가시험군)","FP-Growth 최대빈발 ≥8테이블","p.12"],
     ["규칙","연관규칙 — 경도→위도(22.7)·교과목코드↔분반(124.9)·시도→시군구(39)","lift 2 ~ 125","p.13"],
     ["시각화","연관규칙 lift·support 분포","이슈 규칙 10건","p.14"],
     ["검증","규칙#4 국소→전체 (도로명 → 지번)","conf 1.00 → 0.44","p.15"],
     ["기타","컬럼 표준·값지문 / 상수·스키마불일치 / 교차도메인값","1D 집계·샘플","p.16"],
    ]
    W=EMU(12.09); colws=[EMU(1.30),EMU(6.90),EMU(2.70),EMU(1.19)]
    tbl=make_table(EMU(0.62),EMU(2.05),W,colws,header,rows)
    cap=tbox(EMU(0.62),EMU(6.45),W,EMU(0.3),
             [_para([("표는 이후 상세 분석 페이지(p.10→16) 순서로 정렬 · 1D 집계·샘플 결과는 ‘기타(p.16)’",950,False,DGRAY)],spcaft=0)],name="list-cap")
    return fix_clipping(xml.replace("</p:spTree>",tbl+cap+"</p:spTree>"))

# =============== slide15: reframe 재식별→연계 + 11개 구성 테이블 표 ===============
FAM_TABLES=[  # 데이터에서 계산(scripts/_g3 류): 패밀리 4컬럼 모두 보유한 11개 보건 테이블
 ["1","전북 익산시","의료기관","368"],["2","인천 부평구","의료기관 현황","682"],
 ["3","부산 강서구","병의원 현황","118"],["4","전남 목포시","병원현황","42"],
 ["5","대전 동구","특수의료장비 공동활용병상 의료기관","26"],["6","부산 기장군","한의원 현황","39"],
 ["7","충남 아산시","한의원현황","61"],["8","충남 아산시","치과의원현황","93"],
 ["9","전북 김제시","병원현황","65"],["10","부산 북구","의원급의료기관 현황","339"],
 ["11","부산 동래구","피부과의원","11"],
]
def build_slide15(xml):
    xml=edit_sp_by_y(xml,0,200000,"인텔리전스")
    xml=edit_sp_by_y(xml,300000,700000,"인텔리전스 교차도메인 데이터 연계 기회")
    xml=edit_sp_by_name(xml,"metric-desc","명·주소·전화가 항상 묶여 있어 같은 기관을 여러 데이터셋에서 연결할 수 있는 ‘연계 기회’를 형성함")
    def rm(m):
        b=m.group(0); mo=re.search(r'<a:off x="-?\d+" y="(-?\d+)"/>',b)
        return '' if (mo and 1800000<=int(mo.group(1))<6000000) else b
    xml=re.sub(r"<p:sp>.*?</p:sp>",rm,xml,flags=re.DOTALL)
    LX=EMU(0.62); shapes=[]
    LCW=EMU(5.00); RCX=LX+LCW+EMU(0.20); RCW=EMU(6.89)
    # ① 패밀리 생성 (좌상)
    AY=EMU(2.0); AH=EMU(2.30)
    shapes+=frame(LX,AY,LCW,AH,NAVY,LIGHT,LINE,"① 의료기관 패밀리 생성")
    shapes.append(tbox(LX+EMU(0.13),AY+HEADER_H+EMU(0.05),LCW-EMU(0.26),AH-HEADER_H-EMU(0.10),[
     _para([("장바구니: ",950,True,NAVY),("각 테이블 → 보유 컬럼개념 집합",950,False,BODYTX)],spcaft=6),
     _para([("FP-Growth: ",950,True,NAVY),("최대빈발 항목집합(≥8테이블)",950,False,BODYTX)],spcaft=6),
     _para([("결과: ",950,True,GREEN),("{순번·의료기관명·의료기관전화번호·의료기관주소} 4컬럼 공통",950,False,BODYTX)],spcaft=6),
     _para([("연계 키 후보: ",950,True,NAVY),("전화·주소·명·위경도 (공통 조인 키, PII 아님)",950,False,BODYTX)],spcaft=0)],name="r-fam"))
    # ③ 연관규칙 도표 (좌하, 세로 chip)
    BY=EMU(4.45); BH=EMU(2.45)
    shapes+=frame(LX,BY,LCW,BH,GREEN,LGREEN,LINE,"③ 연관규칙 — 연계 근거")
    cx=LX+EMU(0.55); cw=EMU(3.90); cyy=BY+HEADER_H+EMU(0.10)
    shapes+=chip(cx,cyy,cw,EMU(0.40),"의료기관전화번호",LGREEN,GREEN,GREEN,sz=1000)
    shapes.append(tbox(cx,cyy+EMU(0.40),cw,EMU(0.30),[_para([("▼",1200,True,DGRAY)],algn="ctr",spcaft=0)],anchor="ctr",lins=0,name="r-arrow"))
    shapes+=chip(cx,cyy+EMU(0.70),cw,EMU(0.40),"의료기관명 · 의료기관주소",LIGHT,NAVY,NAVY,sz=1000)
    shapes.append(tbox(LX+EMU(0.16),cyy+EMU(1.22),LCW-EMU(0.30),EMU(0.80),[
     _para([("support 15 · conf 0.88 · lift 46.4",950,True,NAVY)],algn="ctr",spcaft=6),
     _para([("→ 같은 기관을 여러 데이터셋에서 연결 = 연계 기회",950,False,BODYTX)],algn="ctr",spcaft=0)],anchor="t",name="r-rule"))
    # ② 11개 구성 테이블 표 (우측)
    TY=EMU(2.0); TH=EMU(4.90)
    shapes+=frame(RCX,TY,RCW,TH,NAVY,LIGHT,LINE,"② 패밀리 구성 11개 테이블 (보건 · 4컬럼 공통 보유)")
    colws=[EMU(0.45),EMU(1.62),EMU(3.46),EMU(1.06)]
    tbl=make_table(RCX+EMU(0.13),TY+HEADER_H+EMU(0.06),RCW-EMU(0.26),colws,
                   ["#","지자체","데이터셋(의료기관 유형)","레코드"],FAM_TABLES,
                   rowh=EMU(0.32),hh=EMU(0.34),fs=900,hfs=950,headalgn=["ctr","l","l","ctr"])
    shapes.append(tbl)
    shapes.append(tbox(RCX+EMU(0.13),TY+TH-EMU(0.40),RCW-EMU(0.26),EMU(0.34),
     [_para([("4컬럼(순번·명·전화·주소)을 공통 보유 = 통합 가능 단위 · 합계 1,844행",880,False,DGRAY)],spcaft=0)],name="r-tcap"))
    return fix_clipping(xml.replace("</p:spTree>","".join(shapes)+"</p:spTree>"))

def main():
    with zipfile.ZipFile(SRC) as z:
        order=z.infolist(); data={i.filename:z.read(i.filename) for i in order}
    data["ppt/slides/slide9.xml"]=build_slide9(data["ppt/slides/slide9.xml"].decode("utf-8")).encode("utf-8")
    data["ppt/slides/slide15.xml"]=build_slide15(data["ppt/slides/slide15.xml"].decode("utf-8")).encode("utf-8")
    for fn in ["ppt/slides/slide9.xml","ppt/slides/slide15.xml"]:
        ET.fromstring(data[fn]); ids=re.findall(r'<p:cNvPr id="(\d+)"',data[fn].decode()); assert len(ids)==len(set(ids)),"dup "+fn
    if os.path.exists(DST): os.remove(DST)
    with zipfile.ZipFile(DST,"w",zipfile.ZIP_DEFLATED) as zout:
        for i in order: zout.writestr(i,data[i.filename])
    print("V17: slide9 표 재배열 + slide15 연계기회 리프레임 완료")
    print("WROTE",DST)

if __name__=="__main__":
    main()
