# -*- coding: utf-8 -*-
"""V20 (Base=사용자 정리본 V19, 14장): '속성원형(보건+유일키+협소) 테이블이 합치기 좋은 구조인가' 검증 결과를
새 슬라이드로 추가(=새 14p, 연계 기회 13p 뒤). 인텔리전스/통합 주제.

검증(데이터 직접계산, analysis/_v20_merge.md):
 - 속성원형 105개 중 컬럼집합 완전일치 3그룹(11개) + 컬럼 Jaccard>=0.8 동형 쌍 41 → 다수가 동일 스키마
 - 순번(유일키) 보유 + 지자체·유형만 상이 = 수평 조각 → 표준 UNION 병합 가능
 실제 예시 2건:
  · 의료기기 판매업소: 부산 서구206·중구278·동구349·군산553 (동일4컬럼) → 전국 1,386행
  · 의료기관 현황: 익산368·부평682·아산한의원61·부산북구339 (동일5컬럼) → 전국 1,450행

방식(안전): slide13.xml(연계,인텔리전스+배너) 복사→새 slide15.xml(제목·배너·푸터·본문 교체). CT/rels/sldIdLst 배선.
footer 재번호(new=14·로드맵 slide14→15). ET 검증.
"""
import zipfile, os, re
import xml.etree.ElementTree as ET
SRC=r"d:\data_intelligence_2\report\데이터지도_분석보고서_V19.pptx"
DST=r"d:\data_intelligence_2\report\데이터지도_분석보고서_V20.pptx"
def EMU(i): return int(round(i*914400))
NAVY="081F46"; GREEN="00965F"; DGRAY="505050"; BODYTX="282D37"; WHITE="FFFFFF"
LIGHT="EEF2F8"; LGREEN="E8F5F0"; LINE="D0D8E4"; RED="C0392B"; FONT="맑은 고딕"
HEADER_H=EMU(0.36)
_id=[8000]
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
            f'<a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>'
            f'<a:ln w="9525"><a:solidFill><a:srgbClr val="{linecol}"/></a:solidFill></a:ln><a:effectLst/></p:spPr>{STYLE}{EMPTY_TX}</p:sp>')
def _run(t,sz,b_,c):
    return (f'<a:r><a:rPr sz="{sz}" b="{1 if b_ else 0}" i="0"><a:solidFill><a:srgbClr val="{c}"/></a:solidFill>'
            f'<a:latin typeface="{FONT}"/><a:ea typeface="{FONT}"/></a:rPr><a:t>{esc(t)}</a:t></a:r>')
def _para(runs,algn="l",lnspc=112000,spcaft=12):
    return (f'<a:p><a:pPr algn="{algn}"><a:lnSpc><a:spcPct val="{lnspc}"/></a:lnSpc>'
            f'<a:spcBef><a:spcPts val="0"/></a:spcBef><a:spcAft><a:spcPts val="{spcaft}"/></a:spcAft><a:buNone/></a:pPr>'
            f'{"".join(_run(t,sz,bb,c) for t,sz,bb,c in runs)}</a:p>')
def tbox(x,y,w,h,paras,anchor="t",lins=12700,name="txt"):
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{nid()}" name="{name}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
            f'<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
            f'<p:txBody><a:bodyPr wrap="square" lIns="{lins}" tIns="6350" rIns="{lins}" bIns="6350" anchor="{anchor}"><a:normAutofit/></a:bodyPr>'
            f'<a:lstStyle/>{"".join(paras)}</p:txBody></p:sp>')
def frame(x,y,w,h,hdr,bg,line,title):
    return [roundrect(x,y,w,h,bg,line,"f-bg"),roundrect(x,y,w,HEADER_H,hdr,hdr,"f-hdr"),
            tbox(x,y,w,HEADER_H,[_para([(title,1080,True,WHITE)],algn="l",spcaft=0)],anchor="ctr",lins=EMU(0.13),name="f-title")]
def _cell(text,sz,bold,color,fill,algn="l"):
    return (f'<a:tc><a:txBody><a:bodyPr wrap="square" lIns="46000" rIns="23000" tIns="8000" bIns="8000" anchor="ctr"/><a:lstStyle/>'
            f'<a:p><a:pPr algn="{algn}"/><a:r><a:rPr lang="ko-KR" sz="{sz}" b="{1 if bold else 0}">'
            f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill><a:latin typeface="{FONT}"/><a:ea typeface="{FONT}"/></a:rPr>'
            f'<a:t>{esc(text)}</a:t></a:r></a:p></a:txBody><a:tcPr><a:solidFill><a:srgbClr val="{fill}"/></a:solidFill></a:tcPr></a:tc>')
def make_table(x,y,w,colws,header,rows,ha,rowh=EMU(0.70),hh=EMU(0.40),fs=950,hfs=980):
    grid="".join(f'<a:gridCol w="{c}"/>' for c in colws)
    trs=[f'<a:tr h="{hh}">'+"".join(_cell(h,hfs,True,WHITE,NAVY,ha[i]) for i,h in enumerate(header))+'</a:tr>']
    for ri,row in enumerate(rows):
        fill=WHITE if ri%2 else LIGHT; cells=""
        for ci,val in enumerate(row):
            color=NAVY if ci==0 else (GREEN if ci==len(row)-1 else BODYTX); bold=(ci==0 or ci==len(row)-1)
            cells+=_cell(val,fs,bold,color,fill,ha[ci])
        trs.append(f'<a:tr h="{rowh}">{cells}</a:tr>')
    cy=hh+rowh*len(rows)
    return (f'<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="{nid()}" name="ex-tbl"/>'
            f'<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>'
            f'<p:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{cy}"/></p:xfrm>'
            f'<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">'
            f'<a:tbl><a:tblPr firstRow="1" bandRow="1"/><a:tblGrid>{grid}</a:tblGrid>{"".join(trs)}</a:tbl>'
            f'</a:graphicData></a:graphic></p:graphicFrame>')
def merge_runs(sp,txt):
    m=re.search(r'<a:rPr.*?</a:rPr>|<a:rPr[^>]*/>',sp,re.S); rpr=m.group(0) if m else '<a:rPr/>'
    fr=sp.index('<a:r>'); lr=sp.rindex('</a:r>')+len('</a:r>')
    return sp[:fr]+f'<a:r>{rpr}<a:t>{esc(txt)}</a:t></a:r>'+sp[lr:]
def edit_sp_by_y(xml,lo,hi,txt):
    for m in re.finditer(r'<p:sp>.*?</p:sp>',xml,re.S):
        sp=m.group(0); mo=re.search(r'<a:off x="-?\d+" y="(-?\d+)"/>',sp)
        if mo and lo<=int(mo.group(1))<hi and '<a:r>' in sp: return xml[:m.start()]+merge_runs(sp,txt)+xml[m.end():]
    raise RuntimeError(f"no sp y[{lo},{hi})")
def edit_sp_by_name(xml,name,txt):
    m=re.search(r'<p:sp>(?:(?!</p:sp>).)*name="'+re.escape(name)+r'".*?</p:sp>',xml,re.S)
    if not m: raise RuntimeError("no "+name)
    return xml[:m.start()]+merge_runs(m.group(0),txt)+xml[m.end():]
def set_footer(xml,num):
    try: return edit_sp_by_y(xml,6200000,9999999,num)
    except RuntimeError: return xml
def rm_body(xml,lo=1800000,hi=6300000):
    return re.sub(r"<p:sp>.*?</p:sp>",lambda m:('' if (lambda mo:mo and lo<=int(mo.group(1))<hi)(re.search(r'<a:off x="-?\d+" y="(-?\d+)"/>',m.group(0))) else m.group(0)),xml,flags=re.S)
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
def add(xml,shapes): return xml.replace("</p:spTree>","".join(shapes)+"</p:spTree>")

def build_new(s13):
    x=s13
    x=edit_sp_by_y(x,300000,700000,"속성원형 → 데이터 통합(병합) 기회")
    x=edit_sp_by_name(x,"metric-number","41쌍")
    x=edit_sp_by_name(x,"metric-pattern","발견 패턴 ▸ 동일 스키마 지자체별 조각 = 표준 UNION 병합")
    x=edit_sp_by_name(x,"metric-desc","보건 ‘작고 키 있는’ 데이터셋이 동일 스키마로 산재 → 지자체·유형 조각을 전국 마스터로 통합 가능(동형 쌍 41)")
    x=set_footer(x,"14")
    x=rm_body(x)
    x=re.sub(r'<p:graphicFrame>.*?</p:graphicFrame>','',x,flags=re.S)  # 연계 슬라이드의 11개 테이블(graphicFrame) 잔상 제거
    LX=EMU(0.62); shapes=[]
    LCW=EMU(5.85); RCX=LX+LCW+EMU(0.25); RCW=EMU(5.99); Y=EMU(2.0); CH=EMU(2.00)
    # ① 병합 가능성(검증)
    shapes+=frame(LX,Y,LCW,CH,NAVY,LIGHT,LINE,"① 합치기 좋은가? — 검증 결과")
    shapes.append(tbox(LX+EMU(0.14),Y+HEADER_H+EMU(0.06),LCW-EMU(0.28),CH-HEADER_H-EMU(0.12),[
     _para([("속성원형 105개 중 ",980,False,BODYTX),("컬럼집합 완전일치 3그룹(11개) + 동형(Jaccard≥0.8) 41쌍",980,True,RED)],spcaft=8),
     _para([("순번(유일키) 보유 + 지자체·유형만 상이 = 수평 조각(partition)",980,False,BODYTX)],spcaft=8),
     _para([("→ ",980,True,GREEN),("표준 스키마로 UNION 병합 가능 = ‘합치기 좋은 구조’ 확인",980,True,GREEN)],spcaft=0)],name="m-ver"))
    # ② 병합 방법·효과
    shapes+=frame(RCX,Y,RCW,CH,NAVY,LIGHT,LINE,"② 병합 방법 · 효과")
    shapes.append(tbox(RCX+EMU(0.14),Y+HEADER_H+EMU(0.06),RCW-EMU(0.28),CH-HEADER_H-EMU(0.12),[
     _para([("방법: ",980,True,NAVY),("동일 컬럼 확인 → 출처(지자체)·유형 컬럼 추가 → UNION ALL → 전국 마스터",980,False,BODYTX)],spcaft=8),
     _para([("키: ",980,True,NAVY),("순번은 테이블 내 유일 → 전국 통합 시 (지자체+순번) 복합키로 재부여",980,False,BODYTX)],spcaft=8),
     _para([("효과: ",980,True,GREEN),("산재 데이터 전국 단위 분석·중복 제거·갱신 일원화",980,False,BODYTX)],spcaft=0)],name="m-how"))
    # ③ 실제 예시 표
    by=EMU(4.15); shapes+=frame(LX,by,EMU(12.09),EMU(2.75),GREEN,LGREEN,LINE,"③ 실제 병합 예시 (동일 스키마 · 지자체만 상이)")
    rows=[["의료기기 판매업소","순번·영업소명·소재지·전화 (4)","부산 서구206·중구278·동구349·군산553","4개 → 전국 1,386행"],
          ["의료기관 현황","순번·의료기관명·전화·종별·주소 (5)","익산368·부평682·아산한의원61·부산북구339","4개 → 전국 1,450행"]]
    tbl=make_table(LX+EMU(0.14),by+HEADER_H+EMU(0.10),EMU(12.09)-EMU(0.28),
                   [EMU(1.95),EMU(3.25),EMU(4.65),EMU(1.96)],
                   ["예시(주제)","동일 컬럼(스키마)","조각 테이블 (지자체·레코드)","통합 결과"],rows,
                   ha=["l","l","l","ctr"],rowh=EMU(0.62),hh=EMU(0.42))
    shapes.append(tbl)
    shapes.append(tbox(LX+EMU(0.14),by+EMU(2.18),EMU(12.09)-EMU(0.28),EMU(0.34),
     [_para([("두 예시 모두 컬럼·키 구조 동일 → ‘출처’ 컬럼만 더해 UNION하면 전국 표 완성 (속성원형이 병합 친화적 구조임을 실증)",900,False,DGRAY)],spcaft=0)],name="m-cap"))
    return fix_clipping(add(x,shapes))

def main():
    with zipfile.ZipFile(SRC) as z:
        names=z.namelist(); data={n:z.read(n) for n in names}
    new=build_new(data["ppt/slides/slide13.xml"].decode("utf-8")); ET.fromstring(new)
    data["ppt/slides/slide15.xml"]=new.encode("utf-8")
    data["ppt/slides/_rels/slide15.xml.rels"]=data["ppt/slides/_rels/slide13.xml.rels"]
    ct=data["[Content_Types].xml"].decode("utf-8").replace("</Types>",'<Override PartName="/ppt/slides/slide15.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>')
    data["[Content_Types].xml"]=ct.encode("utf-8")
    pr=data["ppt/_rels/presentation.xml.rels"].decode("utf-8")
    NEWRID="rId%d"%(max(int(n) for n in re.findall(r'Id="rId(\d+)"',pr))+1)
    pr=pr.replace("</Relationships>",f'<Relationship Id="{NEWRID}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide15.xml"/></Relationships>')
    data["ppt/_rels/presentation.xml.rels"]=pr.encode("utf-8")
    rid_s13=[i for i,t in re.findall(r'Id="(rId\d+)"[^>]*Target="(slides/slide\d+\.xml)"',pr) if t=="slides/slide13.xml"][0]
    pres=data["ppt/presentation.xml"].decode("utf-8")
    anchor=re.search(r'<p:sldId id="\d+" r:id="'+rid_s13+r'"/>',pres).group(0)
    pres=pres.replace(anchor,anchor+f'<p:sldId id="291" r:id="{NEWRID}"/>')
    data["ppt/presentation.xml"]=pres.encode("utf-8")
    # footer: new slide15->14, slide14(로드맵)->15 ; others 1-13 unchanged
    data["ppt/slides/slide14.xml"]=set_footer(data["ppt/slides/slide14.xml"].decode("utf-8"),"15").encode("utf-8")
    for k in ["ppt/slides/slide15.xml","ppt/presentation.xml","ppt/slides/slide14.xml"]: ET.fromstring(data[k])
    if os.path.exists(DST): os.remove(DST)
    order=names+["ppt/slides/slide15.xml","ppt/slides/_rels/slide15.xml.rels"]
    with zipfile.ZipFile(DST,"w",zipfile.ZIP_DEFLATED) as zo:
        for n in order: zo.writestr(n,data[n])
    print("V20: 속성원형 병합 기회 슬라이드 삽입(연계기회 뒤) 완료")
    print("WROTE",DST)

if __name__=="__main__":
    main()
