# -*- coding: utf-8 -*-
"""V19 (Base=사용자 간소화본 V18, 14장):
 1+2) 11페이지(거버넌스 출처 누락) 심화 분석 새 슬라이드 1장 삽입(=새 12p):
      ① regi_date·url 적재 배치 추적: 999건 모두 2026-03-30 일괄 / 적재순서 499~598=연속 100건 전부
         유아및초중등(타분류 0건 혼입)=단일 배치 출처 일괄 소실 / url도 999건 전부 'url' placeholder
      ② 전체 도메인 비교: cate2별 placeholder율 (유아및초중등 100% ≫ 교육일반15·고등14·보건3·평생2·행정1)
         cate1 교육33%≫보건3%≫공공행정1% → 교육(특히 유아및초중등)에 계통 집중(전체평균14.5% 대비 극단 outlier)
 3) 결론(13·14p) 현재 슬라이드 기준 현행화(제거된 G2/PII/서브스페이스/near-duplicate 참조 제거,
    탐사 대상 규칙·출처 누락·교차도메인 연계 중심으로 재작성)

방식(안전): slide11.xml 복사→새 slide15.xml(제목·배너·푸터·본문 교체). 파트/CT/rels/sldIdLst(slide11 뒤) 배선.
slide13·14 본문 교체. 전 슬라이드 푸터 1~15 재번호. ET 검증.
"""
import zipfile, os, re
import xml.etree.ElementTree as ET
SRC=r"d:\data_intelligence_2\report\데이터지도_분석보고서_V18.pptx"
DST=r"d:\data_intelligence_2\report\데이터지도_분석보고서_V19.pptx"
def EMU(i): return int(round(i*914400))
NAVY="081F46"; GREEN="00965F"; DGRAY="505050"; BODYTX="282D37"; WHITE="FFFFFF"
LIGHT="EEF2F8"; LGREEN="E8F5F0"; LINE="D0D8E4"; RED="C0392B"; FONT="맑은 고딕"
HEADER_H=EMU(0.36)
_id=[7000]
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
def make_table(x,y,w,colws,header,rows,rowh=EMU(0.40),hh=EMU(0.38),fs=950,hfs=980,hi=None):
    grid="".join(f'<a:gridCol w="{c}"/>' for c in colws)
    ha=["l","ctr","ctr"]
    trs=[f'<a:tr h="{hh}">'+"".join(_cell(h,hfs,True,WHITE,NAVY,ha[i] if i<len(ha) else "ctr") for i,h in enumerate(header))+'</a:tr>']
    for ri,row in enumerate(rows):
        em=(hi is not None and ri==hi)
        fill=LGREEN if em else (WHITE if ri%2 else LIGHT)
        cells=""
        for ci,val in enumerate(row):
            color=RED if (em and ci>=1) else (NAVY if ci==0 else BODYTX); bold=(ci==0 or em)
            cells+=_cell(val,fs,bold,color,fill,ha[ci] if ci<len(ha) else "ctr")
        trs.append(f'<a:tr h="{rowh}">{cells}</a:tr>')
    cy=hh+rowh*len(rows)
    return (f'<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="{nid()}" name="cmp-tbl"/>'
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
    def rm(m):
        b=m.group(0); mo=re.search(r'<a:off x="-?\d+" y="(-?\d+)"/>',b)
        return '' if (mo and lo<=int(mo.group(1))<hi) else b
    return re.sub(r"<p:sp>.*?</p:sp>",rm,xml,flags=re.S)
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

# ---------- 새 슬라이드(출처누락 심화) : slide11 복사 후 개조 ----------
def build_deepdive(s11):
    x=s11
    x=edit_sp_by_y(x,300000,750000,"거버넌스 출처 누락 — 심화 (적재 배치 · 도메인 비교)")
    x=edit_sp_by_name(x,"metric-number","100건")
    x=edit_sp_by_name(x,"metric-pattern","발견 패턴 ▸ 적재 순서 연속 100건 = 단일 배치")
    x=edit_sp_by_name(x,"metric-desc","적재 순서 499~598(연속 100건)이 모두 ‘유아및초·중등교육’ — 단일 수집 배치에서 출처 일괄 소실")
    x=set_footer(x,"12")
    x=rm_body(x)
    LX=EMU(0.62); shapes=[]
    LCW=EMU(5.85); RCX=LX+LCW+EMU(0.25); RCW=EMU(5.99); Y=EMU(2.0); H=EMU(4.90)
    # ① regi_date·url
    shapes+=frame(LX,Y,LCW,H,NAVY,LIGHT,LINE,"① 적재 배치 추적 (regi_date · url)")
    shapes.append(tbox(LX+EMU(0.14),Y+HEADER_H+EMU(0.06),LCW-EMU(0.28),H-HEADER_H-EMU(0.12),[
     _para([("전체 999건 regi_date = 2026-03-30 ",980,True,NAVY),("(카탈로그 일괄 적재)",980,False,DGRAY)],spcaft=9),
     _para([("적재 순서 499~598번 = 연속 100건이 ",980,False,BODYTX),("모두 ‘유아및초·중등교육’",980,True,RED),(" (타 분류 0건 혼입)",980,False,DGRAY)],spcaft=9),
     _para([("11:28~11:37 (9분) 연속 → ",980,False,BODYTX),("단일 수집 배치에서 table_source 일괄 소실 확증",980,True,NAVY)],spcaft=9),
     _para([("url 필드도 999건 전부 ‘url’ placeholder ",980,True,NAVY),("= 카탈로그 전반 별도 품질 이슈(출처누락과 독립)",980,False,DGRAY)],spcaft=9),
     _para([("함의: ",980,True,GREEN),("공급원·수집경로 단위로 출처가 통째 누락 → 배치 단위 추적·복구 가능",980,False,BODYTX)],spcaft=0)],name="dd-regi"))
    # ② 도메인 비교
    shapes+=frame(RCX,Y,RCW,H,NAVY,LIGHT,LINE,"② 전체 도메인 비교 (출처 누락율)")
    rows=[["유아및초·중등교육","100/100","100%"],["교육일반","15/100","15%"],["고등교육","14/100","14%"],
          ["보건의료","10/300","3%"],["평생·직업교육","2/100","2%"],["일반행정","4/299","1%"]]
    tbl=make_table(RCX+EMU(0.14),Y+HEADER_H+EMU(0.08),RCW-EMU(0.28),[EMU(2.9),EMU(1.5),EMU(1.0)],
                   ["cate2 (중분류)","placeholder","율"],rows,rowh=EMU(0.40),hh=EMU(0.40),hi=0)
    shapes.append(tbl)
    shapes.append(tbox(RCX+EMU(0.14),Y+EMU(3.45),RCW-EMU(0.28),EMU(1.30),[
     _para([("cate1: ",980,True,NAVY),("교육 33%(131) ≫ 보건 3% ≫ 공공행정 1%",980,False,BODYTX)],spcaft=8),
     _para([("이슈: ",980,True,RED),("출처 누락은 교육(특히 유아및초중등)에 계통 집중 — 전체 평균 14.5% 대비 극단 outlier",980,False,BODYTX)],spcaft=8),
     _para([("대책: ",980,True,GREEN),("특정 공급기관 메타표준 미준수 → 공급원별 적재 검증·출처 필수화",980,False,BODYTX)],spcaft=0)],name="dd-cmp"))
    return fix_clipping(add(x,shapes))

# ---------- 결론 현행화 ----------
def update_slide13(x):
    x=rm_body(x)
    LX=EMU(0.62); W=EMU(12.09); gap=EMU(0.30); cw=(W-2*gap)//3; Y=EMU(2.05); H=EMU(4.4)
    xs=[LX,LX+cw+gap,LX+2*(cw+gap)]
    c1=frame(xs[0],Y,cw,H,NAVY,LIGHT,LINE,"즉시 (거버넌스)")
    c1.append(tbox(xs[0]+EMU(0.13),Y+HEADER_H+EMU(0.06),cw-EMU(0.26),H-HEADER_H-EMU(0.12),[
     _para([("· 출처 누락 단일 배치 provenance 복구",1000,False,BODYTX)],spcaft=9),
     _para([("· 적재 시 출처 필수검증(누락 차단)",1000,False,BODYTX)],spcaft=9),
     _para([("· url placeholder(999건) 정정",1000,False,BODYTX)],spcaft=0)],name="rm-c1"))
    c2=frame(xs[1],Y,cw,H,NAVY,LIGHT,LINE,"단기 (검증 · 표준)")
    c2.append(tbox(xs[1]+EMU(0.13),Y+HEADER_H+EMU(0.06),cw-EMU(0.26),H-HEADER_H-EMU(0.12),[
     _para([("· 탐사 대상 규칙(적절 lift) 도메인 검증",1000,False,BODYTX)],spcaft=9),
     _para([("· 규칙 기반 스키마 자동 태깅·무결성 체크",1000,False,BODYTX)],spcaft=9),
     _para([("· 좌표·연락처 등 표준 데이터모델 정의",1000,False,BODYTX)],spcaft=0)],name="rm-c2"))
    c3=frame(xs[2],Y,cw,H,GREEN,LGREEN,LINE,"중기 (연계 · 인텔리전스)")
    c3.append(tbox(xs[2]+EMU(0.13),Y+HEADER_H+EMU(0.06),cw-EMU(0.26),H-HEADER_H-EMU(0.12),[
     _para([("· 교차도메인 연계 — combined_pair_map",1000,False,BODYTX)],spcaft=4),
     _para([("  .linking_ratio로 조인키 정량 검증",1000,False,BODYTX)],spcaft=9),
     _para([("· 통합 기관 마스터(주소·연락처) 구축",1000,False,BODYTX)],spcaft=0)],name="rm-c3"))
    return fix_clipping(add(x,c1+c2+c3))
def update_slide14(x):
    x=rm_body(x)
    LX=EMU(0.62); W=EMU(12.09); gap=EMU(0.30); cw=(W-gap)//2; Y=EMU(2.05); H=EMU(4.5)
    c1=frame(LX,Y,cw,H,NAVY,LIGHT,LINE,"다음 단계 분석")
    c1.append(tbox(LX+EMU(0.14),Y+HEADER_H+EMU(0.06),cw-EMU(0.28),H-HEADER_H-EMU(0.12),[
     _para([("· 연계: ",1000,True,NAVY),("linking_ratio로 조인키 정량 검증 → 통합 마스터",1000,False,BODYTX)],spcaft=9),
     _para([("· 규칙: ",1000,True,NAVY),("탐사 대상 규칙 전체검증·운영 룰(필수 동반 컬럼) 태깅",1000,False,BODYTX)],spcaft=9),
     _para([("· 거버넌스: ",1000,True,NAVY),("출처·url 메타 표준화 · 공급원별 적재 검증",1000,False,BODYTX)],spcaft=0)],name="nx-c1"))
    c2=frame(LX+cw+gap,Y,cw,H,GREEN,LGREEN,LINE,"한계")
    c2.append(tbox(LX+cw+gap+EMU(0.14),Y+HEADER_H+EMU(0.06),cw-EMU(0.28),H-HEADER_H-EMU(0.12),[
     _para([("· 연관규칙 support 낮음(희소) → 운영 전 도메인 확인 필요",1000,False,BODYTX)],spcaft=9),
     _para([("· 활용/품질 지표 미충전(분석 축 제한)",1000,False,BODYTX)],spcaft=9),
     _para([("· regi_date 전건 2026-03-30 동일 → 시계열·이력 분석 불가",1000,False,BODYTX)],spcaft=9),
     _para([("· 값 표본 정렬 편향 → 국소 규칙 전체검증 필수",1000,False,BODYTX)],spcaft=0)],name="nx-c2"))
    return fix_clipping(add(x,c1+c2))

def main():
    with zipfile.ZipFile(SRC) as z:
        names=z.namelist(); data={n:z.read(n) for n in names}
    # new deepdive slide from slide11 copy
    new=build_deepdive(data["ppt/slides/slide11.xml"].decode("utf-8")); ET.fromstring(new)
    data["ppt/slides/slide15.xml"]=new.encode("utf-8")
    data["ppt/slides/_rels/slide15.xml.rels"]=data["ppt/slides/_rels/slide11.xml.rels"]
    # update conclusions
    data["ppt/slides/slide13.xml"]=update_slide13(data["ppt/slides/slide13.xml"].decode("utf-8")).encode("utf-8")
    data["ppt/slides/slide14.xml"]=update_slide14(data["ppt/slides/slide14.xml"].decode("utf-8")).encode("utf-8")
    # Content_Types
    ct=data["[Content_Types].xml"].decode("utf-8").replace("</Types>",'<Override PartName="/ppt/slides/slide15.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>')
    data["[Content_Types].xml"]=ct.encode("utf-8")
    # presentation rels: next free rId
    pr=data["ppt/_rels/presentation.xml.rels"].decode("utf-8")
    NEWRID="rId%d"%(max(int(n) for n in re.findall(r'Id="rId(\d+)"',pr))+1)
    pr=pr.replace("</Relationships>",f'<Relationship Id="{NEWRID}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide15.xml"/></Relationships>')
    data["ppt/_rels/presentation.xml.rels"]=pr.encode("utf-8")
    # sldIdLst insert after slide11's sldId
    rid_s11=[i for i,t in re.findall(r'Id="(rId\d+)"[^>]*Target="(slides/slide\d+\.xml)"',data["ppt/_rels/presentation.xml.rels"].decode()) if t=="slides/slide11.xml"][0]
    pres=data["ppt/presentation.xml"].decode("utf-8")
    anchor=re.search(r'<p:sldId id="\d+" r:id="'+rid_s11+r'"/>',pres).group(0)
    pres=pres.replace(anchor,anchor+f'<p:sldId id="290" r:id="{NEWRID}"/>')
    data["ppt/presentation.xml"]=pres.encode("utf-8")
    # footer renumber to final positions
    finalpos={"slide9":"9","slide10":"10","slide11":"11","slide15":"12","slide12":"13","slide13":"14","slide14":"15"}
    for i in range(2,9): finalpos[f"slide{i}"]=str(i)
    for fn,num in finalpos.items():
        k=f"ppt/slides/{fn}.xml"; data[k]=set_footer(data[k].decode("utf-8"),num).encode("utf-8")
    for k in ["ppt/slides/slide15.xml","ppt/slides/slide13.xml","ppt/slides/slide14.xml","ppt/presentation.xml"]:
        ET.fromstring(data[k])
    if os.path.exists(DST): os.remove(DST)
    order=names+["ppt/slides/slide15.xml","ppt/slides/_rels/slide15.xml.rels"]
    with zipfile.ZipFile(DST,"w",zipfile.ZIP_DEFLATED) as zo:
        for n in order: zo.writestr(n,data[n])
    print("V19: 출처누락 심화 슬라이드 삽입 + 결론(13·14) 현행화 + 푸터 재번호 완료")
    print("WROTE",DST)

if __name__=="__main__":
    main()
