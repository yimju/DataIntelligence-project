# -*- coding: utf-8 -*-
"""V15: 발견 섹션(9p~)을 '리스트 → 상세 → 기타' 3단 구조로 재편.
- slide9(종합) → '발견 패턴·규칙 리스트'(표)로 재작성
- slide11(값지문) → '기타 분석 결과'(컬럼 표준·값지문 / 상수·스키마불일치 / 교차도메인값)로 재작성, 배너 제거
- presentation.xml sldIdLst 순서 재배열(새 파트 추가 없음): 리스트→상세(G1·재식별·패밀리·규칙·시각화·규칙#4)→기타
- 영향 푸터 번호만 갱신: slide15(재식별) 15→11, slide16(규칙#4) 16→15, slide11(기타) 11→16

새 위치(position : file):
 9 slide9(리스트) · 10 slide10(G1) · 11 slide15(재식별) · 12 slide12(패밀리) · 13 slide13(규칙)
 14 slide14(시각화) · 15 slide16(규칙#4) · 16 slide11(기타) · 17~20 동일

방식(안전·CLAUDE.md §3): Base=V14. 문자열 편집·런 병합·fix_clipping. 파트 추가/삭제 없음(정합 위험 최소).
"""
import zipfile, os, re
import xml.etree.ElementTree as ET

SRC=r"d:\data_intelligence_2\report\데이터지도_분석보고서_V14.pptx"
DST=r"d:\data_intelligence_2\report\데이터지도_분석보고서_V15.pptx"
def EMU(i): return int(round(i*914400))
NAVY="081F46"; GREEN="00965F"; DGRAY="505050"; BODYTX="282D37"; WHITE="FFFFFF"
LIGHT="EEF2F8"; LGREEN="E8F5F0"; LINE="D0D8E4"; RED="C0392B"; FONT="맑은 고딕"
HEADER_H=EMU(0.36)
_id=[4000]
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
    return [roundrect(x,y,w,h,bg,line,name="card-bg"),
            roundrect(x,y,w,HEADER_H,hdr,hdr,name="card-hdr"),
            tbox(x,y,w,HEADER_H,[_para([(title,1100,True,WHITE)],algn="l",spcaft=0)],anchor="ctr",lins=EMU(0.13),name="card-title")]

# ---------- native table ----------
def _cell(text,sz,bold,color,fill,algn="l"):
    return (f'<a:tc><a:txBody><a:bodyPr wrap="square" lIns="50000" rIns="25000" tIns="10000" bIns="10000" anchor="ctr"/><a:lstStyle/>'
            f'<a:p><a:pPr algn="{algn}"/><a:r><a:rPr lang="ko-KR" sz="{sz}" b="{1 if bold else 0}">'
            f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill><a:latin typeface="{FONT}"/><a:ea typeface="{FONT}"/></a:rPr>'
            f'<a:t>{esc(text)}</a:t></a:r></a:p></a:txBody>'
            f'<a:tcPr><a:solidFill><a:srgbClr val="{fill}"/></a:solidFill></a:tcPr></a:tc>')
def make_table(x,y,w,colws,header,rows,rowh=EMU(0.52),hh=EMU(0.42)):
    grid="".join(f'<a:gridCol w="{c}"/>' for c in colws)
    trs=[f'<a:tr h="{hh}">'+"".join(_cell(h,1050,True,WHITE,NAVY,"ctr" if i>=2 else "l") for i,h in enumerate(header))+'</a:tr>']
    for ri,row in enumerate(rows):
        fill=WHITE if ri%2 else LIGHT; cells=""
        for ci,val in enumerate(row):
            color=NAVY if ci==0 else (GREEN if ci==3 else BODYTX)
            bold=(ci==0 or ci==3); algn="ctr" if ci>=2 else "l"; sz=1000
            cells+=_cell(val,sz,bold,color,fill,algn)
        trs.append(f'<a:tr h="{rowh}">{cells}</a:tr>')
    cy=hh+rowh*len(rows)
    return (f'<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="{nid()}" name="list-tbl"/>'
            f'<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>'
            f'<p:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{cy}"/></p:xfrm>'
            f'<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">'
            f'<a:tbl><a:tblPr firstRow="1" bandRow="1"/><a:tblGrid>{grid}</a:tblGrid>{"".join(trs)}</a:tbl>'
            f'</a:graphicData></a:graphic></p:graphicFrame>')

# ---------- shared edit helpers ----------
def merge_runs(sp,newtext):
    rpr_m=re.search(r'<a:rPr.*?</a:rPr>|<a:rPr[^>]*/>',sp,re.S)
    rpr=rpr_m.group(0) if rpr_m else '<a:rPr/>'
    newrun=f'<a:r>{rpr}<a:t>{esc(newtext)}</a:t></a:r>'
    fr=sp.index('<a:r>'); lr=sp.rindex('</a:r>')+len('</a:r>')
    return sp[:fr]+newrun+sp[lr:]
def edit_sp_by_y(xml,ylo,yhi,newtext):
    for m in re.finditer(r'<p:sp>.*?</p:sp>',xml,re.S):
        sp=m.group(0); mo=re.search(r'<a:off x="-?\d+" y="(-?\d+)"/>',sp)
        if mo and ylo<=int(mo.group(1))<yhi and '<a:r>' in sp:
            return xml[:m.start()]+merge_runs(sp,newtext)+xml[m.end():]
    raise RuntimeError(f"no editable sp in y[{ylo},{yhi})")
def remove_sps(xml,pred):
    return re.sub(r'<p:sp>.*?</p:sp>',lambda m:'' if pred(m.group(0)) else m.group(0),xml,flags=re.S)
def yof(blk):
    mo=re.search(r'<a:off x="-?\d+" y="(-?\d+)"/>',blk); return int(mo.group(1)) if mo else None
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
def add_shapes(xml,shapes):
    assert xml.count("</p:spTree>")==1
    return xml.replace("</p:spTree>","".join(shapes)+"</p:spTree>")

# =================== slide9 -> 리스트 ===================
def build_slide9(xml):
    xml=edit_sp_by_y(xml,0,200000,"FINDINGS")
    xml=edit_sp_by_y(xml,300000,700000,"발견 패턴·규칙 리스트")
    xml=edit_sp_by_y(xml,1200000,1400000,"비지도 분석(FP-Growth·연관규칙) 산출 패턴·규칙을 먼저 나열하고, 상세 분석 대상 페이지를 표시함")
    xml=remove_sps(xml,lambda b:(yof(b) is not None and 1800000<=yof(b)<6000000))
    header=["구분","발견 내용","핵심 지표","상세"]
    rows=[
     ["패턴","스키마 패밀리(의료기관·POI·사학재정·국가시험군)","FP-Growth 최대빈발 ≥8테이블","p.12"],
     ["규칙","경도 → 위도 (POI 좌표쌍)","conf 1.00 · lift 22.7","p.13"],
     ["규칙","교과목코드 ↔ 분반","conf 1.00 · lift 124.9","p.13"],
     ["규칙","시도명 → 시군구명 (행정구역 계층)","conf 0.94 · lift 39.0","p.13"],
     ["거버넌스","출처 결손 — ‘유아및초·중등교육’ 100% placeholder","A1 conf 1.00 · lift 6.89","p.10"],
     ["거버넌스","교차도메인 재식별 — 의료기관전화 → {명·주소}","conf 0.88 · lift 46.4","p.11"],
     ["검증","규칙#4 국소→전체 (도로명 → 지번)","conf 1.00 → 0.44","p.15"],
     ["기타","컬럼 표준·값지문 / 상수·스키마불일치 / 교차도메인값","1D 집계·샘플","p.16"],
    ]
    W=EMU(12.09)
    colws=[EMU(1.30),EMU(6.90),EMU(2.70),EMU(1.19)]
    tbl=make_table(EMU(0.62),EMU(2.05),W,colws,header,rows)
    cap=tbox(EMU(0.62),EMU(6.60),W,EMU(0.3),
             [_para([("‘상세’ 열은 해당 발견을 깊이 다루는 페이지 · 1D 집계·샘플 결과는 ‘기타(p.16)’",950,False,DGRAY)],spcaft=0)],name="list-cap")
    return fix_clipping(add_shapes(xml,[tbl,cap]))

# =================== slide11 -> 기타 ===================
def build_slide11(xml):
    xml=edit_sp_by_y(xml,0,200000,"기타")
    xml=edit_sp_by_y(xml,300000,700000,"기타 분석 결과")
    xml=edit_sp_by_y(xml,6000000,9999999,"16")  # footer renumber
    xml=remove_sps(xml,lambda b:('name="metric-' in b) or (yof(b) is not None and 1800000<=yof(b)<6000000))
    core=tbox(EMU(0.62),EMU(1.16),EMU(12.09),EMU(0.56),
              [_para([("핵심 방법(FP-Growth·연관규칙) 산출이 아닌 1D 집계·샘플 기반 부가 결과",1500,True,DGRAY)],spcaft=0)],
              anchor="ctr",name="etc-core")
    W=EMU(12.09); gap=EMU(0.26); cw=(W-2*gap)//3; y=EMU(2.05); h=EMU(4.35)
    xs=[EMU(0.62),EMU(0.62)+cw+gap,EMU(0.62)+2*(cw+gap)]
    cards=[]
    c1=frame(xs[0],y,cw,h,NAVY,LIGHT,LINE,"① 컬럼 표준·품질 (값 지문)")
    c1.append(tbox(xs[0]+EMU(0.13),y+HEADER_H+EMU(0.06),cw-EMU(0.26),h-HEADER_H-EMU(0.12),[
     _para([("연번·순번·번호",1000,True,NAVY),(" = min1·pk≈1.0",1000,False,BODYTX)],spcaft=4),
     _para([("→ 행 식별자",1000,False,DGRAY)],spcaft=10),
     _para([("데이터기준일자·데이터기준일",1000,True,NAVY),(" = pk≈0.0",1000,False,BODYTX)],spcaft=4),
     _para([("→ 스냅샷 날짜(거의 상수)",1000,False,DGRAY)],spcaft=10),
     _para([("같은 개념·다른 이름 → 값 지문(min·pk)으로 판별·표준화",1000,False,BODYTX)],spcaft=0)],name="etc-c1"))
    c2=frame(xs[1],y,cw,h,NAVY,LIGHT,LINE,"② 상수컬럼 · 스키마 불일치")
    c2.append(tbox(xs[1]+EMU(0.13),y+HEADER_H+EMU(0.06),cw-EMU(0.26),h-HEADER_H-EMU(0.12),[
     _para([("상수 컬럼(distinct=1): ",1000,True,NAVY),("1,372개(13.9%)",1000,True,RED)],spcaft=4),
     _para([("→ 전 행 동일값 = 정보량 0",1000,False,DGRAY)],spcaft=10),
     _para([("스키마 불일치: ",1000,True,NAVY),("56개 테이블",1000,True,RED)],spcaft=4),
     _para([("→ column_cnt ≠ 범주+수치(미분류 컬럼)",1000,False,DGRAY)],spcaft=10),
     _para([("단순 1D 집계·정합성 검사로 발견",1000,False,BODYTX)],spcaft=0)],name="etc-c2"))
    c3=frame(xs[2],y,cw,h,NAVY,LIGHT,LINE,"③ 교차도메인 값 공유")
    c3.append(tbox(xs[2]+EMU(0.13),y+HEADER_H+EMU(0.06),cw-EMU(0.26),h-HEADER_H-EMU(0.12),[
     _para([("3개 도메인 공통 값 상위",1000,True,NAVY)],spcaft=4),
     _para([("= 소수·연도(1·8·2018) = 잡음 지배",1000,False,BODYTX)],spcaft=10),
     _para([("샘플 집계 기반 → 전체검증 필요",1000,False,DGRAY)],spcaft=10),
     _para([("의미 코드만 필터해야 연계 신호 추출 가능",1000,False,BODYTX)],spcaft=0)],name="etc-c3"))
    return fix_clipping(add_shapes(xml,[core]+c1+c2+c3))

# =================== main ===================
def main():
    with zipfile.ZipFile(SRC) as z:
        order=z.infolist(); data={i.filename:z.read(i.filename) for i in order}
    # 1) slide9 -> 리스트
    data["ppt/slides/slide9.xml"]=build_slide9(data["ppt/slides/slide9.xml"].decode("utf-8")).encode("utf-8")
    # 2) slide11 -> 기타
    data["ppt/slides/slide11.xml"]=build_slide11(data["ppt/slides/slide11.xml"].decode("utf-8")).encode("utf-8")
    # 3) footer renumber on moved slides: slide15 15->11, slide16 16->15
    for fn,new in [("ppt/slides/slide15.xml","11"),("ppt/slides/slide16.xml","15")]:
        x=data[fn].decode("utf-8"); x=edit_sp_by_y(x,6000000,9999999,new); data[fn]=x.encode("utf-8")
    # 4) reorder sldIdLst
    pres=data["ppt/presentation.xml"].decode("utf-8")
    sldids=re.findall(r'<p:sldId [^>]*/>',pres)
    byrid={re.search(r'r:id="(rId\d+)"',e).group(1):e for e in sldids}
    target=['rId2','rId3','rId4','rId5','rId6','rId7','rId8','rId9','rId10','rId11',
            'rId16','rId13','rId14','rId15','rId17','rId12','rId18','rId19','rId20','rId21']
    assert set(target)==set(byrid), "rId set mismatch"
    newlst="".join(byrid[r] for r in target)
    pres=re.sub(r'(<p:sldIdLst>).*?(</p:sldIdLst>)',lambda m:m.group(1)+newlst+m.group(2),pres,flags=re.S)
    data["ppt/presentation.xml"]=pres.encode("utf-8")

    # validate edited slides
    for fn in ["ppt/slides/slide9.xml","ppt/slides/slide11.xml","ppt/slides/slide15.xml","ppt/slides/slide16.xml","ppt/presentation.xml"]:
        ET.fromstring(data[fn]);
        ids=re.findall(r'<p:cNvPr id="(\d+)"',data[fn].decode("utf-8"))
        assert len(ids)==len(set(ids)), f"dup id in {fn}"
    if os.path.exists(DST): os.remove(DST)
    with zipfile.ZipFile(DST,"w",zipfile.ZIP_DEFLATED) as zout:
        for i in order: zout.writestr(i,data[i.filename])
    print("OK reordered + slide9(리스트)/slide11(기타) rebuilt + footers renumbered")
    print("WROTE",DST)

if __name__=="__main__":
    main()
