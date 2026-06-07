# -*- coding: utf-8 -*-
"""V8 deck: 각 거버넌스 이슈/데이터 인텔리전스 슬라이드(10~17) 상단에 '발견 패턴 + 핵심 수치' 배너 추가.
요청: 어떤 패턴에서 찾은 이슈이며 그 수치가 무엇인지를 각 슬라이드 처음에 '가장 크게' 명기.

설계: 슬라이드의 핵심메시지(코어, 그리드 1.28″ 영역)를 '메트릭 배너'로 교체.
 - 배너 = 연한 띠 + ① '발견 패턴 ▸ …'(작게) + ② 핵심 수치(36pt, 제목보다 큼=가장 큼) + ③ 설명(기존 코어 문구 보존).
 - 기존 코어 텍스트는 추출해 배너 설명줄로 재사용(사용자가 다듬은 10p 문구 등 보존).
 거버넌스=Navy, 인텔리전스/검증=Green. 본문(2.0″~)은 손대지 않음(배너는 1.22~1.96″ 코어 영역만 사용).

수치 근거(보고서/데이터): G1 100%(유아및초중등 100/100, lift6.89) · G2 9,840컬럼×3지표=0 · G3 상수컬럼1,372(13.9%)
 · 패밀리 11테이블({순번·의료기관명·전화·주소}) · 연관규칙 lift124.9(교과목코드↔분반) · 경도→위도 44테이블
 · 재식별 lift46.4(의료기관전화→{명,주소}) · 규칙#4 conf 1.00→0.44(도로명→지번).

방식(안전): Base=V7. slide10~17만 편집(코어 제거 → 배너 삽입), 마지막에 한글잘림 방지. 그 외 슬라이드 byte-동일.
"""
import zipfile, os, re

SRC=r"d:\DataInt_2\report\데이터지도_분석보고서_V7.pptx"
DST=r"d:\DataInt_2\report\데이터지도_분석보고서_V8.pptx"
def EMU(i): return int(round(i*914400))
NAVY="081F46"; GREEN="00965F"; DGRAY="505050"; BODYTX="282D37"; WHITE="FFFFFF"
LIGHT="EEF2F8"; LGREEN="E8F5F0"; LINE="D0D8E4"; LGLINE="C7DED3"; FONT="맑은 고딕"
MARGIN=EMU(0.62); CONTENT_W=EMU(12.09)

_id=[600]
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
def _para(runs,algn="l",lnspc=110000):
    rs="".join(_run(t,sz,bb,c) for (t,sz,bb,c) in runs)
    return (f'<a:p><a:pPr algn="{algn}"><a:lnSpc><a:spcPct val="{lnspc}"/></a:lnSpc>'
            f'<a:spcBef><a:spcPts val="0"/></a:spcBef><a:spcAft><a:spcPts val="0"/></a:spcAft></a:pPr>{rs}</a:p>')
def tbox(x,y,w,h,paras,anchor="t",lins=12700,name="txt"):
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{nid()}" name="{name}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
            f'<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
            f'<p:txBody><a:bodyPr wrap="square" lIns="{lins}" tIns="6350" rIns="{lins}" bIns="6350" anchor="{anchor}"><a:normAutofit/></a:bodyPr>'
            f'<a:lstStyle/>{"".join(paras)}</p:txBody></p:sp>')

# slide -> (발견 패턴, 핵심 수치, accent, bg)
CFG={
 10:("A1 비자명 연관규칙 (lift 6.89·conf 1.0)","100%",NAVY,LIGHT),
 11:("컬럼 분포 통계 ▸ 품질·활용 지표 3종 전부 0","9,840개",NAVY,LIGHT),
 12:("distinct=1 스캔 · col_nm 빈도 집계","1,372개",NAVY,LIGHT),
 13:("FP-Growth 최대빈발 컬럼군 (≥8테이블)","11개 테이블",GREEN,LGREEN),
 14:("컬럼 장바구니 연관규칙 (FP-Growth·lift)","lift 124.9",GREEN,LGREEN),
 15:("연관규칙 support–lift 분포 (경도→위도)","44개 테이블",GREEN,LGREEN),
 16:("의료기관 패밀리 + 연관규칙 (conf 0.88)","lift 46.4",GREEN,LGREEN),
 17:("규칙#4 국소→전체 재계산 (도로명→지번)","1.00 → 0.44",GREEN,LGREEN),
}

def banner(pattern,number,desc,accent,bg,line):
    BY=EMU(1.22); BH=EMU(0.74)
    out=[roundrect(MARGIN,BY,CONTENT_W,BH,bg,line,name="metric-bg")]
    out.append(tbox(MARGIN+EMU(0.26),BY+EMU(0.05),CONTENT_W-EMU(0.5),EMU(0.20),
        [_para([("발견 패턴  ▸  ",1000,True,accent),(pattern,1000,True,DGRAY)])],anchor="t",lins=EMU(0.04),name="metric-pattern"))
    out.append(tbox(MARGIN+EMU(0.24),BY+EMU(0.245),EMU(3.45),EMU(0.46),
        [_para([(number,3600,True,accent)],algn="l")],anchor="ctr",lins=EMU(0.04),name="metric-number"))
    out.append(tbox(MARGIN+EMU(3.82),BY+EMU(0.245),CONTENT_W-EMU(4.05),EMU(0.46),
        [_para([(desc,1150,False,BODYTX)],algn="l")],anchor="ctr",lins=EMU(0.06),name="metric-desc"))
    return "".join(out)

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
    log=[]
    for snum,(pattern,number,accent,bg) in CFG.items():
        name=f"ppt/slides/slide{snum}.xml"
        xml=data[name].decode("utf-8")
        core=[None]
        def rm(m):
            blk=m.group(0); mo=re.search(r'<a:off x="\d+" y="(\d+)"/>',blk)
            if mo and 1100000<=int(mo.group(1))<1800000:
                core[0]="".join(re.findall(r'<a:t>(.*?)</a:t>',blk,re.DOTALL)); return ''
            return blk
        xml=re.sub(r'<p:sp>.*?</p:sp>',rm,xml,flags=re.DOTALL)
        desc=core[0] if core[0] else ""
        line=LINE if accent==NAVY else LGLINE
        assert xml.count("</p:spTree>")==1
        xml=xml.replace("</p:spTree>",banner(pattern,number,desc,accent,bg,line)+"</p:spTree>")
        xml=fix_clipping(xml)
        data[name]=xml.encode("utf-8")
        log.append((snum,number,(desc[:34]+'…') if len(desc)>34 else desc, core[0] is not None))
    if os.path.exists(DST): os.remove(DST)
    with zipfile.ZipFile(DST,"w",zipfile.ZIP_DEFLATED) as zout:
        for i in order: zout.writestr(i, data[i.filename])
    for snum,number,desc,found in log:
        print(f"  slide{snum}: 수치='{number}'  core추출={found}  desc='{desc}'")
    print("WROTE",DST)

if __name__=="__main__":
    main()
