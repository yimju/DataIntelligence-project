# -*- coding: utf-8 -*-
"""V21 (Base=사용자 정리본 V20, 15장): 14페이지(보건 협소 테이블 → 데이터 병합 기회)의
① 카드만 '동형(同型) 판단 방법 + Jaccard 유사도' 설명으로 교체. 나머지(제목·배너·②·③표·사용자 편집)는 전부 보존.

설명(데이터/방법 근거):
 - 동형 판단: 각 테이블 컬럼명 정규화 → '컬럼 집합' → 두 집합 비교(완전일치 또는 Jaccard≥0.8)
 - Jaccard = |A∩B|/|A∪B| = 공통컬럼/합친컬럼 (0~1). 예) 의료기관현황 A{순번·명·전화·주소·종별} vs B{…주소}: 4/5=0.80
 - 왜 Jaccard: 완전일치만 보면 컬럼 1개 차이로 놓침 → ≥0.8을 '동형'으로 보아 거의 같은 표까지 포착(순서·개수 무관, 합집합 정규화로 큰 표 편향 보정)
 - 결과: 105개 중 완전일치 3그룹(11개)+동형 41쌍

방식(안전): slide14.xml만 편집. ①의 f-title(좌상단 x=566928,y=2155971) 텍스트 교체 + m-ver 본문 문단 교체. 그 외 byte-동일. ET 검증.
"""
import zipfile, os, re
import xml.etree.ElementTree as ET
SRC=r"d:\data_intelligence_2\report\데이터지도_분석보고서_V20.pptx"
DST=r"d:\data_intelligence_2\report\데이터지도_분석보고서_V21.pptx"
SLIDE="ppt/slides/slide14.xml"
NAVY="081F46"; GREEN="00965F"; DGRAY="505050"; BODYTX="282D37"; RED="C0392B"; FONT="맑은 고딕"
def esc(s): return s.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")
def _run(t,sz,b_,c):
    return (f'<a:r><a:rPr sz="{sz}" b="{1 if b_ else 0}" i="0"><a:solidFill><a:srgbClr val="{c}"/></a:solidFill>'
            f'<a:latin typeface="{FONT}"/><a:ea typeface="{FONT}"/></a:rPr><a:t>{esc(t)}</a:t></a:r>')
def _para(runs,spcaft=8):
    return (f'<a:p><a:pPr algn="l"><a:lnSpc><a:spcPct val="108000"/></a:lnSpc>'
            f'<a:spcBef><a:spcPts val="0"/></a:spcBef><a:spcAft><a:spcPts val="{spcaft}"/></a:spcAft><a:buNone/></a:pPr>'
            f'{"".join(_run(t,sz,bb,c) for t,sz,bb,c in runs)}</a:p>')
def merge_runs(sp,txt):
    m=re.search(r'<a:rPr.*?</a:rPr>|<a:rPr[^>]*/>',sp,re.S); rpr=m.group(0) if m else '<a:rPr/>'
    fr=sp.index('<a:r>'); lr=sp.rindex('</a:r>')+len('</a:r>')
    return sp[:fr]+f'<a:r>{rpr}<a:t>{esc(txt)}</a:t></a:r>'+sp[lr:]

def edit_ftitle_topleft(xml,newtext):
    for m in re.finditer(r'<p:sp>(?:(?!</p:sp>).)*?name="f-title".*?</p:sp>',xml,re.S):
        sp=m.group(0); off=re.search(r'<a:off x="(-?\d+)" y="(-?\d+)"/>',sp)
        if off and off.group(1)=="566928" and off.group(2)=="2155971":
            return xml[:m.start()]+merge_runs(sp,newtext)+xml[m.end():]
    raise RuntimeError("① f-title not found")
def set_paras(xml,name,paras):
    m=re.search(r'<p:sp>(?:(?!</p:sp>).)*?name="'+re.escape(name)+r'".*?</p:sp>',xml,re.S)
    if not m: raise RuntimeError("no "+name)
    sp=m.group(0); fp=sp.index('<a:p>'); lp=sp.rindex('</a:p>')+len('</a:p>')
    return xml[:m.start()]+sp[:fp]+"".join(paras)+sp[lp:]+xml[m.end():]
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

PARAS=[
 _para([("방법: ",950,True,NAVY),("각 테이블 컬럼명을 정규화 → ‘컬럼 집합’으로 만들어 두 집합을 비교",950,False,BODYTX)]),
 _para([("Jaccard = |A∩B| ÷ |A∪B|",950,True,NAVY),(" = 공통 컬럼 ÷ 합친 컬럼 (0~1)",950,False,BODYTX)]),
 _para([("예) ",950,True,DGRAY),("A{순번·명·전화·주소·종별} · B{순번·명·전화·주소} → 4÷5 = 0.80",950,False,BODYTX)]),
 _para([("왜 Jaccard? ",950,True,GREEN),("완전일치(1.0)만 보면 컬럼 1개 차이로 놓침 → ≥0.8을 ‘동형’으로 보아 거의 같은 표까지 포착 (순서·개수 무관, 합집합 정규화로 큰 표 편향 보정)",950,False,BODYTX)]),
 _para([("결과: ",950,True,NAVY),("105개 중 완전일치 3그룹(11개) + 동형(≥0.8) 41쌍",950,True,RED)],spcaft=0),
]

def main():
    with zipfile.ZipFile(SRC) as z:
        names=z.namelist(); data={n:z.read(n) for n in names}
    x=data[SLIDE].decode("utf-8")
    x=edit_ftitle_topleft(x,"① 동형(同型) 판단 — Jaccard 유사도")
    x=set_paras(x,"m-ver",PARAS)
    x=fix_clipping(x)
    ET.fromstring(x)
    ids=re.findall(r'<p:cNvPr id="(\d+)"',x); assert len(ids)==len(set(ids)),"dup id"
    data[SLIDE]=x.encode("utf-8")
    if os.path.exists(DST): os.remove(DST)
    with zipfile.ZipFile(DST,"w",zipfile.ZIP_DEFLATED) as zo:
        for n in names: zo.writestr(n,data[n])
    print("V21: 14p ① 카드 = 동형 판단/Jaccard 설명으로 교체 (그 외 보존)")
    print("WROTE",DST)

if __name__=="__main__":
    main()
