# -*- coding: utf-8 -*-
"""V18: 사용자 업데이트본 V17(15장) 보존 + 9페이지(연관규칙 상세 지표) 뒤에 새 슬라이드 1장 삽입.
새 슬라이드 = '탐사 대상 규칙 (적절한 lift 구간)' — lift가 매우 높은 자명한 규칙(예 교과목코드↔분반 125)이 아니라
적절한 구간(lift≈4~50)의 비자명·탐사가치 규칙을 9페이지와 동일 양식(배너+표)으로 표현.
포함: 현재 10페이지 항목 '유아및초·중등교육 → 기관=미상(소스)'(conf 1.00·lift 6.89).

데이터 근거(직접계산): 출판사→저자 47.0 / 확진자수→사망자수 39.0 / 시도명→시군구명 39.0 / 연령대→성별 35.2 /
 설립→학교명 21.4 / 유아및초중등→기관미상 6.89 / {보건·유일키·협소}→{보건의료·상수없음} 4.5

방식(안전): Base=V17. slide9.xml 복사→slide16.xml(제목·섹션·배너·푸터·표행 치환). 파트/Content_Types/rels/sldIdLst 추가.
이후 위치 슬라이드 푸터 재번호(9→9·신규→10·구10~15→11~16). ET 검증.
"""
import zipfile, os, re
import xml.etree.ElementTree as ET
SRC=r"d:\data_intelligence_2\report\데이터지도_분석보고서_V17.pptx"
DST=r"d:\data_intelligence_2\report\데이터지도_분석보고서_V18.pptx"
def esc(s): return s.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")

# ---- explore-worthy rules (적절 lift 구간, 자명 초고-lift 제외) ----
ROWS=[
 ["도서관","출판사 → 저자","12","0.012","0.71","47.0","0.0118","3.3"],
 ["보건통계","확진자수 → 사망자수","10","0.010","0.62","39.0","0.0098","2.6"],
 ["행정구역","시도명 → 시군구명","15","0.015","0.94","39.0","0.0146","15.6"],
 ["인구통계","연령대 → 성별","11","0.011","0.85","35.2","0.0107","6.3"],
 ["교육","설립 → 학교명","9","0.009","0.90","21.4","0.0086","9.6"],
 ["거버넌스","유아및초·중등교육 → 기관=미상(소스)","100","0.100","1.00","6.89","0.0856","∞"],
 ["속성원형","{보건·유일키·협소폭} → {보건의료·상수없음}","68","0.068","0.65","4.5","0.0530","2.4"],
]
NEW_TITLE="탐사 대상 규칙 — 적절한 lift 구간 (초고-lift 자명규칙 제외)"
NEW_SECTION="탐사 대상"
NEW_DESC="발견 패턴 ▸ lift 매우 높음(예 교과목코드↔분반 125)은 단일 템플릿 내 자명한 동반 → 적절 구간(lift 4~50)의 비자명 규칙을 탐사"

def merge_runs(sp,newtext):
    m=re.search(r'<a:rPr.*?</a:rPr>|<a:rPr[^>]*/>',sp,re.S); rpr=m.group(0) if m else '<a:rPr/>'
    fr=sp.index('<a:r>'); lr=sp.rindex('</a:r>')+len('</a:r>')
    return sp[:fr]+f'<a:r>{rpr}<a:t>{esc(newtext)}</a:t></a:r>'+sp[lr:]
def edit_sp_by_y(xml,lo,hi,txt):
    for m in re.finditer(r'<p:sp>.*?</p:sp>',xml,re.S):
        sp=m.group(0); mo=re.search(r'<a:off x="-?\d+" y="(-?\d+)"/>',sp)
        if mo and lo<=int(mo.group(1))<hi and '<a:r>' in sp:
            return xml[:m.start()]+merge_runs(sp,txt)+xml[m.end():]
    raise RuntimeError(f"no sp y[{lo},{hi})")
def edit_sp_by_name(xml,name,txt):
    m=re.search(r'<p:sp>(?:(?!</p:sp>).)*name="'+re.escape(name)+r'".*?</p:sp>',xml,re.S)
    if not m: raise RuntimeError("no "+name)
    return xml[:m.start()]+merge_runs(m.group(0),txt)+xml[m.end():]
def set_footer(xml,num):
    return edit_sp_by_y(xml,6200000,9999999,num)

def fill_row(tmpl,values):
    parts=re.split(r'<a:tc>.*?</a:tc>',tmpl,flags=re.S)
    cells=re.findall(r'<a:tc>.*?</a:tc>',tmpl,re.S)
    res=parts[0]
    for i,cell in enumerate(cells):
        if i<len(values) and '<a:t>' in cell:
            cell=re.sub(r'<a:t>[^<]*</a:t>','<a:t>'+esc(values[i])+'</a:t>',cell,count=1)
        res+=cell+parts[i+1]
    return res

def build_new_slide(slide9xml):
    x=slide9xml
    x=edit_sp_by_y(x,0,200000,NEW_SECTION)
    x=edit_sp_by_y(x,300000,750000,NEW_TITLE)
    x=edit_sp_by_name(x,"metric-desc",NEW_DESC)
    x=set_footer(x,"10")
    # replace table data rows
    tm=re.search(r'<a:tbl>.*?</a:tbl>',x,re.S); tbl=tm.group(0)
    trs=re.findall(r'<a:tr\b.*?</a:tr>',tbl,re.S)
    header=trs[0]; tA=trs[1]; tB=trs[2] if len(trs)>2 else trs[1]
    newrows=[header]
    for i,vals in enumerate(ROWS):
        newrows.append(fill_row(tA if i%2==0 else tB, vals))
    # rebuild tbl: keep everything before first <a:tr>, then rows, then closing
    pre=tbl[:tbl.index('<a:tr')];
    newtbl=pre+"".join(newrows)+'</a:tbl>'
    x=x[:tm.start()]+newtbl+x[tm.end():]
    return x

def main():
    with zipfile.ZipFile(SRC) as z:
        names=z.namelist(); data={n:z.read(n) for n in names}
    s9=data["ppt/slides/slide9.xml"].decode("utf-8")
    new=build_new_slide(s9)
    ET.fromstring(new)
    # add new parts
    data["ppt/slides/slide16.xml"]=new.encode("utf-8")
    data["ppt/slides/_rels/slide16.xml.rels"]=data["ppt/slides/_rels/slide9.xml.rels"]
    # Content_Types
    ct=data["[Content_Types].xml"].decode("utf-8")
    ct=ct.replace("</Types>",'<Override PartName="/ppt/slides/slide16.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>')
    data["[Content_Types].xml"]=ct.encode("utf-8")
    # presentation rels — use next free rId across ALL relationships
    pr=data["ppt/_rels/presentation.xml.rels"].decode("utf-8")
    nextid=max(int(n) for n in re.findall(r'Id="rId(\d+)"',pr))+1
    NEWRID=f"rId{nextid}"
    pr=pr.replace("</Relationships>",f'<Relationship Id="{NEWRID}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide16.xml"/></Relationships>')
    data["ppt/_rels/presentation.xml.rels"]=pr.encode("utf-8")
    # sldIdLst insert after rId10 (slide9, pos9)
    pres=data["ppt/presentation.xml"].decode("utf-8")
    anchor=re.search(r'<p:sldId id="\d+" r:id="rId10"/>',pres).group(0)
    assert pres.count(anchor)==1
    pres=pres.replace(anchor, anchor+f'<p:sldId id="277" r:id="{NEWRID}"/>')
    data["ppt/presentation.xml"]=pres.encode("utf-8")
    # footer renumber: slide9->9, slide10->11, slide11->12, slide12->13, slide13->14, slide14->15, slide15->16
    for fn,num in [("slide9","9"),("slide10","11"),("slide11","12"),("slide12","13"),("slide13","14"),("slide14","15"),("slide15","16")]:
        k=f"ppt/slides/{fn}.xml"; data[k]=set_footer(data[k].decode("utf-8"),num).encode("utf-8")
    # validate edited slides
    for k in ["ppt/slides/slide16.xml","ppt/presentation.xml"]+[f"ppt/slides/slide{i}.xml" for i in [9,10,11,12,13,14,15]]:
        ET.fromstring(data[k])
    if os.path.exists(DST): os.remove(DST)
    order=names+["ppt/slides/slide16.xml","ppt/slides/_rels/slide16.xml.rels"]
    with zipfile.ZipFile(DST,"w",zipfile.ZIP_DEFLATED) as zo:
        for n in order: zo.writestr(n,data[n])
    print("V18: 새 슬라이드(탐사 대상 규칙) 9 뒤 삽입 + 푸터 재번호 완료")
    print("WROTE",DST)

if __name__=="__main__":
    main()
