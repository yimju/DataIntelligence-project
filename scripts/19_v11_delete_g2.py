# -*- coding: utf-8 -*-
"""V11 deck: '거버넌스 G2 죽은 메트릭' 항목(슬라이드 11) 삭제 + 보고서 정합성 업데이트.

삭제 절차(슬라이드 1장 제거 = 4개 파트 정리):
 1) [Content_Types].xml 에서 slide11 Override 제거
 2) ppt/presentation.xml 의 sldIdLst 에서 slide11 sldId(rId12) 제거
 3) ppt/_rels/presentation.xml.rels 에서 rId12 Relationship 제거
 4) zip 에서 ppt/slides/slide11.xml + 그 .rels 제외(미기록)
정합성 업데이트:
 5) 페이지번호 재번호: 기존 slide12~21(파일)은 위치가 11~20으로 당겨지므로 푸터 번호 12→11 … 21→20
 6) slide9(종합): 코어 ‘3대 이슈’→‘2대 이슈’, 카드 ‘거버넌스 G2·G3’→‘거버넌스 G3’, ‘죽은 메트릭’ bullet 삭제
 7) slide20(로드맵): ‘죽은 메트릭·네이밍 정정’→‘메타 네이밍·상수컬럼 정정’
docProps/app.xml(슬라이드 수 메타)은 PowerPoint가 저장 시 자동 보정 → 무손상 위해 미변경.
"""
import zipfile, os, re

SRC=r"d:\DataInt_2\report\데이터지도_분석보고서_V10.pptx"
DST=r"d:\DataInt_2\report\데이터지도_분석보고서_V11.pptx"
DROP={"ppt/slides/slide11.xml","ppt/slides/_rels/slide11.xml.rels"}
def esc(s): return s.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")

def set_para_text(block,new):
    cnt=[0]
    def t(m):
        cnt[0]+=1
        return f"<a:t>{esc(new)}</a:t>" if cnt[0]==1 else "<a:t></a:t>"
    return re.sub(r"<a:t>.*?</a:t>",t,block,flags=re.DOTALL)

def para_sub(xml,fn):
    def repl(m):
        block=m.group(0); text="".join(re.findall(r"<a:t>(.*?)</a:t>",block,re.DOTALL))
        nb=fn(block,text); return nb if nb is not None else block
    return re.sub(r"<a:p>.*?</a:p>",repl,xml,flags=re.DOTALL)

def dec_page(xml):
    def repl(m):
        block=m.group(0)
        if 'y="6437376"' not in block: return block
        cur="".join(re.findall(r"<a:t>(.*?)</a:t>",block,re.DOTALL)).strip()
        try: newn=str(int(cur)-1)
        except: return block
        return set_para_text(block,newn)
    return re.sub(r"<p:sp>.*?</p:sp>",repl,xml,flags=re.DOTALL)

def fn9(block,text):
    if "3대 이슈" in text: return set_para_text(block,"거버넌스 2대 이슈와 인텔리전스 4대 발견을 도출")
    if "G2" in text:       return set_para_text(block,"거버넌스 G3")
    if "죽은 메트릭" in text: return ""   # remove bullet paragraph
    return None
def fn20(block,text):
    if "죽은 메트릭" in text: return set_para_text(block,"• 메타 네이밍·상수컬럼 정정")
    return None

def main():
    with zipfile.ZipFile(SRC) as z:
        order=z.infolist(); data={i.filename:z.read(i.filename) for i in order}

    # 1) Content_Types
    ct=data["[Content_Types].xml"].decode("utf-8")
    ovr='<Override PartName="/ppt/slides/slide11.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
    assert ovr in ct; ct=ct.replace(ovr,""); data["[Content_Types].xml"]=ct.encode("utf-8")
    # 2) presentation.xml sldId
    pres=data["ppt/presentation.xml"].decode("utf-8")
    sld='<p:sldId id="266" r:id="rId12"/>'
    assert sld in pres; pres=pres.replace(sld,""); data["ppt/presentation.xml"]=pres.encode("utf-8")
    # 3) presentation rels
    rels=data["ppt/_rels/presentation.xml.rels"].decode("utf-8")
    before=rels
    rels=re.sub(r'<Relationship Id="rId12"[^>]*Target="slides/slide11\.xml"[^>]*/>',"",rels)
    assert rels!=before; data["ppt/_rels/presentation.xml.rels"]=rels.encode("utf-8")
    # 5) page renumber on files slide12..21
    for n in range(12,22):
        f=f"ppt/slides/slide{n}.xml"; data[f]=dec_page(data[f].decode("utf-8")).encode("utf-8")
    # 6) slide9, 7) slide20
    data["ppt/slides/slide9.xml"]=para_sub(data["ppt/slides/slide9.xml"].decode("utf-8"),fn9).encode("utf-8")
    data["ppt/slides/slide20.xml"]=para_sub(data["ppt/slides/slide20.xml"].decode("utf-8"),fn20).encode("utf-8")

    # 4) write without slide11
    if os.path.exists(DST): os.remove(DST)
    with zipfile.ZipFile(DST,"w",zipfile.ZIP_DEFLATED) as zout:
        for i in order:
            if i.filename in DROP: continue
            zout.writestr(i, data[i.filename])
    print("removed slide11 (G2) + rels; updated CT/presentation/rels; renumbered 12-21; cleaned slide9/20")
    print("WROTE",DST)

if __name__=="__main__":
    main()
