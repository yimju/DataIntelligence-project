# -*- coding: utf-8 -*-
"""V5 deck: 7페이지(AOI 롤업) 과정을 자세히 — 롤업 로직 텍스트 보강 + 과정 도표 추가.

추가 내용:
 - 핵심메시지/bullet을 'AOI(속성지향 귀납) 롤업 = 구체값→상위개념 다계층 일반화' 관점으로 보강.
 - 우측에 3단계 과정 도표(① 원시 table_source → ② 정규식 접미사·키워드 규칙 → ③ 7개 기관유형/개수).
 - 기존 fig_aoi.png(scripts/09) 제거: 그 그림은 09의 정규식이라 개수가 254/29/35로 보고서와 불일치 →
   도표는 분석 파이프라인(scripts/06)의 규칙으로 재계산(기초404·공공245·미상145·광역83·교육49·중앙39·기타34,
   I1_REPORT §0.2와 정확히 일치)하여 표기 → 슬라이드를 보고서 근거에 맞게 교정.

Base=V4. slide7.xml만 문자열 편집(기존 바이트 보존) + python-pptx 검증 도형구조 복제(<p:style> 포함,
채운 도형은 빈 txBody·텍스트는 별도 박스)로 도표 삽입. 다른 슬라이드는 V4와 byte-동일.
"""
import zipfile, os, re, csv, io
from collections import Counter

SRC=r"d:\DataInt_2\report\데이터지도_분석보고서_V4.pptx"
DST=r"d:\DataInt_2\report\데이터지도_분석보고서_V5.pptx"
TM =r"d:\DataInt_2\data\table_map.csv"
SLIDE="ppt/slides/slide7.xml"

def EMU(i): return int(round(i*914400))
NAVY="081F46"; GREEN="00965F"; DGRAY="505050"; BODYTX="282D37"; WHITE="FFFFFF"
LIGHT="EEF2F8"; LGREEN="E8F5F0"; LINE="D0D8E4"; LGLINE="C7DED3"; FONT="맑은 고딕"
CORE_Y=str(EMU(1.28))

# ---------- AOI counts via scripts/06 rule (== report §0.2) ----------
def src_type(s):
    if s is None or s=="" or s=="소스": return "미상(소스)"
    if re.search(r"(교육청|교육지원청|교육연수원|교육원|교육연구원|도서관|대학교|대학원|폴리텍|학교$)",s): return "교육기관"
    if re.search(r"(부$|청$|처$|위원회$|^교육부|^보건복지부|^행정안전부|^고용노동부|^국방부|^법무부)",s): return "중앙행정"
    if re.search(r"(공단|공사|진흥원|개발원|재단|연구원|시험원|평가원|관리공단|센터$|기금|공제회|진흥회|연수원|병원|암센터|진흥재단|장학재단|보험공단)",s): return "공공기관"
    if re.search(r"(특별시$|광역시$|특별자치시$|도$|특별자치도$)",s): return "광역지자체"
    if re.search(r"(시$|군$|구$)",s): return "기초지자체"
    return "기타"
rows=list(csv.DictReader(io.open(TM,encoding="utf-8-sig")))
C=Counter(src_type((r.get("table_source") or "").strip()) for r in rows)
TYPE_ORDER=["기초지자체","공공기관","미상(소스)","광역지자체","교육기관","중앙행정","기타"]

# ---------- PowerPoint-safe shape builders (mimic python-pptx) ----------
_id=[300]
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
def _para(text,sz,b,color,algn="l",lnspc=120000,spcaft=0):
    return (f'<a:p><a:pPr algn="{algn}"><a:lnSpc><a:spcPct val="{lnspc}"/></a:lnSpc>'
            f'<a:spcBef><a:spcPts val="0"/></a:spcBef><a:spcAft><a:spcPts val="{spcaft}"/></a:spcAft></a:pPr>'
            f'<a:r><a:rPr sz="{sz}" b="{1 if b else 0}" i="0"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
            f'<a:latin typeface="{FONT}"/><a:ea typeface="{FONT}"/></a:rPr><a:t>{text}</a:t></a:r></a:p>')
def tbox(x,y,w,h,lines,anchor="t",algn="l",lins=12700,lnspc=120000,spcaft=0,name="txt"):
    paras="".join(_para(t,sz,b,c,algn,lnspc,spcaft) for (t,sz,b,c) in lines)
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{nid()}" name="{name}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
            f'<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
            f'<p:txBody><a:bodyPr wrap="square" lIns="{lins}" tIns="6350" rIns="{lins}" bIns="6350" anchor="{anchor}"/>'
            f'<a:lstStyle/>{paras}</p:txBody></p:sp>')

HEADER_H=EMU(0.34)
def card(x,y,w,h,hdr_color,bg,line,title,body_lines):
    out=[roundrect(x,y,w,h,bg,line,name="card-bg"),
         roundrect(x,y,w,HEADER_H,hdr_color,hdr_color,name="card-hdr"),
         tbox(x,y,w,HEADER_H,[(title,1050,True,WHITE)],anchor="ctr",algn="ctr",name="card-title")]
    out.append(tbox(x+EMU(0.10),y+HEADER_H+EMU(0.05),w-EMU(0.20),h-HEADER_H-EMU(0.10),
                    body_lines,anchor="t",algn="l",lnspc=124000,spcaft=100,name="card-body"))
    return out

# ---------- build the 3-stage rollup diagram ----------
CY=EMU(2.12); CH=EMU(4.05)
AX=EMU(5.12); AW=EMU(1.98)
B1X=EMU(7.12); ARW=EMU(0.34)
BX=EMU(7.48); BW=EMU(2.46)
B2X=EMU(9.96)
CX=EMU(10.32); CW=EMU(2.38)
shapes=[]
# Stage ① raw values
shapes+=card(AX,CY,AW,CH,NAVY,LIGHT,LINE,"① 원시 table_source",
    [("999개 자유 기관명",1000,True,NAVY),
     ("· 서울특별시 금천구",950,False,BODYTX),
     ("· 한국사학진흥재단",950,False,BODYTX),
     ("· 경상남도",950,False,BODYTX),
     ("· ‘소스’(빈값)",950,False,BODYTX)])
# arrow 1
shapes.append(tbox(B1X,CY,ARW,CH,[("▶",2200,True,GREEN)],anchor="ctr",algn="ctr",name="arrow1"))
# Stage ② regex rules (priority order = 06 평가 순서)
shapes+=card(BX,CY,BW,CH,GREEN,LGREEN,LGLINE,"② 정규식 규칙",
    [("기관명 접미사·키워드 매칭",850,False,DGRAY),
     ("· 교육청·대학교·학교 → 교육",950,False,BODYTX),
     ("· 부·청·위원회 → 중앙행정",950,False,BODYTX),
     ("· 공단·재단·병원 → 공공기관",950,False,BODYTX),
     ("· 특별시·도 → 광역지자체",950,False,BODYTX),
     ("· 시·군·구 → 기초지자체",950,False,BODYTX),
     ("· ‘소스’·빈값 → 미상",950,True,GREEN),
     ("· 그 외 → 기타",950,False,BODYTX)])
# arrow 2
shapes.append(tbox(B2X,CY,ARW,CH,[("▶",2200,True,GREEN)],anchor="ctr",algn="ctr",name="arrow2"))
# Stage ③ 7 types + counts
clines=[("7개 기관유형 (테이블 수)",900,True,NAVY)]
for t in TYPE_ORDER:
    disp = "미상(소스)" if t=="미상(소스)" else t
    clines.append((f"· {disp}  {C[t]}", 980, t=="미상(소스)", GREEN if t=="미상(소스)" else BODYTX))
shapes+=card(CX,CY,CW,CH,NAVY,LIGHT,LINE,"③ 7개 기관유형",clines)
# diagram caption
shapes.append(tbox(EMU(5.12),EMU(6.32),EMU(7.58),EMU(0.30),
    [("AOI 롤업: table_source(999) → 접미사·키워드 정규식(우선순위 평가) → 7개 기관유형 · scripts/06",900,False,DGRAY)],
    anchor="t",algn="l",name="diag-caption"))

# ---------- new text ----------
CORE_OLD="999개 자유 기관명과 컬럼명을 상위 개념으로 일반화하여 패턴 가시화"
CORE_NEW="제각각인 999개 기관명을 정규식 규칙으로 7개 상위 개념(기관유형)으로 롤업(상향 일반화)"
BULLETS_NEW=[
 "• AOI(속성지향 귀납) 롤업 = 구체적 값을 상위 개념으로 끌어올리는 다계층 일반화(roll-up)",
 "– 제각각인 999개 기관명은 그대로면 빈발패턴이 안 잡힘 → 개념으로 묶어 통계적 의미 확보",
 "• 출처 롤업: table_source의 접미사·키워드를 정규식으로 매칭해 7개 기관유형으로 축약(우측 도표)",
 "– 우선순위로 1건씩 분류(교육→중앙→공공→광역→기초→미상→기타), ‘소스’·빈값은 미상으로 분리",
 "• 컬럼명 롤업: 괄호·공백·숫자접두 제거(결정적 정규화); 동의어(연번≈순번)는 미병합 — BI 오염 방지(G3)",
]

def edit_core(xml):
    assert f"<a:t>{CORE_OLD}</a:t>" in xml, "core old not found"
    return xml.replace(f"<a:t>{CORE_OLD}</a:t>", f"<a:t>{CORE_NEW}</a:t>")

def edit_bullets(xml):
    # locate the bullets <p:sp> (contains '출처 일반화' originally), narrow it, rewrite each <a:p>
    def sp_repl(m):
        block=m.group(0)
        # bullets box identified by its grid offset (text is split across runs, so don't match on text)
        if '<a:off x="566928" y="1828800"/>' not in block: return block
        block=block.replace('<a:ext cx="6035040" cy="4526280"/>','<a:ext cx="4069080" cy="4526280"/>')
        pc=[0]
        def para_repl(pm):
            i=pc[0]; pc[0]+=1
            if i>=len(BULLETS_NEW): return pm.group(0)
            new=BULLETS_NEW[i]; tc=[0]
            def t_repl(tm):
                tc[0]+=1
                return f"<a:t>{new}</a:t>" if tc[0]==1 else "<a:t></a:t>"
            return re.sub(r"<a:t>.*?</a:t>", t_repl, pm.group(0), flags=re.DOTALL)
        nb=re.sub(r"<a:p>.*?</a:p>", para_repl, block, flags=re.DOTALL)
        assert pc[0]==5, f"expected 5 paragraphs, got {pc[0]}"
        return nb
    new=re.sub(r"<p:sp>.*?</p:sp>", sp_repl, xml, flags=re.DOTALL)
    return new

def main():
    with zipfile.ZipFile(SRC) as z:
        data={i.filename:z.read(i.filename) for i in z.infolist()}; order=z.infolist()
    xml=data[SLIDE].decode("utf-8")
    n_pic=xml.count("<p:pic>")
    xml=re.sub(r"<p:pic>.*?</p:pic>","",xml,flags=re.DOTALL)   # remove fig_aoi
    xml=edit_core(xml)
    xml=edit_bullets(xml)
    assert CORE_NEW in xml and BULLETS_NEW[0] in xml, "text edit failed"
    assert xml.count("</p:spTree>")==1
    xml=xml.replace("</p:spTree>", "".join(shapes)+"</p:spTree>")
    data[SLIDE]=xml.encode("utf-8")
    if os.path.exists(DST): os.remove(DST)
    with zipfile.ZipFile(DST,"w",zipfile.ZIP_DEFLATED) as zout:
        for i in order: zout.writestr(i, data[i.filename])
    print("AOI counts:",[(t,C[t]) for t in TYPE_ORDER],"sum",sum(C.values()))
    print("pic removed:",n_pic,"| shapes added:",len(shapes))
    print("WROTE",DST)

if __name__=="__main__":
    main()
