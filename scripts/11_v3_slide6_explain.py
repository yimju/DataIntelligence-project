# -*- coding: utf-8 -*-
"""V3 deck: make slide 6 (1D 이산화) friendlier.
- '멱법칙' 같은 어려운 용어를 쉬운 말로 바꾸고(핵심메시지·본문 bullet),
- 좌측 하단에 '멱법칙(쏠림 분포)이란?' 설명 콜아웃 박스를 신규 추가.

Base = V2(헤드폰트 18pt). python-pptx 미설치/오프라인이라 ZIP/XML을 직접 편집(stdlib ElementTree,
네임스페이스 a/r/p 3종만 사용 → register_namespace로 prefix 보존). slide6.xml만 교체, 나머지 불변.
"""
import zipfile, os, copy
import xml.etree.ElementTree as ET

SRC = r"d:\DataInt_2\report\데이터지도_분석보고서_V2.pptx"
DST = r"d:\DataInt_2\report\데이터지도_분석보고서_V3.pptx"
SLIDE = "ppt/slides/slide6.xml"

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
ns = {"a": A, "r": R, "p": P}
for k, v in ns.items():
    ET.register_namespace(k, v)
def qn(t):  # 'a:off' -> '{uri}off'
    pfx, loc = t.split(":"); return f"{{{ns[pfx]}}}{loc}"

CORE_Y = str(int(1.28 * 914400))   # core message box top

# ---- new copy (plain language) -------------------------------------------
CORE_NEW = "한쪽으로 심하게 쏠린 수치(멱법칙)를 로그로 고르게 편 뒤 자동으로 구간을 나눔"
BULLETS_NEW = [
    "• 문제: 레코드수가 1~833,466로 극단 쏠림 → 그대로면 큰 값이 결과를 지배",
    "• 해결: log(로그) 변환으로 쏠림을 눌러 편 뒤 1차원 KMeans로 자동 구간화",
    "• 설명가능: 각 구간 = 원래값 [최소~최대] 범위로 환원(우측 초록선=경계)",
    "• 레코드수: 극소[1-22]·소[23-130]·중[133-829]·대[843-8019]·극대[8578~]",
    "• 컬럼폭: 협소[1-5]·표준[6-11]·광폭[12-44]·초광폭[49-338]",
]
CALLOUT_TITLE = "▶ 쉬운 설명 — ‘멱법칙(쏠림 분포)’이란?"
CALLOUT_BODY = ("소수의 값만 아주 크고 대부분은 작은, 한쪽으로 치우친 분포예요. "
                "예) 소수 데이터셋만 수십만 건인데 대다수는 수십~수백 건. "
                "이런 값은 log(로그)를 씌우면 큰 값이 눌려 고르게 펴지므로, "
                "공정하게 비교하고 구간으로 나눌 수 있습니다.")

def edit_core(sp):
    """Collapse the 11 word-split runs into one run (keep first run's rPr), set new text."""
    p = sp.find(qn("p:txBody")).find(qn("a:p"))
    runs = p.findall(qn("a:r"))
    rpr = copy.deepcopy(runs[0].find(qn("a:rPr")))   # preserves 18pt/bold/color/font
    pPr = p.find(qn("a:pPr"))
    for r in runs:
        p.remove(r)
    new_r = ET.Element(qn("a:r"))
    if rpr is not None:
        new_r.append(rpr)
    t = ET.SubElement(new_r, qn("a:t")); t.text = CORE_NEW
    idx = (list(p).index(pPr) + 1) if pPr is not None else 0
    p.insert(idx, new_r)

def edit_bullets(sp):
    paras = sp.find(qn("p:txBody")).findall(qn("a:p"))
    assert len(paras) == len(BULLETS_NEW), f"expected {len(BULLETS_NEW)} bullets, got {len(paras)}"
    for para, newtxt in zip(paras, BULLETS_NEW):
        t = para.find(qn("a:r")).find(qn("a:t"))
        t.text = newtxt

CALLOUT_XML = f"""<p:sp xmlns:a="{A}" xmlns:p="{P}">
<p:nvSpPr><p:cNvPr id="30" name="V3 Callout 멱법칙 설명"/><p:cNvSpPr txBox="0"/><p:nvPr/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="566928" y="4526280"/><a:ext cx="4846320" cy="1463040"/></a:xfrm>
<a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>
<a:solidFill><a:srgbClr val="E8F5F0"/></a:solidFill>
<a:ln w="9525"><a:solidFill><a:srgbClr val="C7DED3"/></a:solidFill></a:ln>
<a:effectLst/></p:spPr>
<p:txBody>
<a:bodyPr wrap="square" lIns="91440" tIns="64008" rIns="91440" bIns="64008" anchor="t"><a:normAutofit/></a:bodyPr><a:lstStyle/>
<a:p><a:pPr><a:lnSpc><a:spcPct val="106000"/></a:lnSpc><a:spcAft><a:spcPts val="300"/></a:spcAft></a:pPr>
<a:r><a:rPr lang="ko-KR" altLang="en-US" sz="1250" b="1" dirty="0"><a:solidFill><a:srgbClr val="00965F"/></a:solidFill><a:latin typeface="맑은 고딕"/><a:ea typeface="맑은 고딕"/></a:rPr><a:t>{CALLOUT_TITLE}</a:t></a:r></a:p>
<a:p><a:pPr><a:lnSpc><a:spcPct val="106000"/></a:lnSpc></a:pPr>
<a:r><a:rPr lang="ko-KR" altLang="en-US" sz="1150" dirty="0"><a:solidFill><a:srgbClr val="282D37"/></a:solidFill><a:latin typeface="맑은 고딕"/><a:ea typeface="맑은 고딕"/></a:rPr><a:t>{CALLOUT_BODY}</a:t></a:r></a:p>
</p:txBody></p:sp>"""

def main():
    with zipfile.ZipFile(SRC) as z:
        data = {i.filename: z.read(i.filename) for i in z.infolist()}
        order = z.infolist()
    root = ET.fromstring(data[SLIDE])
    spTree = root.find(qn("p:cSld")).find(qn("p:spTree"))
    did_core = did_bul = False
    for sp in spTree.findall(qn("p:sp")):
        off = sp.find(f'.//{qn("a:off")}')
        y = off.get("y") if off is not None else None
        txt = "".join(t.text or "" for t in sp.iter(qn("a:t")))
        if y == CORE_Y:
            edit_core(sp); did_core = True
        elif txt.startswith("• 문제:") or "rec_cnt" in txt or "레코드수가" in txt:
            edit_bullets(sp); did_bul = True
    assert did_core and did_bul, f"core={did_core} bullets={did_bul}"
    # append explanation callout
    spTree.append(ET.fromstring(CALLOUT_XML))
    body = ET.tostring(root, encoding="unicode")
    decl = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    data[SLIDE] = (decl + body).encode("utf-8")

    if os.path.exists(DST):
        os.remove(DST)
    with zipfile.ZipFile(DST, "w", zipfile.ZIP_DEFLATED) as zout:
        for i in order:
            zout.writestr(i, data[i.filename])
    print("core message reworded :", did_core)
    print("bullets reworded      :", did_bul)
    print("callout added         : True")
    print("WROTE", DST)

if __name__ == "__main__":
    main()
