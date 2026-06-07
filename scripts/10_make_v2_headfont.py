# -*- coding: utf-8 -*-
"""Make a V2 deck from the existing report: enlarge every slide's CORE (head) message
font 15pt -> 18pt. The core-message text box is uniquely identified by its fixed grid
position top=1.28in (EMU y="1170432", set in scripts/08_build_pptx.py base()).

No external deps (python-pptx unavailable / offline): a .pptx is a ZIP of XML, so we
edit the slide parts in place and rewrite the archive. Only the <p:sp> shape sitting at
y=1170432 is touched, so titles/bullets/tables/cards are left exactly as-is.
"""
import zipfile, re, shutil, os

SRC = r"d:\DataInt_2\report\데이터지도_분석보고서.pptx"
DST = r"d:\DataInt_2\report\데이터지도_분석보고서_V2.pptx"
CORE_Y = str(int(1.28 * 914400))   # 1170432  (CORE_TOP grid line)
OLD_SZ, NEW_SZ = 'sz="1500"', 'sz="1800"'   # 15pt -> 18pt (head message)

SP_RE = re.compile(r"<p:sp>.*?</p:sp>", re.DOTALL)
SLIDE_RE = re.compile(r"ppt/slides/slide\d+\.xml$")

def patch_core_sp(block: str):
    """Within a single <p:sp> that is the core message, bump head font to 18pt."""
    if f'y="{CORE_Y}"' not in block:
        return block, 0
    new_block, n = re.subn(re.escape(OLD_SZ), NEW_SZ, block)
    return new_block, n

def patch_slide(xml: str):
    total = 0
    def repl(m):
        nonlocal total
        nb, n = patch_core_sp(m.group(0))
        total += n
        return nb
    return SP_RE.sub(repl, xml), total

def main():
    if not os.path.exists(SRC):
        raise SystemExit(f"source not found: {SRC}")
    changed = {}
    with zipfile.ZipFile(SRC) as zin:
        infos = zin.infolist()
        data = {i.filename: zin.read(i.filename) for i in infos}
        slide_runs = 0
        slide_hits = 0
        for name in data:
            if SLIDE_RE.search(name):
                xml = data[name].decode("utf-8")
                new_xml, n = patch_slide(xml)
                if n:
                    data[name] = new_xml.encode("utf-8")
                    changed[name] = n
                    slide_hits += 1
                    slide_runs += n
    # rewrite archive preserving order/compression
    if os.path.exists(DST):
        os.remove(DST)
    with zipfile.ZipFile(SRC) as zin, zipfile.ZipFile(DST, "w", zipfile.ZIP_DEFLATED) as zout:
        for i in zin.infolist():
            zout.writestr(i, data[i.filename])
    print(f"slides patched : {slide_hits}")
    print(f"runs bumped 15->18pt : {slide_runs}")
    for k in sorted(changed, key=lambda s: int(re.search(r'(\d+)', s).group(1))):
        print(f"  {k}: {changed[k]} run(s)")
    print("WROTE", DST)

if __name__ == "__main__":
    main()
