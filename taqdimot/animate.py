"""Add slide transitions and entrance animations (Float In, auto-play cascade) to a pptx.

Usage: python3 animate.py deck.pptx
Every shape on every slide floats in top-to-bottom, left-to-right, automatically
after the slide opens (no clicks needed). Each slide gets a transition.
"""
import re
import sys
import zipfile
import shutil

EMU = 914400
TRANSITIONS = ['<p:fade/>', '<p:push dir="u"/>', '<p:wipe dir="r"/>', '<p:cover dir="l"/>']


def shapes(xml):
    """Return [(id, kind, y, x)] for top-level shapes in the slide's spTree."""
    out = []
    for m in re.finditer(r'<p:(sp|pic|graphicFrame)>(.*?)</p:\1>', xml, re.S):
        kind, body = m.group(1), m.group(2)
        sid = re.search(r'<p:cNvPr id="(\d+)"', body)
        off = re.search(r'<a:off x="(-?\d+)" y="(-?\d+)"', body)
        if not sid or not off:
            continue
        out.append((int(sid.group(1)), kind, int(off.group(2)), int(off.group(1)), 'type="title"' in body))
    return out


def effect(ctn_id, spid, delay, first):
    node = "afterEffect" if first else "withEffect"
    tgt = f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>'
    ids = iter(range(ctn_id + 1, ctn_id + 10))
    return ctn_id + 5, (
        f'<p:par><p:cTn id="{ctn_id}" fill="hold"><p:stCondLst><p:cond delay="{delay}"/></p:stCondLst><p:childTnLst>'
        f'<p:par><p:cTn id="{next(ids)}" presetID="42" presetClass="entr" presetSubtype="0" fill="hold" grpId="0" nodeType="{node}">'
        f'<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
        f'<p:set><p:cBhvr><p:cTn id="{next(ids)}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>{tgt}'
        f'<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>'
        f'<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="{next(ids)}" dur="500"/>{tgt}</p:cBhvr></p:animEffect>'
        f'<p:anim calcmode="lin" valueType="num"><p:cBhvr><p:cTn id="{next(ids)}" dur="500" decel="100000" fill="hold"/>{tgt}'
        f'<p:attrNameLst><p:attrName>ppt_y</p:attrName></p:attrNameLst></p:cBhvr><p:tavLst>'
        f'<p:tav tm="0"><p:val><p:strVal val="#ppt_y+.1"/></p:val></p:tav>'
        f'<p:tav tm="100000"><p:val><p:strVal val="#ppt_y"/></p:val></p:tav></p:tavLst></p:anim>'
        f'</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>'
    )


def timing(items):
    # items sorted; title first, then rows top->bottom
    parts, cid = [], 5
    for k, (spid, kind, *_rest) in enumerate(items):
        cid, xml = effect(cid, spid, k * 90, k == 0)
        parts.append(xml)
    # Each effect is wrapped in its own par with its delay; all start after slide begins
    inner = "".join(parts)
    bld = "".join(
        f'<p:bldGraphic spid="{spid}" grpId="0"><p:bldAsOne/></p:bldGraphic>' if kind == "graphicFrame"
        else (f'<p:bldP spid="{spid}" grpId="0" animBg="1"/>' if kind == "sp" else "")
        for spid, kind, *_ in items
    )
    bld = bld.replace("", "")
    return (
        '<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>'
        '<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>'
        '<p:par><p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="indefinite"/>'
        '<p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond></p:stCondLst><p:childTnLst>'
        f'<p:par><p:cTn id="4" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>{inner}'
        '</p:childTnLst></p:cTn></p:par>'
        '</p:childTnLst></p:cTn></p:par>'
        '</p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
        '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>'
        f'</p:childTnLst></p:cTn></p:par></p:tnLst>{"<p:bldLst>" + bld + "</p:bldLst>" if bld else ""}</p:timing>'
    )


def main(path):
    tmp = path + ".tmp"
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            m = re.fullmatch(r"ppt/slides/slide(\d+)\.xml", info.filename)
            if m:
                n = int(m.group(1))
                xml = data.decode("utf-8")
                xml = re.sub(r"<p:transition.*?</p:transition>|<p:transition[^>]*/>", "", xml, flags=re.S)
                xml = re.sub(r"<p:timing>.*?</p:timing>", "", xml, flags=re.S)
                items = shapes(xml)
                # title first, then by row (0.15" bands) and x
                items.sort(key=lambda t: (not t[4], round(t[2] / (EMU * 0.15)), t[3]))
                trans = f'<p:transition spd="med">{TRANSITIONS[(n - 1) % len(TRANSITIONS)]}</p:transition>'
                add = trans + (timing(items) if items else "")
                # transition + timing go right after clrMapOvr (or cSld)
                if "</p:clrMapOvr>" in xml:
                    xml = xml.replace("</p:clrMapOvr>", "</p:clrMapOvr>" + add, 1)
                else:
                    xml = xml.replace("</p:cSld>", "</p:cSld>" + add, 1)
                data = xml.encode("utf-8")
            zout.writestr(info, data)
    shutil.move(tmp, path)
    print("animated", path)


if __name__ == "__main__":
    main(sys.argv[1])
