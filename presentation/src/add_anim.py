"""pptx 슬라이드에 클릭 단계별 등장 애니메이션을 넣는다. 도형 이름이 'aN ...' 이면 N번째 클릭, 'a0 ...' 이면 진입 시 자동.
이름에 '|wipe' / '|wheel' 이 있으면 그 효과, 없으면 서서히 나타나기."""
import re, sys, zipfile, shutil, os
src, dst = sys.argv[1], sys.argv[2]
EFF = {'fade': ('10', '0', 'fade', 450), 'wipe': ('22', '8', 'wipe(left)', 750), 'wheel': ('21', '1', 'wheel(1)', 1300)}
def timing(xml):
    items = []
    for m in re.finditer(r'<p:cNvPr id="(\d+)" name="a(\d+)(?:\|(\w+))? [^"]*"', xml):
        before = xml[:m.start()]
        kind = 'pic' if before.rfind('<p:pic>') > before.rfind('<p:sp>') else 'sp'
        items.append((int(m.group(2)), m.group(1), m.group(3) or 'fade', kind))
    if not items: return None
    cid = [2]
    def nid():
        cid[0] += 1; return cid[0]
    groups = {}
    for step, spid, eff, kind in items: groups.setdefault(step, []).append((spid, eff, kind))
    pars = []
    for step in sorted(groups):
        effs = []
        for i, (spid, eff, kind) in enumerate(groups[step]):
            pid, sub, flt, dur = EFF[eff]
            node = ('afterEffect' if step == 0 else 'clickEffect') if i == 0 else 'withEffect'
            effs.append(f'<p:par><p:cTn id="{nid()}" presetID="{pid}" presetClass="entr" presetSubtype="{sub}" fill="hold" grpId="0" nodeType="{node}"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
                        f'<p:set><p:cBhvr><p:cTn id="{nid()}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>'
                        f'<p:animEffect transition="in" filter="{flt}"><p:cBhvr><p:cTn id="{nid()}" dur="{dur}"/><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl></p:cBhvr></p:animEffect>'
                        f'</p:childTnLst></p:cTn></p:par>')
        cond = '<p:cond delay="indefinite"/>' + ('<p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond>' if step == 0 else '')
        pars.append(f'<p:par><p:cTn id="{nid()}" fill="hold"><p:stCondLst>{cond}</p:stCondLst><p:childTnLst><p:par><p:cTn id="{nid()}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>{"".join(effs)}</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>')
    bld = ''.join(f'<p:bldP spid="{spid}" grpId="0" animBg="1"/>' for _, spid, _, kind in items if kind == 'sp')
    return ('<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>'
            '<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>' + ''.join(pars) +
            '</p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
            '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>'
            '</p:childTnLst></p:cTn></p:par></p:tnLst><p:bldLst>' + bld + '</p:bldLst></p:timing>')
zin = zipfile.ZipFile(src); zout = zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED); cnt = 0
for it in zin.infolist():
    if it.filename.endswith('/'): continue
    data = zin.read(it.filename)
    if re.fullmatch(r'ppt/slides/slide\d+\.xml', it.filename):
        x = data.decode('utf8'); t = timing(x)
        if t:
            x = x.replace('</p:sld>', t + '</p:sld>'); cnt += 1
        # 도형 이름에서 애니메이션 표시 제거
        x = re.sub(r'name="a\d+(?:\|\w+)? ', 'name="', x)
        data = x.encode('utf8')
    zout.writestr(it, data)
zout.close(); print('animated slides:', cnt)
