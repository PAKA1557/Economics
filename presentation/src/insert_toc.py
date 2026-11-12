import sys, re
un = sys.argv[1]
E = lambda v: int(round(v*914400))
INK, GR, OR, BG, WH, LG = '1B1E24','5E6672','EE4E22','EFF0F2','FFFFFF','C3C8D0'
S=[('01','사태 개요',['배경','ELS란','손실 규모']),
('02','금융소비자의 법적권리',['권리 강화','분쟁 시 권익 확보','처벌 강화']),
('03','실제 구제와 제재',['실제 구제','대표사례','유사 사례','제재']),
('04','문제점',['해외 사례','주요 문제점']),
('05','개선방안과 결론',['주체별 개선방안','결론'])]
nid=[1]
def rpr(sz,color,b=False):
    return (f'<a:rPr lang="ko-KR" altLang="en-US" sz="{sz}"{" b=\"1\"" if b else ""} dirty="0"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
            '<a:latin typeface="Pretendard" pitchFamily="34" charset="0"/><a:ea typeface="Pretendard" pitchFamily="34" charset="-122"/><a:cs typeface="Pretendard" pitchFamily="34" charset="-120"/></a:rPr>')
def box(name,x,y,w,h,paras,anchor='ctr'):
    nid[0]+=1
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{nid[0]}" name="{name}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="{E(x)}" y="{E(y)}"/><a:ext cx="{E(w)}" cy="{E(h)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln/></p:spPr>'
            f'<p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0" anchor="{anchor}"/><a:lstStyle/>{paras}</p:txBody></p:sp>')
def para(t,sz,color,b=False):
    return f'<a:p><a:pPr marL="0" indent="0" algn="l"><a:buNone/></a:pPr><a:r>{rpr(sz,color,b)}<a:t>{t}</a:t></a:r></a:p>'
def bullets(items):
    return ''.join(f'<a:p><a:pPr marL="216000" indent="-216000" algn="l"><a:lnSpc><a:spcPct val="115000"/></a:lnSpc><a:buFontTx/><a:buChar char="&#8226;"/></a:pPr><a:r>{rpr(1400,GR)}<a:t>{t}</a:t></a:r></a:p>' for t in items)
nid[0]+=1
sh=[f'<p:sp><p:nvSpPr><p:cNvPr id="{nid[0]}" name="목차 패널"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{E(4.3)}" cy="6858000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="{INK}"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="ko-KR" altLang="en-US"/></a:p></p:txBody></p:sp>']
sh.append(box('목차 제목',0.889,0.75,3.3,0.75,para('Contents',3600,WH,True)))
sh.append(box('목차 조명',0.889,6.45,3.0,0.4,para('소비자의 권리 · 1조',1100,LG)))
ROWY=[1.0,2.85,4.7]  # 01·04 and 02·05 share a row
for col,x in (((0,1,2),5.2),((3,4),9.2)):
    for k,i in enumerate(col):
        y=ROWY[k]
        n,name,subs=S[i]
        sh.append(box(f'목차 {n} 번호',x,y,0.55,0.4,para(n,1900,OR,True)))
        sh.append(box(f'목차 {n} 항목',x+0.55,y,3.4,0.4,para(name,1900,INK,True)))
        sh.append(box(f'목차 {n} 하위',x+0.55,y+0.5,3.2,len(subs)*0.3,bullets(subs),'t'))
sp=f'{un}/ppt/slides/slide25.xml'
d=open(sp,encoding='utf8').read()
head=d[:d.index('<p:spTree>')]
grp='<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
tail=d[d.index('</p:spTree>'):]
tail=re.sub(r'<p:extLst>.*?</p:extLst>','',tail,flags=re.S)  # drop cloned creationId
head=head.replace('name="Slide 12"','name="Contents"')
open(sp,'w',encoding='utf8').write(head+'<p:spTree>'+grp+''.join(sh)+tail)
# notes
nd=open(f'{un}/ppt/notesSlides/notesSlide14.xml',encoding='utf8').read()
a=nd.index('<a:lstStyle/>',nd.index('Notes Placeholder'))+len('<a:lstStyle/>'); b=nd.index('</p:txBody>',a)
note='[안내] 발표 순서: 사태 개요 → 금융소비자의 법적권리 → 실제 구제와 제재 → 문제점 → 개선방안과 결론. 항목을 하나씩 읽지 않고 한 문장으로 넘김(10초 이내).'
nd=nd[:a]+f'<a:p><a:r><a:t>{note}</a:t></a:r></a:p>'+nd[b:]
nd=nd.replace('<a:t>14</a:t>','<a:t>2</a:t>')
nd=re.sub(r'<p:extLst>.*?</p:extLst>','',nd,flags=re.S)
open(f'{un}/ppt/notesSlides/notesSlide25.xml','w',encoding='utf8').write(nd)
r=open(f'{un}/ppt/notesSlides/_rels/notesSlide14.xml.rels',encoding='utf8').read().replace('slide17.xml','slide25.xml')
open(f'{un}/ppt/notesSlides/_rels/notesSlide25.xml.rels','w',encoding='utf8').write(r)
rp=f'{un}/ppt/slides/_rels/slide25.xml.rels'
r=open(rp,encoding='utf8').read().replace('</Relationships>','<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide25.xml"/></Relationships>')
open(rp,'w',encoding='utf8').write(r)
cp=f'{un}/[Content_Types].xml'
c=open(cp,encoding='utf8').read()
c=c.replace('</Types>','<Override PartName="/ppt/notesSlides/notesSlide25.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/></Types>')
open(cp,'w',encoding='utf8').write(c)
