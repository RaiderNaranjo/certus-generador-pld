import sys,re,json
from docx import Document
from docx.oxml.ns import qn
CH=re.compile(r'^(Capítulo|CAPÍTULO)\s+(\d+)[\.:]?\s*(.*)$')
AN=re.compile(r'^Anexo\s+([A-D])[\.:]\s*(.*)$')
SUB=re.compile(r'^(\d+)\.(\d+)(\.\d+)?\s+\S')
def text(el):
    return ''.join(t.text or '' for t in el.iter(qn('w:t')))
def split(path):
    d=Document(path); body=d.element.body
    out=[]; cur={'key':'pre','title':'','els':[]}
    for el in body:
        if el.tag==qn('w:sectPr'): continue
        t=text(el).strip()
        m=CH.match(t) if el.tag==qn('w:p') else None
        a=AN.match(t) if el.tag==qn('w:p') else None
        if m or a:
            out.append(cur)
            cur={'key':('C'+m.group(2)) if m else ('A'+a.group(1)),'title':t,'els':[]}
        cur['els'].append(el)
    out.append(cur); return d,out
if __name__=='__main__':
    for f in sys.argv[1:]:
        d,ch=split(f)
        print('====',f)
        for c in ch:
            subs=[text(e).strip()[:60] for e in c['els'] if e.tag==qn('w:p') and SUB.match(text(e).strip())]
            words=sum(len(text(e).split()) for e in c['els'])
            print(f"{c['key']:4} {words:5}w  {c['title'][:70]}")
            for s in subs: print('        ',s)
