import sys,re,json
sys.path.insert(0,'/home/claude/d5/tools')
from split import split,text,SUB
from docx.oxml.ns import qn
def bodies(f):
    d,ch=split(f); res={}
    for c in ch:
        cur=c['key']; res[cur]=[]
        for e in c['els'][1:] if c['key']!='pre' else c['els']:
            t=text(e).strip()
            if e.tag==qn('w:p') and SUB.match(t): cur=c['key']+'.'+t.split()[0]; res[cur]=[]
            res[cur].append(('T|' if e.tag==qn('w:tbl') else '')+t)
    return res
if __name__=='__main__':
    a,b=bodies(sys.argv[1]),bodies(sys.argv[2])
    import difflib
    for k in b:
        if '.' in k: continue
        x=a.get(k,[]); y=b[k]
        if x!=y:
            print('=====',k)
            for l in difflib.unified_diff(x,y,lineterm='',n=0):
                if l.startswith(('---','+++','@@')): continue
                print(l[:600])
