import re, sys, zipfile, shutil, os
INM = {16:18, 6:5, 9:14, 13:15, 18:19, 15:16, 20:17, 11:9, 4:4, 8:6, 17:12, 19:13}
def mapnums(s): return re.sub(r'\d+', lambda m: str(INM[int(m.group())]), s)
def inm(t):
    t = re.sub(r'(D5,? Caps?\. )(\d+(?:(?:, | y )\d+)*)', lambda m: m.group(1)+mapnums(m.group(2)), t)
    t = re.sub(r'Capítulo (\d+) del Documento D5', lambda m: 'Capítulo %s del Documento D5' % INM[int(m.group(1))], t)
    return t
FIJOS = [('(ver D5-COR, capítulo 1)', '(Transitorios Segundo a Noveno del Acuerdo 115/2026)'),
         ('(ver D5-PROF, capítulo 21)', '(Transitorios Segundo a Noveno del Acuerdo 115/2026)'),
         ('(D5-PROF, capítulos 2 y 20)', '(D5-PROF, capítulos 2 y 21)')]
def proc(path, fn):
    z = zipfile.ZipFile(path); tmp = path + '.tmp'
    out = zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED); n = 0
    for it in z.infolist():
        data = z.read(it.filename)
        if it.filename == 'word/document.xml':
            s = data.decode('utf8')
            def f(m):
                nonlocal n
                new = fn(m.group(2))
                if new != m.group(2): n += 1; print('   ', m.group(2)[-90:], '=>', new[-60:])
                return m.group(1) + new + m.group(3)
            s = re.sub(r'(<w:t(?: [^>]*)?>)([^<]*)(</w:t>)', f, s)
            data = s.encode('utf8')
        out.writestr(it, data)
    out.close(); z.close(); shutil.move(tmp, path); return n
def fijos(t):
    for a, b in FIJOS: t = t.replace(a, b)
    return t
if __name__ == "__main__":
  for p in sys.argv[1:]:
      fn = inm if '/inmobiliaria/' in p else fijos
      print(p, proc(p, fn))
