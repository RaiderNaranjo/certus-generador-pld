"""Modelo de párrafo y normalización de formato para los D5 homologados.

Un párrafo se convierte en una lista de caracteres con sus propiedades (negrita, cursiva,
subrayado y estilo de variante). Sobre esa cadena se aplican las reglas de texto; después
se reconstruyen los runs con el formato canónico. Así se conservan las variantes [PF: …]
/ [PM: …] y las negritas de entrada, y se elimina el formato directo heterogéneo.
"""
import copy
import re
from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
XML = 'http://www.w3.org/XML/1998/namespace'
def q(tag):
    p, t = tag.split(':')
    return '{%s}%s' % ({'w': W}[p], t)

VARIANTES = ('VariantePF', 'VariantePM', 'NotaCertus')
LOG = []

def log(msg):
    LOG.append(msg)

# ---------------------------------------------------------------- estilos canónicos
ESTILOS_XML = '''
<w:styles xmlns:w="%s">
<w:style w:type="paragraph" w:customStyle="1" w:styleId="CertusTexto">
  <w:name w:val="Certus Texto"/><w:basedOn w:val="Normal"/><w:qFormat/>
  <w:pPr><w:spacing w:before="0" w:after="110" w:line="259" w:lineRule="auto"/><w:jc w:val="both"/></w:pPr>
  <w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr>
</w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="CertusCapitulo">
  <w:name w:val="Certus Capítulo"/><w:basedOn w:val="CertusTexto"/><w:next w:val="CertusTexto"/><w:qFormat/>
  <w:pPr><w:keepNext/><w:keepLines/><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="2" w:color="1F3864"/></w:pBdr>
    <w:spacing w:before="280" w:after="120"/><w:jc w:val="left"/><w:outlineLvl w:val="0"/></w:pPr>
  <w:rPr><w:b/><w:bCs/><w:color w:val="1F3864"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>
</w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="CertusSeccion">
  <w:name w:val="Certus Sección"/><w:basedOn w:val="CertusTexto"/><w:next w:val="CertusTexto"/><w:qFormat/>
  <w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="160" w:after="80"/><w:jc w:val="left"/><w:outlineLvl w:val="1"/></w:pPr>
  <w:rPr><w:b/><w:bCs/><w:color w:val="1F3864"/></w:rPr>
</w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="CertusVineta">
  <w:name w:val="Certus Viñeta"/><w:basedOn w:val="CertusTexto"/><w:qFormat/>
  <w:pPr><w:tabs><w:tab w:val="left" w:pos="340"/></w:tabs><w:spacing w:after="70"/><w:ind w:left="340" w:hanging="340"/></w:pPr>
</w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="CertusIndice">
  <w:name w:val="Certus Índice"/><w:basedOn w:val="CertusTexto"/>
  <w:pPr><w:spacing w:before="60" w:after="160"/><w:jc w:val="left"/></w:pPr>
  <w:rPr><w:i/><w:iCs/><w:color w:val="404040"/><w:sz w:val="19"/><w:szCs w:val="19"/></w:rPr>
</w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="CertusTabla">
  <w:name w:val="Certus Tabla"/><w:basedOn w:val="CertusTexto"/>
  <w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/><w:jc w:val="left"/></w:pPr>
  <w:rPr><w:sz w:val="19"/><w:szCs w:val="19"/></w:rPr>
</w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="CertusTablaEncabezado">
  <w:name w:val="Certus Tabla Encabezado"/><w:basedOn w:val="CertusTabla"/>
  <w:pPr><w:keepNext/></w:pPr>
  <w:rPr><w:b/><w:bCs/><w:color w:val="FFFFFF"/></w:rPr>
</w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="CertusPie">
  <w:name w:val="Certus Pie"/><w:basedOn w:val="CertusTexto"/>
  <w:pPr><w:spacing w:before="240" w:after="0"/><w:jc w:val="center"/></w:pPr>
  <w:rPr><w:color w:val="595959"/><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr>
</w:style>
</w:styles>''' % W

def instalar_estilos(doc):
    st = doc.styles.element
    nuevos = etree.fromstring(ESTILOS_XML)
    for s in nuevos:
        sid = s.get(q('w:styleId'))
        for old in st.findall(q('w:style')):
            if old.get(q('w:styleId')) == sid:
                st.remove(old)
        st.append(s)

# ---------------------------------------------------------------- modelo de párrafo
class Par:
    """Texto de un párrafo con propiedades por carácter."""
    def __init__(self, chars=None, props=None, mark=None, bullet=False):
        self.chars = chars or []
        self.props = props or []
        self.mark = mark        # estilo de variante de la marca de párrafo (párrafo completo)
        self.bullet = bullet

    @property
    def text(self):
        return ''.join(self.chars)

    @classmethod
    def from_text(cls, s, base=frozenset()):
        """Texto con marcado mínimo: **negrita**, y [PF: …] / [PM: …] como variantes."""
        p = cls()
        bold = False
        i = 0
        while i < len(s):
            if s.startswith('**', i):
                bold = not bold; i += 2; continue
            m = re.match(r'\[(PF|PM): ', s[i:])
            if m:
                j = s.index(']', i)
                seg = s[i:j + 1]
                st = 'VariantePF' if m.group(1) == 'PF' else 'VariantePM'
                for ch in seg:
                    p.chars.append(ch)
                    p.props.append(frozenset(base | ({'b'} if bold else set()) | {'rs:' + st}))
                i = j + 1
                continue
            p.chars.append(s[i])
            p.props.append(frozenset(base | ({'b'} if bold else set())))
            i += 1
        return p

    def replace(self, pattern, repl, flags=0, protect=True, count=0):
        """Reemplaza en el texto. repl: str | callable(m)->str | list[(texto, set_extra)].
        El texto nuevo hereda las propiedades del primer carácter coincidente. No toca los
        marcadores del generador ([…] que no son variantes) si protect=True."""
        text = self.text
        prot = []
        if protect:
            prot = [(m.start(), m.end()) for m in re.finditer(r'\[(?!PF: |PM: )[^\[\]]*\]', text)]
        out_c, out_p = [], []
        last = 0
        n = 0
        for m in re.finditer(pattern, text, flags):
            if any(a < m.end() and m.start() < b for a, b in prot):
                continue
            if count and n >= count:
                break
            n += 1
            out_c += self.chars[last:m.start()]
            out_p += self.props[last:m.start()]
            base = self.props[m.start()] if m.start() < len(self.props) else (self.props[-1] if self.props else frozenset())
            r = repl(m) if callable(repl) else repl
            if isinstance(r, str):
                r = m.expand(r) if not callable(repl) else r
                out_c += list(r); out_p += [base] * len(r)
            else:
                for seg, extra in r:
                    pr = frozenset([x for x in base if not x.startswith('rs:')] + list(extra))
                    out_c += list(seg); out_p += [pr] * len(seg)
            last = m.end()
        if n:
            out_c += self.chars[last:]; out_p += self.props[last:]
            self.chars, self.props = out_c, out_p
        return n

    def strip(self):
        while self.chars and self.chars[0] in ' \t':
            self.chars.pop(0); self.props.pop(0)
        while self.chars and self.chars[-1] in ' \t\n':
            self.chars.pop(); self.props.pop()

def leer_par(p):
    """lxml w:p -> Par. Devuelve None si el párrafo tiene contenido que no se puede rehacer."""
    par = Par()
    ppr = p.find(q('w:pPr'))
    if ppr is not None:
        rs = ppr.find('%s/%s' % (q('w:rPr'), q('w:rStyle')))
        if rs is not None and rs.get(q('w:val')) in VARIANTES:
            par.mark = rs.get(q('w:val'))
        if ppr.find(q('w:numPr')) is not None:
            par.bullet = True
        ps = ppr.find(q('w:pStyle'))
        if ps is not None and ps.get(q('w:val')) in ('ListParagraph', 'Prrafodelista'):
            par.bullet = True
    complejo = False
    def runs(el):
        for ch in el:
            if ch.tag == q('w:r'):
                yield ch
            elif ch.tag in (q('w:hyperlink'), q('w:ins'), q('w:smartTag'), q('w:fldSimple')):
                yield from runs(ch)
    for r in runs(p):
        rpr = r.find(q('w:rPr'))
        props = set()
        if rpr is not None:
            for tag, key in (('w:b', 'b'), ('w:i', 'i'), ('w:u', 'u')):
                e = rpr.find(q(tag))
                if e is not None and e.get(q('w:val')) not in ('false', '0', 'none'):
                    props.add(key)
            rs = rpr.find(q('w:rStyle'))
            if rs is not None and rs.get(q('w:val')) in VARIANTES:
                props.add('rs:' + rs.get(q('w:val')))
        props = frozenset(props)
        for ch in r:
            if ch.tag == q('w:t'):
                for c in (ch.text or ''):
                    par.chars.append(c); par.props.append(props)
            elif ch.tag == q('w:tab'):
                par.chars.append('\t'); par.props.append(props)
            elif ch.tag in (q('w:br'), q('w:cr')):
                par.chars.append('\n'); par.props.append(props)
            elif ch.tag == q('w:noBreakHyphen'):
                par.chars.append('-'); par.props.append(props)
            elif ch.tag in (q('w:drawing'), q('w:pict'), q('w:fldChar'), q('w:instrText'), q('w:object')):
                complejo = True
    if complejo:
        return None
    # viñetas escritas a mano
    t = par.text
    m = re.match(r'^(\[P[FM]: )?\s*[•●▪◦·\-–]\s+', t)
    if m and (t.lstrip().startswith(('•', '●', '▪', '◦')) or (m.group(1) and re.match(r'^\[P[FM]: \s*[•●▪◦]', t))):
        par.bullet = True
        pre = m.group(1) or ''
        del par.chars[len(pre):m.end()]
        del par.props[len(pre):m.end()]
    return par

def escribir_par(par, estilo):
    """Par -> lxml w:p con estilo de párrafo canónico."""
    p = etree.Element(q('w:p'))
    ppr = etree.SubElement(p, q('w:pPr'))
    etree.SubElement(ppr, q('w:pStyle')).set(q('w:val'), estilo)
    if par.mark:
        rpr = etree.SubElement(ppr, q('w:rPr'))
        etree.SubElement(rpr, q('w:rStyle')).set(q('w:val'), par.mark)
    chars, props = list(par.chars), list(par.props)
    if par.bullet and estilo == 'CertusVineta':
        base = frozenset()
        chars = ['•', '\t'] + chars
        props = [base, base] + props
    i = 0
    while i < len(chars):
        j = i
        while j < len(chars) and props[j] == props[i]:
            j += 1
        r = etree.SubElement(p, q('w:r'))
        pr = props[i]
        if pr:
            rpr = etree.SubElement(r, q('w:rPr'))
            for x in sorted(pr):
                if x.startswith('rs:'):
                    etree.SubElement(rpr, q('w:rStyle')).set(q('w:val'), x[3:])
            if 'b' in pr:
                etree.SubElement(rpr, q('w:b')); etree.SubElement(rpr, q('w:bCs'))
            if 'i' in pr:
                etree.SubElement(rpr, q('w:i')); etree.SubElement(rpr, q('w:iCs'))
            if 'u' in pr:
                etree.SubElement(rpr, q('w:u')).set(q('w:val'), 'single')
            # el estilo de carácter debe ir primero en rPr
            rs = rpr.find(q('w:rStyle'))
            if rs is not None:
                rpr.remove(rs); rpr.insert(0, rs)
        buf = ''
        def flush():
            nonlocal buf
            if buf:
                t = etree.SubElement(r, q('w:t')); t.text = buf
                t.set('{%s}space' % XML, 'preserve')
                buf = ''
        for c in chars[i:j]:
            if c == '\t':
                flush(); etree.SubElement(r, q('w:tab'))
            elif c == '\n':
                flush(); etree.SubElement(r, q('w:br'))
            else:
                buf += c
        flush()
        i = j
    return p

def normalizar_parrafo_complejo(p, estilo):
    """Para párrafos que no se rehacen: solo se sustituye pPr y se quita formato directo."""
    ppr = p.find(q('w:pPr'))
    mark = None
    if ppr is not None:
        rs = ppr.find('%s/%s' % (q('w:rPr'), q('w:rStyle')))
        mark = rs.get(q('w:val')) if rs is not None and rs.get(q('w:val')) in VARIANTES else None
        p.remove(ppr)
    ppr = etree.Element(q('w:pPr'))
    etree.SubElement(ppr, q('w:pStyle')).set(q('w:val'), estilo)
    if mark:
        etree.SubElement(etree.SubElement(ppr, q('w:rPr')), q('w:rStyle')).set(q('w:val'), mark)
    p.insert(0, ppr)
    for rpr in p.iter(q('w:rPr')):
        if rpr.getparent() is ppr:
            continue
        for tag in ('w:rFonts', 'w:sz', 'w:szCs', 'w:color', 'w:shd', 'w:highlight'):
            for e in rpr.findall(q(tag)):
                rpr.remove(e)
    return p

# ---------------------------------------------------------------- tablas
ANCHO = 9500

def normalizar_tabla(tbl, proc_par, encabezado=None):
    """Formato canónico de tabla: ancho 9,500, bordes simples, encabezado azul marino.
    proc_par(Par) aplica las reglas de texto a cada párrafo de celda."""
    tblpr = tbl.find(q('w:tblPr'))
    if tblpr is not None:
        tbl.remove(tblpr)
    tblpr = etree.fromstring(
        '<w:tblPr xmlns:w="%s"><w:tblW w:w="%d" w:type="dxa"/><w:jc w:val="center"/>'
        '<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="7F7F7F"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="7F7F7F"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="7F7F7F"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="7F7F7F"/>'
        '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="7F7F7F"/>'
        '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="7F7F7F"/></w:tblBorders>'
        '<w:tblLayout w:type="fixed"/>'
        '<w:tblCellMar><w:top w:w="50" w:type="dxa"/><w:left w:w="100" w:type="dxa"/>'
        '<w:bottom w:w="50" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tblCellMar>'
        '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="1" w:noVBand="1"/>'
        '</w:tblPr>' % (W, ANCHO))
    tbl.insert(0, tblpr)
    grid = tbl.find(q('w:tblGrid'))
    cols = [int(g.get(q('w:w'))) for g in grid.findall(q('w:gridCol'))] if grid is not None else []
    total = sum(cols) or ANCHO
    f = ANCHO / total
    if grid is not None:
        acc = 0
        gcs = grid.findall(q('w:gridCol'))
        for k, g in enumerate(gcs):
            w = round(int(g.get(q('w:w'))) * f) if k < len(gcs) - 1 else ANCHO - acc
            g.set(q('w:w'), str(w)); acc += w
    filas = tbl.findall(q('w:tr'))
    if encabezado is None:
        encabezado = es_encabezado(filas[0]) if filas else False
    for ri, tr in enumerate(filas):
        trpr = tr.find(q('w:trPr'))
        if trpr is not None:
            for e in list(trpr):
                if e.tag not in (q('w:tblHeader'), q('w:cantSplit')):
                    trpr.remove(e)
        hdr = encabezado and ri == 0
        if hdr:
            if trpr is None:
                trpr = etree.Element(q('w:trPr')); tr.insert(0, trpr)
            if trpr.find(q('w:tblHeader')) is None:
                etree.SubElement(trpr, q('w:tblHeader'))
        for tc in tr.findall(q('w:tc')):
            tcpr = tc.find(q('w:tcPr'))
            keep = {}
            if tcpr is not None:
                for e in tcpr:
                    if e.tag in (q('w:tcW'), q('w:gridSpan'), q('w:vMerge'), q('w:hMerge')):
                        keep[e.tag] = e
                tc.remove(tcpr)
            tcpr = etree.Element(q('w:tcPr'))
            tcw = keep.get(q('w:tcW'))
            if tcw is not None and tcw.get(q('w:type')) in ('dxa', None) and tcw.get(q('w:w')):
                tcw.set(q('w:w'), str(round(int(tcw.get(q('w:w'))) * f))); tcw.set(q('w:type'), 'dxa')
            for tag in ('w:tcW', 'w:gridSpan', 'w:hMerge', 'w:vMerge'):
                if q(tag) in keep:
                    tcpr.append(keep[q(tag)])
            if hdr:
                shd = etree.SubElement(tcpr, q('w:shd'))
                shd.set(q('w:val'), 'clear'); shd.set(q('w:color'), 'auto'); shd.set(q('w:fill'), '1F3864')
            etree.SubElement(tcpr, q('w:vAlign')).set(q('w:val'), 'center')
            tc.insert(0, tcpr)
            for p in tc.findall(q('w:p')):
                par = leer_par(p)
                est = 'CertusTablaEncabezado' if hdr else 'CertusTabla'
                if par is None:
                    normalizar_parrafo_complejo(p, est); continue
                proc_par(par)
                par.strip()
                if par.bullet:
                    par.chars = ['•', ' '] + par.chars; par.props = [frozenset()] * 2 + par.props
                    par.bullet = False
                new = escribir_par(par, est)
                tc.replace(p, new)
            if tc.find(q('w:p')) is None:
                etree.SubElement(tc, q('w:p'))
    return tbl

def es_encabezado(tr):
    tcs = tr.findall(q('w:tc'))
    if not tcs:
        return False
    sombreadas = 0
    negritas = 0
    for tc in tcs:
        shd = tc.find('%s/%s' % (q('w:tcPr'), q('w:shd')))
        if shd is not None and (shd.get(q('w:fill')) or 'auto').upper() not in ('AUTO', 'FFFFFF'):
            sombreadas += 1
        par = None
        ps = tc.findall(q('w:p'))
        texto = ''.join(''.join(t.text or '' for t in p.iter(q('w:t'))) for p in ps).strip()
        bs = [r for p in ps for r in p.iter(q('w:r')) if ''.join(t.text or '' for t in r.iter(q('w:t'))).strip()]
        if bs and all(r.find('%s/%s' % (q('w:rPr'), q('w:b'))) is not None for r in bs):
            negritas += 1
        elif not texto:
            negritas += 1
    return sombreadas == len(tcs) or (negritas == len(tcs) and len(tcs) > 1)

def tabla_simple(filas, anchos, encabezado=True):
    """Tabla nueva a partir de texto (marcado de Par.from_text en cada celda)."""
    tbl = etree.Element(q('w:tbl'))
    etree.SubElement(tbl, q('w:tblPr'))
    grid = etree.SubElement(tbl, q('w:tblGrid'))
    for w in anchos:
        etree.SubElement(grid, q('w:gridCol')).set(q('w:w'), str(w))
    for fila in filas:
        tr = etree.SubElement(tbl, q('w:tr'))
        for k, celda in enumerate(fila):
            tc = etree.SubElement(tr, q('w:tc'))
            tcpr = etree.SubElement(tc, q('w:tcPr'))
            tcw = etree.SubElement(tcpr, q('w:tcW')); tcw.set(q('w:w'), str(anchos[k])); tcw.set(q('w:type'), 'dxa')
            tc.append(escribir_par(Par.from_text(celda), 'CertusTabla'))
    return tbl
