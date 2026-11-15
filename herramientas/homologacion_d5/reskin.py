"""Aplica la presentación y las reglas de redacción Certus a cualquier plantilla, conservando su contenido.

Uso: python3 reskin.py <contenedor.docx> <entrada.docx> <salida.docx>
- Portada: la primera tabla con «CERTUS PLD» se rehace como banda canónica.
- Títulos: SECCIÓN/Capítulo/romanos/Anexo -> «Certus Capítulo»; A.1, 2.1, G.1… -> «Certus Sección».
- Cuerpo, viñetas y tablas con estilos canónicos; reglas de redacción comunes (build_d5.GLOBAL).
"""
import copy, os, re, sys
sys.path.insert(0, os.path.dirname(__file__))
from lxml import etree
from docx import Document
from docxlib import (q, W, Par, leer_par, escribir_par, normalizar_parrafo_complejo, normalizar_tabla,
                     instalar_estilos, espaciar_tras_tablas, numero_de_pagina, LOG, log)
import build_d5 as B

EXTRA = [
    (re.compile(r'Manual de Cumplimiento(?: PLD/FT)?'), lambda m: 'Manual de Políticas Internas' + (' PLD/FT' if 'PLD/FT' in m.group(0) else '')),
    (re.compile(r'CERTUS CONSULTORES(?= *[│|] *Certus PLD)'), 'CERTUS CONSULTORES · ESPINOSA-SALAZAR ABOGADOS CORPORATIVOS'),
    (re.compile(r'\bRepr\. de Cumplimiento\b'), 'Representante Encargado de Cumplimiento'),
    (re.compile(r'\bRepresentante Cumpl\.'), 'Representante Encargado de Cumplimiento'),
    (re.compile(r'\bPersonas? Expuestas? Pol[ií]ticamente\b'), lambda m: m.group(0).replace('Expuestas Políticamente', 'Políticamente Expuestas').replace('Expuesta Políticamente', 'Políticamente Expuesta')),
]

def reglas(par):
    for pat, repl in EXTRA:
        par.replace(pat, repl)
    B.aplicar_globales(par)

TOP = re.compile(r'^(?:(SECCI[ÓO]N|Secci[óo]n|CAP[ÍI]TULO|Cap[íi]tulo|ANEXO|Anexo|PARTE|Parte)\s+([A-Z0-9IVXL]+(?:\.\d+)?)\s*[:.—–-]?\s*(.*)'
                 r'|([IVXL]+)\.\s+(\S.*))$')
SUBT = re.compile(r'^(?:[A-Z]\.\d+|\d+\.\d+|[IVXL]+\.\d+)(?:\.\d+)?\.?\s+\S')

def es_negrita(par):
    vis = [pr for c, pr in zip(par.chars, par.props) if c.strip()]
    return bool(vis) and all('b' in pr for pr in vis)

def titulo_canonico(texto):
    m = TOP.match(texto.strip())
    if not m:
        return None
    if m.group(4):
        return 'Sección %s: %s' % (m.group(4), m.group(5).strip())
    tipo, num, resto = m.group(1), m.group(2), (m.group(3) or '').strip()
    tipo = {'SECCIÓN': 'Sección', 'SECCION': 'Sección', 'Seccion': 'Sección', 'CAPÍTULO': 'Capítulo',
            'CAPITULO': 'Capítulo', 'Capitulo': 'Capítulo', 'ANEXO': 'Anexo', 'PARTE': 'Parte'}.get(tipo, tipo)
    return '%s %s: %s' % (tipo, num, resto) if resto else '%s %s' % (tipo, num)

def pstyle(p):
    ps = p.find('%s/%s' % (q('w:pPr'), q('w:pStyle')))
    return ps.get(q('w:val')) if ps is not None else ''

def sombreado_oscuro(p):
    shd = p.find('%s/%s' % (q('w:pPr'), q('w:shd')))
    return shd is not None and (shd.get(q('w:fill')) or '').upper() in ('1F3864', '0A1628', '162D52', '002060', '1F4E79')

def jc(p):
    j = p.find('%s/%s' % (q('w:pPr'), q('w:jc')))
    return j.get(q('w:val')) if j is not None else None

def banda(tbl):
    """Banda de portada canónica a partir del texto de la tabla original."""
    textos = [''.join(t.text or '' for t in pp.iter(q('w:t'))).strip() for pp in tbl.iter(q('w:p'))]
    textos = [t for t in textos if t]
    todo = ' '.join(textos)
    m = re.match(r'CERTUS PLD\s*[│|:—–-]?\s*(.*)', todo)
    celdas = tbl.findall('.//' + q('w:tc'))
    unica = len(celdas) == 1
    if not m and not unica:
        return None, None
    resto = m.group(1) if m else todo
    # primera línea: código y título; segunda: lo que siga a un segundo separador de sujeto
    l1 = textos[0]; l2 = ' '.join(textos[1:])
    l1 = re.sub(r'^CERTUS PLD\s*[│|:—–-]?\s*', '', l1).strip()
    if not l1 and len(textos) > 1:
        l1, l2 = textos[1], ' '.join(textos[2:])
    if not l2 and re.search(r'\s[│|]\s', l1):
        partes = re.split(r'\s+[│|]\s+', l1)
        l1, l2 = partes[0], '   │   '.join(partes[1:])
    l1 = re.sub(r'^(D\d+[A-Z]?(?:-[A-Z]+)*)\s+[—–-]\s+', r'\1 ', l1)
    l1 = re.sub(r'\s*\(Manual de Cumplimiento\)', '', l1)
    t = etree.fromstring(
        '<w:tbl xmlns:w="%s"><w:tblPr><w:tblW w:w="9500" w:type="dxa"/><w:jc w:val="center"/>'
        '<w:tblBorders><w:top w:val="single" w:sz="12" w:space="0" w:color="1F3864"/><w:left w:val="nil"/>'
        '<w:bottom w:val="single" w:sz="12" w:space="0" w:color="1F3864"/><w:right w:val="nil"/>'
        '<w:insideH w:val="nil"/><w:insideV w:val="nil"/></w:tblBorders><w:tblLayout w:type="fixed"/>'
        '<w:tblCellMar><w:top w:w="100" w:type="dxa"/><w:left w:w="120" w:type="dxa"/><w:bottom w:w="100" w:type="dxa"/>'
        '<w:right w:w="120" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="9500"/></w:tblGrid>'
        '<w:tr><w:tc><w:tcPr><w:tcW w:w="9500" w:type="dxa"/></w:tcPr></w:tc></w:tr></w:tbl>' % W)
    tc = t.find('.//' + q('w:tc'))
    p1 = Par.from_text('CERTUS PLD   │   ' + l1)
    p1.props = [frozenset(set(pr) | {'b'}) for pr in p1.props]
    reglas(p1)
    e1 = escribir_par(p1, 'CertusTabla')
    for r in e1.iter(q('w:r')):
        rpr = r.find(q('w:rPr'))
        etree.SubElement(rpr, q('w:color')).set(q('w:val'), '1F3864')
        etree.SubElement(rpr, q('w:sz')).set(q('w:val'), '26'); etree.SubElement(rpr, q('w:szCs')).set(q('w:val'), '26')
    tc.append(e1)
    if l2:
        p2 = Par.from_text(re.sub(r'\s*[│|]\s*', '   │   ', l2))
        reglas(p2)
        tc.append(escribir_par(p2, 'CertusTabla'))
    return t, l1

def procesar(contenedor, entrada, salida):
    LOG.clear()
    src = Document(entrada)
    out = []
    titulo_doc = None
    primera_tabla = True
    for el in src.element.body:
        if el.tag == q('w:sectPr'):
            continue
        if el.tag == q('w:tbl'):
            if primera_tabla:
                primera_tabla = False
                b, titulo_doc = banda(el)
                if b is not None:
                    out.append(b); continue
            t = copy.deepcopy(el)
            if len(t.findall('.//' + q('w:tc'))) == 1:
                normalizar_tabla(t, reglas, encabezado=False)
                tcpr = t.find('.//' + q('w:tcPr'))
                shd = etree.Element(q('w:shd')); shd.set(q('w:val'), 'clear'); shd.set(q('w:color'), 'auto'); shd.set(q('w:fill'), 'EDF0F5')
                tcw = tcpr.find(q('w:tcW'))
                (tcw.addnext(shd) if tcw is not None else tcpr.insert(0, shd))
                bd = t.find('%s/%s' % (q('w:tblPr'), q('w:tblBorders')))
                for e in list(bd):
                    e.set(q('w:val'), 'nil')
                l = bd.find(q('w:left')); l.set(q('w:val'), 'single'); l.set(q('w:sz'), '24'); l.set(q('w:color'), '1F3864'); l.set(q('w:space'), '0')
            else:
                normalizar_tabla(t, reglas)
            out.append(t); continue
        if el.tag != q('w:p'):
            continue
        texto = ''.join(t.text or '' for t in el.iter(q('w:t'))).strip()
        par = leer_par(el)
        if par is None:
            out.append(normalizar_parrafo_complejo(copy.deepcopy(el), 'CertusTexto')); continue
        if not texto and '\f' not in par.text:
            continue
        # nota de edición
        if par.mark == 'NotaCertus' or texto.startswith('NOTA DE EDICIÓN CERTUS'):
            n = Par.from_text(B.NOTA)
            n.props = [frozenset({'rs:NotaCertus'})] * len(n.chars); n.mark = 'NotaCertus'
            out.append(escribir_par(n, 'CertusTexto')); continue
        # línea de cierre
        if re.match(r'^(CERTUS CONSULTORES|Certus Consultores)\b.*(Mérida|999 163 1363)', texto) and len(texto) < 220:
            out.append(escribir_par(Par.from_text(B.PIE), 'CertusPie')); continue
        st = pstyle(el)
        heading_style = st.lower().startswith(('heading', 'ttulo', 'título', 'titulo'))
        negrita = es_negrita(par)
        canon_top = titulo_canonico(texto) if (negrita or heading_style or sombreado_oscuro(el)) and len(texto) < 170 else None
        if canon_top and not texto.endswith('.') or (canon_top and heading_style):
            np = Par.from_text(canon_top)
            # conserva variantes del título original, si las hay
            if any(any(x.startswith('rs:') for x in pr) for pr in par.props):
                np = par; np.replace(r'^\s*', '', count=1, protect=False)
                np.props = [frozenset(x for x in pr if x != 'b') for pr in np.props]
            reglas(np)
            out.append(escribir_par(np, 'CertusCapitulo')); continue
        if (heading_style or sombreado_oscuro(el)) and len(texto) < 170:
            np = par; np.props = [frozenset(x for x in pr if x not in ('b', 'i')) for pr in np.props]
            reglas(np); np.strip()
            out.append(escribir_par(np, 'CertusCapitulo' if (st.lower().endswith(('1', '2')) or sombreado_oscuro(el)) else 'CertusSeccion')); continue
        if negrita and len(texto) < 140 and not par.bullet and (SUBT.match(texto) or not texto.endswith(('.', ':', ';', ','))):
            np = par; np.props = [frozenset(x for x in pr if x not in ('b',)) for pr in np.props]
            reglas(np); np.strip()
            out.append(escribir_par(np, 'CertusSeccion')); continue
        reglas(par); par.strip()
        if not par.chars:
            continue
        nuevo = escribir_par(par, 'CertusVineta' if par.bullet else 'CertusTexto')
        a = jc(el)
        if a in ('center', 'right') and not par.bullet:
            ppr = nuevo.find(q('w:pPr'))
            etree.SubElement(ppr, q('w:jc')).set(q('w:val'), a)
            # el orden del esquema: pStyle … jc … rPr
            rpr = ppr.find(q('w:rPr'))
            if rpr is not None:
                ppr.remove(rpr); ppr.append(rpr)
        out.append(nuevo)
    # tablas consecutivas: separador para que Word no las fusione
    final = []
    for el in out:
        if final and final[-1].tag == q('w:tbl') and el.tag == q('w:tbl'):
            final.append(escribir_par(Par.from_text(''), 'CertusTexto'))
        final.append(el)
    espaciar_tras_tablas(final)
    doc = Document(contenedor)
    instalar_estilos(doc)
    body = doc.element.body; sect = body.find(q('w:sectPr'))
    for el in list(body):
        if el is not sect:
            body.remove(el)
    for el in final:
        sect.addprevious(el)
    numero_de_pagina(doc)
    cp = doc.core_properties
    cp.title = ('CERTUS PLD — ' + titulo_doc) if titulo_doc else os.path.splitext(os.path.basename(salida))[0]
    cp.author = 'Certus Consultores · Espinosa-Salazar Abogados Corporativos'
    cp.subject = 'Programa Certus PLD — edición homologada noviembre 2026'
    doc.save(salida)
    return list(LOG)

if __name__ == '__main__':
    print('\n'.join(procesar(*sys.argv[1:4])))
