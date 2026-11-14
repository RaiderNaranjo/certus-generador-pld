"""D11B — Casos Prácticos, homologado para los cuatro kits.

Uso: python3 build_d11b.py <dir_contenedor> <dir_salida>
Toma los casos vigentes de cada kit (src/d11b/<KIT>.docx), normaliza títulos y redacción, y agrega el caso del
Módulo 7 donde falta. Un caso por módulo del D11 (COR conserva un segundo caso del Módulo 5).
"""
import os, sys, re, copy
sys.path.insert(0, os.path.dirname(__file__))
import build_d5 as B
from build_d5 import Ctx, p, tabla, portada, pie, escribir_par, Par, aplicar_globales
from docxlib import instalar_estilos, q, espaciar_tras_tablas, leer_par
from docx import Document
import canon
import build_d11 as D11

MODULOS = {1: 'Marco normativo PLD/FT', 2: 'Señales de alerta del sector',
           3: 'Identificación y conocimiento del cliente; Beneficiario Controlador',
           4: "PEP's y debida diligencia reforzada", 6: 'Avisos, Informes y operaciones inusuales',
           7: 'Reforma RCG 2026: EBR, Grado de Riesgo, mecanismos automatizados y auditoría'}

def texto(el):
    return ''.join(t.text or '' for t in el.iter(q('w:t')))

def leer_casos(path):
    """Devuelve [{titulo, modulo, situacion:[Par], preguntas:[Par], respuesta: Par}]."""
    doc = Document(path)
    casos = []; cur = None; modo = None
    for el in doc.element.body:
        tag = el.tag.split('}')[1]
        if tag == 'p':
            t = texto(el).strip()
            m = re.match(r'Caso (\d+): (.*)', t)
            if m:
                cur = dict(titulo=m.group(2), modulo=None, situacion=[], preguntas=[], respuesta=None); casos.append(cur); modo = None; continue
            if cur is None or not t:
                continue
            mm = re.match(r'Módulo (\d+)', t)
            if mm and cur['modulo'] is None:
                cur['modulo'] = int(mm.group(1)); continue
            if t == 'Situación':
                modo = 'situacion'; continue
            if t.startswith('Preguntas de análisis'):
                modo = 'preguntas'; continue
            if t.startswith('CERTUS CONSULTORES'):
                continue
            if modo:
                cur[modo].append(leer_par(el))
        elif tag == 'tbl' and cur is not None:
            t = texto(el).strip()
            if t.startswith('RESPUESTA MODELO'):
                pars = [leer_par(pp) for pp in el.iter(q('w:p'))]
                pars = [pp for pp in pars if ''.join(pp.chars).strip()]
                # la primera línea es el rótulo "RESPUESTA MODELO — Certus PLD" o va pegado al texto
                cuerpo = []
                for pp in pars:
                    s = ''.join(pp.chars)
                    if re.fullmatch(r'RESPUESTA MODELO\s*[—-]?\s*(Certus PLD)?\s*', s.strip()):
                        continue
                    pp.replace(r'^RESPUESTA MODELO\s*[—-]\s*(Certus PLD\s*)?', '', protect=False)
                    cuerpo.append(pp)
                cur['respuesta'] = cuerpo
    return casos

REEMPLAZOS = [(r'\bel RC\b', 'el Representante Encargado de Cumplimiento'),
              (r'\bal RC\b', 'al Representante Encargado de Cumplimiento'),
              (r'\bEl RC\b', 'El Representante Encargado de Cumplimiento'),
              (r'\bdel RC\b', 'del Representante Encargado de Cumplimiento'),
              (r'\bRC\b', 'Representante Encargado de Cumplimiento'),
              (r'(D8(?:-[A-Z]+)?), 6 pasos', r'\1, procedimiento de identificación'),
              (r'\b6 pasos de identificación', 'procedimiento de identificación'),
              (r'SPPLD', 'Portal en Internet del SAT'),
              (r'Secc\. ', 'Sección ')]

def limpiar(par):
    for a, b in REEMPLAZOS:
        par.replace(a, b)
    aplicar_globales(par)
    return par

NUEVOS = {
 'NOT': dict(titulo='El programa que se quedó en papel', modulo=7,
    situacion=['En enero de 2028 el Notario Titular revisa su programa PLD. Durante 2027 clasificó a todos sus clientes como '
               'de riesgo bajo "porque son conocidos de la Notaría", no ha reevaluado a ninguno desde su primer acto, lleva el '
               'control de actos y Avisos en un cuaderno, y no ha contratado la auditoría porque considera que "eso es para '
               'empresas grandes".'],
    preguntas=['¿Puede clasificarse a todos los clientes como de riesgo bajo sin aplicar la Metodología EBR?',
               '¿Cada cuánto debe reevaluarse el Grado de Riesgo de un cliente?',
               '¿Un cuaderno cumple con los mecanismos automatizados que exige la Ley?',
               '¿Está obligada la Notaría a la auditoría anual y quién puede practicarla?'],
    respuesta=['No. El Grado de Riesgo de cada cliente se asigna con un modelo coherente con la Metodología EBR, con al menos '
               'tres grados (art. 23 Bis RCG); el riesgo bajo solo procede con los criterios que fije el Manual (art. 17 RCG). '
               'La clasificación es exigible para los actos realizados a partir del 1 de marzo de 2027 (Transitorio Cuarto del '
               'Acuerdo 115/2026). El Grado de Riesgo se determina al alta y se reevalúa al menos cada seis meses, con mayor '
               'frecuencia a mayor riesgo (art. 23 Bis 1 RCG; D6-NOT, sección IX). Los mecanismos automatizados pueden apoyarse '
               'en hojas de cálculo o bases de datos, siempre que cumplan las funciones del art. 41 RCG y sean verificables (art. '
               '3, fr. XI Ter, RCG); un cuaderno no consolida actos, no genera alertas ni conserva el historial, y debían estar '
               'en operación desde el 1 de junio de 2027 (Transitorio Noveno). La auditoría anual es obligatoria para todo '
               'sujeto obligado, sin importar su tamaño (art. 18, fr. XI, de la Ley; art. 42 RCG); el primer periodo es 2028 '
               '(Transitorio Octavo) y el dictamen se entrega a más tardar el último día hábil de marzo de 2029 (art. 50 RCG). '
               'Si el riesgo de la Notaría es bajo o medio, puede dictaminar un área de auditoría o control interno '
               'independiente; si es alto, o si así lo elige, un auditor externo independiente con certificación vigente de '
               'la UIF (arts. 44 y 45 RCG; D13-NOT).']),
 'COR': dict(titulo='El programa que se quedó en papel', modulo=7,
    situacion=['En enero de 2028 el Corredor Público Titular revisa su programa PLD. Durante 2027 clasificó a todos sus clientes '
               'como de riesgo bajo "porque son empresas que ya conoce", no ha reevaluado a ninguno, lleva el control de actos y '
               'Avisos en un cuaderno, y no ha contratado la auditoría porque cree que "solo aplica a las sociedades".'],
    preguntas=['¿Puede clasificarse a todos los clientes como de riesgo bajo sin aplicar la Metodología EBR?',
               '¿Cada cuánto debe reevaluarse el Grado de Riesgo de un cliente?',
               '¿Un cuaderno cumple con los mecanismos automatizados que exige la Ley?',
               '¿Está obligada la Correduría a la auditoría anual y quién puede practicarla?'],
    respuesta=['No. El Grado de Riesgo de cada cliente se asigna con un modelo coherente con la Metodología EBR, con al menos '
               'tres grados (art. 23 Bis RCG); el riesgo bajo solo procede con los criterios que fije el Manual (art. 17 RCG). '
               'La clasificación es exigible para los actos realizados a partir del 1 de marzo de 2027 (Transitorio Cuarto del '
               'Acuerdo 115/2026). El Grado de Riesgo se determina al alta y se reevalúa al menos cada seis meses, con mayor '
               'frecuencia a mayor riesgo (art. 23 Bis 1 RCG; D6-COR, sección IX). Los mecanismos automatizados pueden apoyarse '
               'en hojas de cálculo o bases de datos, siempre que cumplan las funciones del art. 41 RCG y sean verificables (art. '
               '3, fr. XI Ter, RCG); un cuaderno no consolida actos, no genera alertas ni conserva el historial, y debían estar '
               'en operación desde el 1 de junio de 2027 (Transitorio Noveno). La auditoría anual es obligatoria para todo '
               'sujeto obligado, persona física o moral (art. 18, fr. XI, de la Ley; art. 42 RCG); el primer periodo es 2028 '
               '(Transitorio Octavo) y el dictamen se entrega a más tardar el último día hábil de marzo de 2029 (art. 50 RCG). '
               'Si el riesgo de la Correduría es bajo o medio, puede dictaminar un área de auditoría o control interno '
               'independiente; si es alto, o si así lo elige, un auditor externo independiente con certificación vigente de '
               'la UIF (arts. 44 y 45 RCG; D13-COR).']),
 'PROF': dict(titulo='El programa que se quedó en papel', modulo=7,
    situacion=['En enero de 2028 la contadora Laura, que organiza aportaciones de capital y administra sociedades de sus '
               'clientes (incisos d) y e)), revisa su programa PLD. Durante 2027 clasificó a todos sus clientes como de riesgo '
               'bajo "porque son clientes de años", no ha reevaluado a ninguno, lleva el control en un cuaderno y no ha '
               'contratado la auditoría porque cree que "eso es para los despachos grandes".'],
    preguntas=['¿Puede clasificarse a todos los clientes como de riesgo bajo sin aplicar la Metodología EBR?',
               '¿Cada cuánto debe reevaluarse el Grado de Riesgo de un cliente?',
               '¿Un cuaderno cumple con los mecanismos automatizados que exige la Ley?',
               '¿Está obligada a la auditoría anual y quién puede practicarla?'],
    respuesta=['No. El Grado de Riesgo de cada cliente se asigna con un modelo coherente con la Metodología EBR, con al menos '
               'tres grados (art. 23 Bis RCG); el riesgo bajo solo procede con los criterios que fije el Manual (art. 17 RCG). '
               'La clasificación es exigible para los actos realizados a partir del 1 de marzo de 2027 (Transitorio Cuarto del '
               'Acuerdo 115/2026). El Grado de Riesgo se determina al alta y se reevalúa al menos cada seis meses, con mayor '
               'frecuencia a mayor riesgo (art. 23 Bis 1 RCG; D6-PROF, sección IX). Los mecanismos automatizados pueden apoyarse '
               'en hojas de cálculo o bases de datos, siempre que cumplan las funciones del art. 41 RCG y sean verificables (art. '
               '3, fr. XI Ter, RCG); un cuaderno no consolida actos, no genera alertas ni conserva el historial, y debían estar '
               'en operación desde el 1 de junio de 2027 (Transitorio Noveno). La auditoría anual es obligatoria para todo '
               'sujeto obligado, sin importar su tamaño (art. 18, fr. XI, de la Ley; art. 42 RCG); el primer periodo es 2028 '
               '(Transitorio Octavo) y el dictamen se entrega a más tardar el último día hábil de marzo de 2029 (art. 50 RCG). '
               'Si su riesgo es bajo o medio, puede dictaminar un área de auditoría o control interno independiente; si es '
               'alto, o si así lo elige, un auditor externo independiente con certificación vigente de la UIF (arts. 44 y 45 '
               'RCG; D13-PROF).']),
}

def callout(pars):
    """Recuadro de respuesta modelo: celda sombreada con borde izquierdo azul."""
    from lxml import etree
    W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    t = etree.fromstring(
        '<w:tbl xmlns:w="%s"><w:tblPr><w:tblW w:w="9500" w:type="dxa"/><w:jc w:val="center"/>'
        '<w:tblBorders><w:top w:val="nil"/><w:left w:val="single" w:sz="24" w:space="0" w:color="1F3864"/>'
        '<w:bottom w:val="nil"/><w:right w:val="nil"/><w:insideH w:val="nil"/><w:insideV w:val="nil"/></w:tblBorders>'
        '<w:tblLayout w:type="fixed"/><w:tblCellMar><w:top w:w="100" w:type="dxa"/><w:left w:w="160" w:type="dxa"/>'
        '<w:bottom w:w="100" w:type="dxa"/><w:right w:w="160" w:type="dxa"/></w:tblCellMar></w:tblPr>'
        '<w:tblGrid><w:gridCol w:w="9500"/></w:tblGrid><w:tr><w:trPr><w:cantSplit/></w:trPr><w:tc><w:tcPr><w:tcW w:w="9500" w:type="dxa"/>'
        '<w:shd w:val="clear" w:color="auto" w:fill="EDF0F5"/></w:tcPr></w:tc></w:tr></w:tbl>' % W)
    tc = t.find('.//' + q('w:tc'))
    tc.append(escribir_par(Par.from_text('**Respuesta modelo**'), 'CertusTabla'))
    for pp in pars:
        tc.append(escribir_par(pp, 'CertusTabla'))
    return t

def construir(kit):
    ctx = Ctx(kit); s = D11.S[kit]['suf']
    m5 = D11.S[kit]['m5'][0]
    casos = leer_casos(os.path.join(os.path.dirname(__file__), '..', 'src', 'd11b', '%s.docx' % kit))
    if kit in NUEVOS:
        n = NUEVOS[kit]
        casos.append(dict(titulo=n['titulo'], modulo=n['modulo'],
                          situacion=[Par.from_text(x) for x in n['situacion']],
                          preguntas=[Par.from_text('%d. %s' % (i, x)) for i, x in enumerate(n['preguntas'], 1)],
                          respuesta=[Par.from_text(x) for x in n['respuesta']]))
    casos.sort(key=lambda c: c['modulo'])
    portada(ctx, codigo='D11B', titulo='Casos Prácticos',
            indice='Índice: ' + ' │ '.join('Caso %d. %s' % (i, c['titulo']) for i, c in enumerate(casos, 1)))
    p(ctx, 'Material de capacitación del Programa Anual de Capacitación (D11%s). Cada caso corresponde a un módulo: se lee en '
           'grupo, se responden las preguntas de análisis y después se revisa la respuesta modelo. Las cifras en pesos usan la '
           'UMA de 2026 ($117.31).' % s)
    for i, c in enumerate(casos, 1):
        ctx.body.append(escribir_par(Par.from_text('Caso %d: %s' % (i, c['titulo'])), 'CertusCapitulo'))
        nombre = m5 if c['modulo'] == 5 else MODULOS[c['modulo']]
        p(ctx, '**Módulo %d — %s**' % (c['modulo'], nombre))
        ctx.body.append(escribir_par(Par.from_text('Situación'), 'CertusSeccion'))
        for pp in c['situacion']:
            ctx.body.append(escribir_par(limpiar(pp), 'CertusTexto'))
        ctx.body.append(escribir_par(Par.from_text('Preguntas de análisis'), 'CertusSeccion'))
        for pp in c['preguntas']:
            ctx.body.append(escribir_par(limpiar(pp), 'CertusTexto'))
        ctx.body.append(callout([limpiar(pp) for pp in c['respuesta']]))
    pie(ctx)
    return ctx, len(casos)

NOMBRES = {'PROF': 'profesionistas/D11B-PROF_Casos_Practicos_v2.docx', 'INM': 'inmobiliaria/D11B_Casos_Practicos_CertusPLD_v5.docx',
           'NOT': 'notarial/D11B-NOT_Casos_Practicos_v3.docx', 'COR': 'corredores/D11B-COR_Casos_Practicos_v2.docx'}

if __name__ == '__main__':
    B.SRC, OUT = sys.argv[1], sys.argv[2]
    for k in ('PROF', 'INM', 'NOT', 'COR'):
        ctx, n = construir(k)
        doc = Document(os.path.join(B.SRC, 'PROF_pr12.docx'))
        instalar_estilos(doc)
        body = doc.element.body; sect = body.find(q('w:sectPr'))
        for el in list(body):
            if el is not sect: body.remove(el)
        espaciar_tras_tablas(ctx.body)
        for el in ctx.body: sect.addprevious(el)
        cp = doc.core_properties
        cp.title = 'D11B%s Casos Prácticos — %s' % (canon.KITS[k]['suf'], canon.KITS[k]['variante'])
        cp.author = 'Certus Consultores · Espinosa-Salazar Abogados Corporativos'
        cp.subject = 'Programa Certus PLD — edición homologada noviembre 2026'
        dest = os.path.join(OUT, NOMBRES[k]); os.makedirs(os.path.dirname(dest), exist_ok=True)
        from docxlib import numero_de_pagina; numero_de_pagina(doc)
        doc.save(dest); print(k, n, 'casos', dest)
