"""Construye los cuatro D5 homologados (Manual de Políticas Internas) desde una sola especificación.

Uso: python3 build_d5.py <dir_fuentes> <dir_salida>
Fuentes esperadas en <dir_fuentes>: PROF_pr12, PROF_pr2, INM_pr13, INM_pr2, NOT_pr2, COR_pr2 (.docx)
Contenedor (encabezado/pie con logotipos, estilos de variantes): PROF_pr12.docx
"""
import copy, re, sys, os
sys.path.insert(0, os.path.dirname(__file__))
from docx import Document
from lxml import etree
from docxlib import (q, Par, leer_par, escribir_par, normalizar_parrafo_complejo, normalizar_tabla,
                     tabla_simple, instalar_estilos, LOG, log, W)
import canon

SRC, OUT = sys.argv[1], sys.argv[2]
CH_RE = re.compile(r'^(Capítulo|CAPÍTULO)\s+(\d+)[\.:]?\s*(.*)$')
AN_RE = re.compile(r'^Anexo\s+([A-D])[\.:]\s*(.*)$')
SUB_RE = re.compile(r'^(\d+\.\d+(?:\.\d+)?)\s+\S')

def txt(el):
    return ''.join(t.text or '' for t in el.iter(q('w:t'))).strip()

# ---------------------------------------------------------------- fuentes
class Fuente:
    def __init__(self, nombre, chmap, submap):
        self.nombre = nombre
        self.doc = Document(os.path.join(SRC, nombre + '.docx'))
        self.chmap, self.submap = chmap, submap
        self.caps = {}     # 'C5' / 'AA' -> {'body': [els], 'subs': {'5.1': (head, [els])}, 'orden': [...]}
        cur = {'body': [], 'subs': {}, 'orden': []}; self.caps['pre'] = cur; sub = None
        for el in self.doc.element.body:
            if el.tag == q('w:sectPr'):
                continue
            t = txt(el) if el.tag == q('w:p') else ''
            if el.tag == q('w:p') and not t and el.find('.//' + q('w:drawing')) is None:
                continue
            m, a = CH_RE.match(t), AN_RE.match(t)
            if m or a:
                key = ('C' + m.group(2)) if m else ('A' + a.group(1))
                cur = {'body': [], 'subs': {}, 'orden': []}; self.caps[key] = cur; sub = None
                continue
            ms = SUB_RE.match(t) if el.tag == q('w:p') else None
            if ms and key_ok(ms.group(1), self.caps, cur):
                sub = ms.group(1); cur['subs'][sub] = (el, []); cur['orden'].append(sub)
                continue
            if sub:
                cur['subs'][sub][1].append(el)
            else:
                cur['body'].append(el)

    def body(self, ch):
        return self.caps['C%s' % ch if isinstance(ch, int) else ch]['body']
    def sub(self, ch, num):
        return self.caps['C%s' % ch]['subs'][num]

def key_ok(num, caps, cur):
    return True

# Mapas de remisión: número en la fuente -> número canónico
def m(**kw):
    return {k.replace('_', '.')[1:]: v for k, v in kw.items()}
PROF12_CH = {1: 1, 2: 2, 3: 3, 4: 4, 5: 5, 6: 6, 7: 7, 8: 8, 9: 9, 10: 10, 11: 11, 12: 12, 13: 13, 14: 14,
             15: 16, 16: 17, 17: 18, 18: 19, 19: 20, 20: 21, 21: 22, 22: 23}
PR2_CH = {1: 1, 2: 2, 3: 3, 4: 4, 5: 5, 6: 6, 7: 8, 8: 9, 9: 10, 10: 11, 11: 12, 12: 13, 13: 14, 14: 15,
          15: 16, 16: 17, 17: 18, 18: 19, 19: 20, 20: 21, 23: 22, 24: 23}
FEP_CH = {1: 1, 2: 2, 3: 3, 4: 4, 5: 5, 6: 6, 7: 8, 8: 9, 9: 10, 10: 11, 11: 12, 12: 13, 13: 14, 14: 15,
          15: 16, 16: 17, 17: 21, 18: 21, 19: 18, 20: 19, 21: 20, 23: 22, 24: 23}
INM_CH = {1: 1, 2: 2, 3: 3, 4: 4, 5: 5, 6: 5, 7: 7, 8: 6, 9: 14, 10: 8, 11: 9, 12: 10, 13: 15, 14: 11,
          15: 16, 16: 18, 17: 12, 18: 19, 19: 13, 20: 17, 22: 22, 23: 23}
SUB_COMUN = {'1.2': '1.2', '2.1': '2.2', '2.2': '2.3', '3.1': '3.1', '3.2': '3.2', '3.3': '3.3',
             '4.1': '4.1', '4.2': '4.2', '5.1': '5.1', '5.2': '5.2', '5.3': '5.3', '5.4': '5.4',
             '6.1': '6.1', '9.1': '10.1', '10.1': '11.1', '11.1': '12.1', '12.1': '13.1',
             '13.1': '14.1', '13.2': '14.2', '15.1': '16.1', '16.1': '17.1', '23.1': '22.1'}
SUB_COR = dict(SUB_COMUN, **{'1.3': '1.2', '1.2': None})
SUB_PROF12 = {'1.1': '1.1', '4.1': '4.1'}
SUB_INM = {'1.1': '1.1', '1.2': '1.2', '1.3': '3.3', '2.1': '2', '2.2': '2.1', '2.3': '2.2', '2.4': '2.3',
           '3.1': '3.1', '3.1.1': '3.2', '3.2': '20', '4.1': '4', '4.2': '4.1', '4.3': '4', '4.4': '4.2',
           '5.1': '5.1', '6.1': '5.2', '6.2': '5.3', '6.3': '5.4', '8.1': '6.1', '9.1': '14.1', '9.2': '14.2',
           '10.1': '8.1', '12.1': '10.1', '14.1': '11.1', '15.1': '16.1', '17.1': '12.1', '19.1': '13.1',
           '20.1': '17.1', '22.1': '22.1'}

REF_RE = re.compile(r'(?P<kw>\b(?:[Cc]ap[ií]tulos?|Caps?\.|[Ss]ecci(?:ón|ones))\s+)'
                    r'(?P<nums>\d+(?:\.\d+)*(?:(?:\s*,\s*|\s+y\s+|\s+a\s+)\d+(?:\.\d+)*)*)')

def remapear(par, fuente):
    if fuente.chmap is None:
        return
    text = par.text
    def f(mm):
        cola = text[mm.end():mm.end() + 40]
        if re.match(r'\s*(?:,\s*)?(?:del|de la|en el) (?:Documento |documento )?D\d', cola) or \
           re.match(r'\s*(?:del|de las|de la) (?:RCG|Reglas|Ley|LFPIORPI)', cola):
            return mm.group(0)
        kw = mm.group('kw')
        if mm.start() > 0:
            kw = kw[0].lower() + kw[1:]
        def num(n):
            if '.' in n:
                v = fuente.submap.get(n, 'X')
                if v == 'X':
                    log('%s: remisión a %s sin mapa — se conserva' % (fuente.nombre, n)); return n
                if v is None:
                    log('%s: remisión a %s (suprimida)' % (fuente.nombre, n)); return n
                return v
            v = fuente.chmap.get(int(n))
            if v is None:
                log('%s: remisión a capítulo %s sin destino (contexto: %s)' % (fuente.nombre, n, text[max(0, mm.start()-60):mm.end()+30]))
                return n
            return str(v)
        nums = re.sub(r'\d+(?:\.\d+)*', lambda x: num(x.group(0)), mm.group('nums'))
        nuevo = kw + nums
        if nuevo != mm.group(0):
            log('%s: «%s» → «%s»' % (fuente.nombre, mm.group(0), nuevo))
        return nuevo
    par.replace(REF_RE, f, protect=False)

# ---------------------------------------------------------------- reglas de texto
def V(pf, pm):
    """Segmentos de variante."""
    out = []
    if pf is not None:
        out.append(('[PF: %s]' % pf, {'rs:VariantePF'}))
    if pm is not None:
        out.append(('[PM: %s]' % pm, {'rs:VariantePM'}))
    return out

def es_inicio(s, i):
    j = i - 1
    while j >= 0 and s[j] in ' \t':
        j -= 1
    return j < 0 or s[j] in '.?!\n' or s[j:i].strip() == ''

GLOBAL = []
def regla(pat, repl, flags=0):
    GLOBAL.append((re.compile(pat, flags), repl))

def _art(mm):
    s = mm.string
    return mm.group(0) if es_inicio(s, mm.start()) else mm.group(0).lower()
regla(r'\bArts?\.(?=\s)', _art)
regla(r'\bArtículos?(?=\s+\d)', lambda mm: mm.group(0) if es_inicio(mm.string, mm.start()) else mm.group(0).lower())
regla(r'\bart,\.\s*', 'art. ')
regla(r'\bfracc?\.\s*(?=[IVXL])', 'fr. ')
regla(r'\bfracc\s+(?=[IVXL])', 'fr. ')
regla(r'\bFracc\.', 'Fr.')
regla(r'\b([Aa])rt\. 2, (?:fracción|fr\.) XI Bis', r'\1rt. 3, fr. XI Bis')
regla(r'(?<=[IVX])RCG', ' RCG')
regla(r'Personas? Expuestas? Políticamente', lambda mm: mm.group(0).replace('Expuestas Políticamente', 'Políticamente Expuestas').replace('Expuesta Políticamente', 'Políticamente Expuesta'))
regla(r'(?<!Encargado )(?<!Encargada )\bRepresentante de Cumplimiento\b', 'Representante Encargado de Cumplimiento')
regla(r'Certus Consultores(?= *[│|] *Certus PLD)', 'Certus Consultores · Espinosa-Salazar Abogados Corporativos')
regla(r'desarrollado por Certus Consultores a través', 'desarrollado por Certus Consultores y Espinosa-Salazar Abogados Corporativos a través')
regla(r'Certus Consultores, a través del Programa Certus PLD, notificará', 'Certus Consultores y Espinosa-Salazar Abogados Corporativos, a través del Programa Certus PLD, notificarán')
regla(r'con el apoyo de Certus Consultores(?! y)', 'con el apoyo de Certus Consultores y Espinosa-Salazar Abogados Corporativos')
regla(r'edición octubre 2026', 'edición noviembre 2026')
regla(r'https?://ppld\.sat\.gob\.mx/?', 'https://sppld.sat.gob.mx/pld/index.html')
regla(r'(?<=\S)  +(?=\S)', ' ')
regla(r'\(\s+', '(')
regla(r'\bFracción (XI|XII|V)\b(?= del)', r'fracción \1')

def aplicar_globales(par):
    for pat, repl in GLOBAL:
        par.replace(pat, repl)

# Reglas PF/PM para el texto del PR #2 (redactado antes de la convención de variantes)
REGLAS_PROF = [
    (r'\bal Profesional Titular o, en el despacho, el Representante (?:Encargado )?de Cumplimiento', V('al Profesional Titular', 'al Representante Encargado de Cumplimiento')),
    (r'\bpor el Profesional Titular o, en el despacho, el Representante (?:Encargado )?de Cumplimiento', V('por el Profesional Titular', 'por el Representante Encargado de Cumplimiento')),
    (r'\bEl Profesional Titular o, en el despacho, el Representante (?:Encargado )?de Cumplimiento', V('El Profesional Titular', 'El Representante Encargado de Cumplimiento')),
    (r'\bel Profesional Titular o, en el despacho, el Representante (?:Encargado )?de Cumplimiento', V('el Profesional Titular', 'el Representante Encargado de Cumplimiento')),
    (r'\bdel Profesional Titular o el órgano de administración del despacho', V('del Profesional Titular', 'del órgano de administración')),
    (r'\bel Profesional Titular o el órgano de administración del despacho', V('el Profesional Titular', 'el órgano de administración')),
    (r'\b([Ee]l|del) profesional o despacho', r'\1 profesional'),
]
REGLAS_INM = [
    (r'\bal Representante Encargado de Cumplimiento', V('al titular', 'al Representante Encargado de Cumplimiento')),
    (r'\bdel Representante Encargado de Cumplimiento', V('del titular', 'del Representante Encargado de Cumplimiento')),
    (r'\bEl Representante Encargado de Cumplimiento', V('El titular', 'El Representante Encargado de Cumplimiento')),
    (r'\bel Representante Encargado de Cumplimiento', V('el titular', 'el Representante Encargado de Cumplimiento')),
    (r'\bdel órgano de administración', V('del titular', 'del órgano de administración')),
    (r'\bel órgano de administración', V('el titular', 'el órgano de administración')),
]
PEP_ALTO_PROF = ('Cuando el cliente sea PEP y además de Riesgo alto, el acto u operación requiere '
                 '[PM: la aprobación previa y por escrito del socio director o su equivalente señalado en el Anexo C (art. 23 Ter 5, primer párrafo, RCG)]'
                 '[PF: la constancia firmada por el Profesional Titular en la que señale los motivos considerados (art. 23 Ter 5, segundo párrafo, RCG)], '
                 'que toma en cuenta el cuestionario reforzado, el origen de los recursos y el monitoreo propuesto.')
PEP_ALTO_INM = ('Cuando el cliente sea PEP y además de Riesgo alto, el acto u operación requiere '
                '[PM: la aprobación previa y por escrito del directivo o su equivalente señalado en el Anexo C (art. 23 Ter 5, primer párrafo, RCG)]'
                '[PF: la constancia firmada por el titular en la que señale los motivos considerados (art. 23 Ter 5, segundo párrafo, RCG)], '
                'que toma en cuenta el cuestionario reforzado, el origen de los recursos y el monitoreo propuesto.')
# Párrafos que se sustituyen completos (texto de la fuente -> texto nuevo con marcado)
PARRAFOS = {
    # PROF (PR #2)
    'En el despacho, el Representante reporta directamente al órgano de administración y su evaluación no depende de la facturación de los asuntos que revisa (criterio del Programa).':
        ('PM', 'Independencia: el Representante Encargado de Cumplimiento reporta directamente al órgano de administración y su evaluación no depende de la facturación de los asuntos que revisa (criterio del Programa).'),
    'Suplencia: en ausencias temporales, [NOMBRE Y CARGO DEL SUPLENTE] auxilia en las funciones operativas; todo cambio de Representante del despacho se actualiza en el Portal. El profesional persona física no delega su responsabilidad (art. 20, último párrafo, de la Ley).':
        (None, 'Suplencia: en ausencias temporales, [NOMBRE Y CARGO DEL SUPLENTE] auxilia en las funciones operativas[PM: ; todo cambio de Representante Encargado de Cumplimiento se actualiza en el Portal][PF: ; el profesional no delega su responsabilidad (art. 20, último párrafo, de la Ley)].'),
    'Cuando el cliente sea PEP y además de Riesgo alto, el acto u operación requiere si se trata de un despacho, la aprobación previa y por escrito de [SOCIO DIRECTOR O EQUIVALENTE QUE APRUEBA] (Art. 23 Ter 5, primer párrafo, RCG); si es persona física, la constancia firmada por el Profesional Titular con los motivos considerados (segundo párrafo), que considera el cuestionario reforzado, el origen de los recursos y el monitoreo propuesto.':
        (None, PEP_ALTO_PROF),
    'Vehículo: Portal en Internet del SAT, en el formato oficial vigente y con la e.firma con la que se realizó el alta y registro: la del profesional, si es persona física, o la del despacho, si es persona moral — en este caso no puede utilizarse la e.firma de su representante legal (Art. 24, segundo párrafo, RCG).':
        (None, 'Vehículo: Portal en Internet del SAT, en el formato oficial vigente y con la e.firma con la que se realizó el alta y registro: [PF: la del profesional (art. 24, primer párrafo, RCG).][PM: la de la persona moral; no puede utilizarse la e.firma de su representante legal ni la del Representante Encargado de Cumplimiento (art. 24, segundo párrafo, RCG).]'),
    # INM (PR #2)
    'Independencia de las áreas comerciales: reporta directamente al órgano de administración, y su evaluación y remuneración no dependen de metas de ventas (criterio del Programa).':
        ('PM', 'Independencia de las áreas comerciales: el Representante Encargado de Cumplimiento reporta directamente al órgano de administración, y su evaluación y remuneración no dependen de metas de ventas (criterio del Programa).'),
    'Recursos asignados por el órgano de administración: [RECURSOS ASIGNADOS — tiempo dedicado, personal de apoyo, herramientas y presupuesto de capacitación y asesoría].':
        (None, 'Recursos asignados[PM: \u00a0por el órgano de administración]: [RECURSOS ASIGNADOS — tiempo dedicado, personal de apoyo, herramientas y presupuesto de capacitación y asesoría].'),
    'Suplencia: en ausencias temporales, [NOMBRE Y CARGO DEL SUPLENTE] auxilia en las funciones operativas, sin que ello sustituya la designación ante la autoridad; todo cambio de Representante se actualiza en el Portal y en el Anexo C.':
        (None, 'Suplencia: en ausencias temporales, [NOMBRE Y CARGO DEL SUPLENTE] auxilia en las funciones operativas[PM: , sin que ello sustituya la designación ante la autoridad; todo cambio de Representante Encargado de Cumplimiento se actualiza en el Portal y en el Anexo C][PF: ; el titular no delega su responsabilidad (art. 20, último párrafo, de la Ley)].'),
    'Responsabilidad: responde ante el órgano de administración por la aplicación de este Manual; la responsabilidad ante la autoridad se rige por el Art. 20 de la Ley.':
        (None, 'Responsabilidad: [PM: el Representante Encargado de Cumplimiento responde ante el órgano de administración por la aplicación de este Manual; la responsabilidad ante la autoridad se rige por el art. 20 de la Ley.][PF: el titular responde personal y directamente ante la autoridad por el cumplimiento de la Ley (art. 20, último párrafo, de la Ley).]'),
    'Cuando el cliente sea PEP y además de Riesgo alto, el acto u operación requiere la aprobación previa y por escrito de [CARGO DEL DIRECTIVO QUE APRUEBA OPERACIONES CON PEP DE RIESGO ALTO] (Art. 23 Ter 5, primer párrafo, RCG), que considera el cuestionario reforzado, el origen de los recursos y el monitoreo propuesto.':
        (None, PEP_ALTO_INM),
    'Vehículo: Portal en Internet del SAT, en el formato oficial vigente y con la e.firma de la empresa (persona moral); no puede utilizarse la e.firma del representante legal ni del Representante Encargado de Cumplimiento (Art. 24, segundo párrafo, RCG).':
        (None, 'Vehículo: Portal en Internet del SAT, en el formato oficial vigente y con la e.firma [PF: del titular (art. 24, primer párrafo, RCG).][PM: de la empresa (persona moral); no puede utilizarse la e.firma del representante legal ni la del Representante Encargado de Cumplimiento (art. 24, segundo párrafo, RCG).]'),
    'Informe: el Representante Encargado de Cumplimiento documenta semestralmente los resultados, hallazgos y acciones correctivas con responsable y fecha y los presenta al órgano de administración; el auditor revisa su seguimiento en la auditoría anual.':
        (None, 'Informe: [PF: el titular documenta semestralmente los resultados, hallazgos y acciones correctivas con responsable y fecha][PM: el Representante Encargado de Cumplimiento documenta semestralmente los resultados, hallazgos y acciones correctivas con responsable y fecha y los presenta al órgano de administración]; el auditor revisa su seguimiento en la auditoría anual.'),
    'Modificaciones ordenadas por el SAT (Art. 37 Bis 3 RCG): el Representante Encargado de Cumplimiento registra el requerimiento, informa al órgano de administración dentro de los tres días hábiles siguientes, prepara la modificación del Manual o de los documentos señalados dentro del plazo que fije la autoridad, la somete a aprobación, la difunde y conserva la evidencia de su atención.':
        (None, 'Modificaciones ordenadas por el SAT (art. 37 Bis 3 RCG): [PF: el titular registra el requerimiento][PM: el Representante Encargado de Cumplimiento registra el requerimiento, informa al órgano de administración dentro de los tres días hábiles siguientes], prepara la modificación del Manual o de los documentos señalados dentro del plazo que fije la autoridad, la somete a aprobación, la difunde y conserva la evidencia de su atención.'),
    'Aprobación: el órgano de administración aprueba cada versión mediante acta o resolución; la versión rige desde su aprobación.':
        (None, 'Aprobación: [PF: el titular aprueba cada versión mediante constancia firmada][PM: el órgano de administración aprueba cada versión mediante acta o resolución]; la versión rige desde su aprobación.'),
    'Acceso restringido a expedientes, Avisos y alertas: únicamente el Representante Encargado de Cumplimiento, [PERSONAS AUTORIZADAS] y el auditor. En sistemas digitales se usan credenciales individuales y bitácora de accesos; el archivo físico se resguarda bajo llave.':
        (None, 'Acceso restringido a expedientes, Avisos y alertas: únicamente [PF: el titular][PM: el Representante Encargado de Cumplimiento], las personas que autorice ([PERSONAS AUTORIZADAS]) y el auditor. En sistemas digitales se usan credenciales individuales y bitácora de accesos; el archivo físico se resguarda bajo llave.'),
}

# ---------------------------------------------------------------- proceso de elementos
class Ctx:
    def __init__(self, kit):
        self.kit = kit
        self.body = []           # elementos de salida
        self.cap = 0
        self.sub = 0

def preparar_par(par, fuente, reglas_pfpm, edits):
    remapear(par, fuente)
    t = par.text.strip()
    if t in PARRAFOS:
        modo, nuevo = PARRAFOS[t]
        b = par.bullet
        npar = Par.from_text(nuevo)
        npar.bullet = b
        if modo == 'PM':
            npar = variante_completa(npar, 'PM')
        par.chars, par.props, par.mark, par.bullet = npar.chars, npar.props, npar.mark, npar.bullet
        log('%s: párrafo sustituido (variantes) «%s…»' % (fuente.nombre, t[:60]))
    else:
        for pat, repl in reglas_pfpm:
            n = par.replace(pat, repl)
            if n:
                log('%s: variante PF/PM en «%s…»' % (fuente.nombre, t[:70]))
    for pat, repl in edits:
        if par.replace(pat, repl):
            log('%s: ajuste «%s»' % (fuente.nombre, pat if isinstance(pat, str) else pat.pattern))
    aplicar_globales(par)

def variante_completa(par, tipo):
    """Convierte un párrafo completo en variante PF o PM (los marcadores […] quedan fuera del estilo)."""
    st = 'rs:Variante' + tipo
    text = par.text
    prot = [(mm.start(), mm.end()) for mm in re.finditer(r'\[(?!PF: |PM: )[^\[\]]*\]', text)]
    chars = list('[%s: ' % tipo); props = [frozenset({st})] * len(chars)
    for k, (c, p) in enumerate(zip(par.chars, par.props)):
        if any(a <= k < b for a, b in prot):
            chars.append(c); props.append(frozenset(x for x in p if not x.startswith('rs:')))
        else:
            chars.append(c); props.append(frozenset([x for x in p if not x.startswith('rs:')] + [st]))
    chars.append(']'); props.append(frozenset({st}))
    out = Par(chars, props, mark='Variante' + tipo, bullet=par.bullet)
    return out

def emitir(ctx, els, fuente, reglas_pfpm=(), edits=(), filtro=None, pm=False, quitar_inicio=None):
    for el in els:
        if el.tag == q('w:p'):
            t = txt(el)
            if filtro and not filtro(t):
                continue
            par = leer_par(el)
            if par is None:
                new = normalizar_parrafo_complejo(copy.deepcopy(el), 'CertusTexto')
                log('%s: párrafo complejo conservado «%s»' % (fuente.nombre, t[:60]))
                ctx.body.append(new); continue
            if not par.text.strip():
                continue
            preparar_par(par, fuente, reglas_pfpm, edits)
            if quitar_inicio:
                par.replace(quitar_inicio, '', count=1)
                par.strip()
                if par.chars:
                    par.chars[0] = par.chars[0].upper()
            par.strip()
            if not par.text.strip():
                continue
            if pm:
                par = variante_completa(par, 'PM')
            ctx.body.append(escribir_par(par, 'CertusVineta' if par.bullet else 'CertusTexto'))
        elif el.tag == q('w:tbl'):
            if filtro and not filtro('T|' + txt(el)):
                continue
            tbl = copy.deepcopy(el)
            def proc(par):
                preparar_par(par, fuente, reglas_pfpm, edits)
            normalizar_tabla(tbl, proc)
            ctx.body.append(tbl)

def capitulo(ctx, n):
    ctx.cap, ctx.sub = n, 0
    ctx.body.append(escribir_par(Par.from_text('Capítulo %d: %s' % (n, canon.titulo(ctx.kit, n))), 'CertusCapitulo'))

def seccion(ctx, titulo, pm=False, num=None):
    ctx.sub += 1
    par = Par.from_text('%d.%d %s' % (ctx.cap, ctx.sub, titulo))
    aplicar_globales(par)
    if pm:
        par = variante_completa(par, 'PM')
    ctx.body.append(escribir_par(par, 'CertusSeccion'))

def titulo_de(head_el, fuente):
    """Título de una sección de la fuente sin su número, con citas normalizadas."""
    par = leer_par(head_el)
    t = par.text
    par.replace(r'^\s*\d+\.\d+(?:\.\d+)?\s+', '', count=1, protect=False)
    return par

def sub(ctx, fuente, ch, num, titulo=None, reglas_pfpm=(), edits=(), pm=False, cuerpo=True):
    head, els = fuente.sub(ch, num)
    ctx.sub += 1
    if titulo is None:
        par = titulo_de(head, fuente)
    else:
        par = Par.from_text(titulo)
    for pat, repl in reglas_pfpm:
        if not isinstance(repl, list):
            par.replace(pat, repl)
    pre = '%d.%d ' % (ctx.cap, ctx.sub)
    par.chars = list(pre) + par.chars; par.props = [frozenset()] * len(pre) + par.props
    aplicar_globales(par)
    par.replace(r'\s+—\s+solo despachos \(personas morales\)$', '', protect=False)
    if pm:
        par = variante_completa(par, 'PM')
    ctx.body.append(escribir_par(par, 'CertusSeccion'))
    if cuerpo:
        emitir(ctx, els, fuente, reglas_pfpm, edits, pm=pm)

def p(ctx, texto, bullet=False, pm=False, estilo=None):
    par = Par.from_text(texto)
    par.bullet = bullet
    aplicar_globales(par)
    if pm:
        par = variante_completa(par, 'PM')
    ctx.body.append(escribir_par(par, estilo or ('CertusVineta' if bullet else 'CertusTexto')))

def tabla(ctx, filas, anchos, encabezado=True):
    tbl = tabla_simple(filas, anchos)
    normalizar_tabla(tbl, aplicar_globales, encabezado=encabezado)
    ctx.body.append(tbl)

def anexo(ctx, letra):
    for l, t, _ in canon.ANEXOS:
        if l == letra:
            ctx.body.append(escribir_par(Par.from_text('Anexo %s: %s' % (l, t)), 'CertusCapitulo'))
    ctx.cap = letra

# ---------------------------------------------------------------- portada, índice, glosario, pie
PORTADA = {
    'PROF': dict(etiqueta='[PF: Profesional:][PM: Razón o denominación social:]', valor='[Nombre]',
                 l2='RFC:', v2='[RFC]', banner2='Profesional / Despacho: [Nombre]',
                 aprobo='[PF: Profesional Titular][PM: Órgano de administración (Consejo de Administración o Administrador Único)]'),
    'INM':  dict(etiqueta='[PF: Titular:][PM: Razón o denominación social:]', valor='[EMPRESA]',
                 l2='RFC:', v2='[RFC_EMPRESA]', banner2='Empresa: [EMPRESA]',
                 aprobo='[PF: Titular (persona física)][PM: Órgano de administración (Consejo de Administración o Administrador Único)]'),
    'NOT':  dict(etiqueta='Notaría:', valor='[Denominación completa]', l2='Notario Titular:',
                 v2='[Nombre del Notario Titular]', banner2='Notaría: [Denominación completa]', aprobo='Notario Titular'),
    'COR':  dict(etiqueta='Correduría:', valor='[Denominación de la Correduría]', l2='Corredor Público Titular:',
                 v2='[Nombre del Corredor Público Titular]', banner2='Correduría: [Denominación de la Correduría]',
                 aprobo='Corredor Público Titular'),
}
NOTA = ('NOTA DE EDICIÓN CERTUS — se elimina al personalizar. Este Manual contiene variantes para cliente persona '
        'física, marcadas [PF: …] con sombreado verde, y para cliente persona moral, marcadas [PM: …] con sombreado '
        'amarillo. El texto sin marca aplica a ambos. Para personalizar, ejecute la macro ResolverComoPF o '
        'ResolverComoPM, o elimine a mano las variantes que no apliquen y retire las etiquetas.')
PIE = 'CERTUS CONSULTORES · ESPINOSA-SALAZAR ABOGADOS CORPORATIVOS  |  Certus PLD  |  Mérida, Yucatán  |  999 163 1363 · 999 575 3129'

def portada(ctx):
    k = ctx.kit; d = PORTADA[k]; info = canon.KITS[k]
    # Banda de título
    tbl = etree.fromstring(
        '<w:tbl xmlns:w="%s"><w:tblPr><w:tblW w:w="9500" w:type="dxa"/><w:jc w:val="center"/>'
        '<w:tblBorders><w:top w:val="single" w:sz="12" w:space="0" w:color="1F3864"/><w:left w:val="nil"/>'
        '<w:bottom w:val="single" w:sz="12" w:space="0" w:color="1F3864"/><w:right w:val="nil"/>'
        '<w:insideH w:val="nil"/><w:insideV w:val="nil"/></w:tblBorders><w:tblLayout w:type="fixed"/>'
        '<w:tblCellMar><w:top w:w="100" w:type="dxa"/><w:left w:w="120" w:type="dxa"/><w:bottom w:w="100" w:type="dxa"/>'
        '<w:right w:w="120" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="9500"/></w:tblGrid>'
        '<w:tr><w:tc><w:tcPr><w:tcW w:w="9500" w:type="dxa"/></w:tcPr></w:tc></w:tr></w:tbl>' % W)
    tc = tbl.find('.//' + q('w:tc'))
    l1 = Par.from_text('**CERTUS PLD   │   D5%s Manual de Políticas Internas PLD/FT — %s**' % (info['suf'], info['variante']))
    l2 = Par.from_text('%s   │   Variante: %s — %s' % (d['banner2'], info['variante'], info['frac']))
    e1 = escribir_par(l1, 'CertusTabla'); e2 = escribir_par(l2, 'CertusTabla')
    for r in e1.iter(q('w:r')):
        rpr = r.find(q('w:rPr'))
        if rpr is None: rpr = etree.SubElement(r, q('w:rPr')); r.insert(0, rpr)
        etree.SubElement(rpr, q('w:color')).set(q('w:val'), '1F3864')
        s = etree.SubElement(rpr, q('w:sz')); s.set(q('w:val'), '26')
        s = etree.SubElement(rpr, q('w:szCs')); s.set(q('w:val'), '26')
    tc.append(e1); tc.append(e2)
    ctx.body.append(tbl)
    ctx.body.append(escribir_par(Par.from_text(''), 'CertusTexto'))
    filas = [
        ('**%s**' % d['etiqueta'] if not d['etiqueta'].startswith('[') else d['etiqueta'], d['valor']),
        ('**%s**' % d['l2'], d['v2']),
        ('**Versión:**', '[VERSION] — [Fecha de adopción] │ Plantilla Certus PLD, edición noviembre 2026 (Acuerdo 115/2026; Capítulo X RCG)'),
        ('**Aprobó:**', d['aprobo']),
        ('**Elaboró:**', 'Certus Consultores · Espinosa-Salazar Abogados Corporativos │ Certus PLD'),
        ('**Vigencia:**', 'Un año a partir de la fecha de adopción — revisión anual obligatoria y cada vez que se reformen la Ley, su Reglamento o las Reglas de Carácter General'),
    ]
    t = tabla_simple(filas, [2900, 6600])
    normalizar_tabla(t, aplicar_globales, encabezado=False)
    # etiquetas en negrita también para las variantes de la primera fila
    if d['etiqueta'].startswith('['):
        for r in t.find(q('w:tr')).find(q('w:tc')).iter(q('w:r')):
            rpr = r.find(q('w:rPr'))
            etree.SubElement(rpr, q('w:b'))
    # columna de etiquetas sombreada
    for tr in t.findall(q('w:tr')):
        tcpr = tr.find(q('w:tc')).find(q('w:tcPr'))
        shd = etree.Element(q('w:shd')); shd.set(q('w:val'), 'clear'); shd.set(q('w:color'), 'auto'); shd.set(q('w:fill'), 'EDF0F5')
        tcpr.insert(1, shd)
    ctx.body.append(t)
    if info['pfpm']:
        par = Par.from_text(NOTA, base=frozenset({'rs:NotaCertus'}))
        # la nota no es variante: se quita el estilo de variante que from_text puso en […]
        par.props = [frozenset({'rs:NotaCertus'})] * len(par.chars)
        par.mark = 'NotaCertus'
        ctx.body.append(escribir_par(par, 'CertusTexto'))
    ctx.body.append(escribir_par(Par.from_text(canon.indice(k)), 'CertusIndice'))

GLOSARIO_COMUN = [
    ('Aviso', 'el que debe presentarse en términos de los artículos 17 y 18, fracción VI, de la Ley (art. 3, fr. II, RCG).'),
    ('Beneficiario Controlador', 'persona física o grupo de personas físicas que, directamente o por medio de otra persona, obtiene en última instancia el beneficio del acto u operación o ejerce el control efectivo de la persona moral cliente (art. 3, fr. IV, RCG).'),
    ('Certus PLD', 'Programa de Cumplimiento PLD/FT de Certus Consultores y Espinosa-Salazar Abogados Corporativos.'),
    ('EBR', 'Enfoque Basado en Riesgo: metodología de evaluación de Riesgos del Capítulo II Quáter de las RCG.'),
    ('Evaluación Nacional de Riesgos', 'documento que emite y actualiza la UIF y da a conocer en el Portal en Internet (art. 3, fr. VIII Bis, RCG).'),
    ('Grado de Riesgo', 'clasificación que se otorga a cada Cliente o Usuaria con base en la evaluación de su Riesgo (art. 3, fr. IX Bis, RCG).'),
    ('KYC', 'conocimiento del cliente (Know Your Customer): identificación y verificación del Cliente o Usuaria.'),
    ('Ley o LFPIORPI', 'Ley Federal para la Prevención e Identificación de Operaciones con Recursos de Procedencia Ilícita.'),
    ('Manual de Políticas Internas', 'el presente documento (art. 3, fr. XI Bis, y Capítulo X RCG).'),
    ('Mecanismos automatizados', 'procesos, sistemas o herramientas tecnológicas que permiten revisar de manera continua los actos u operaciones (art. 3, fr. XI Ter, RCG).'),
    ('Mitigantes', 'políticas, criterios, medidas y procedimientos que contribuyen a administrar y disminuir la exposición a los Riesgos identificados (art. 3, fr. XI Quinquies, RCG).'),
    ("PEP", 'Persona Políticamente Expuesta (art. 23 Quáter RCG).'),
    ('Perfil Transaccional', 'conjunto de elementos que permiten identificar el comportamiento esperado de los actos u operaciones de un Cliente o Usuaria (art. 3, fr. XI Sexties, RCG).'),
    ('PLD/FT', 'prevención de lavado de dinero y financiamiento al terrorismo.'),
    ('Portal en Internet', 'Sistema del Portal de Prevención de Lavado de Dinero del SAT (https://sppld.sat.gob.mx/pld/index.html), para el alta y registro y la presentación de Avisos e Informes (art. 3, fr. XII, RCG).'),
    ('RCG o Reglas', 'Reglas de Carácter General a que se refiere la Ley, reformadas por el Acuerdo 115/2026 (DOF 07/08/2026).'),
    ('Relación de negocios', 'la establecida de manera formal y habitual con un Cliente o Usuaria, excluidos los actos u operaciones ocasionales y los servicios de fe pública del art. 17, fr. XII, de la Ley (art. 3, fr. XIV, RCG).'),
    ('Riesgo', 'probabilidad de que las Actividades Vulnerables o quienes las realizan sean utilizados para actos u operaciones vinculados con Delitos de Operaciones con Recursos de Procedencia Ilícita, delitos relacionados o su financiamiento (art. 3, fr. XVI, RCG).'),
    ('SAT', 'Servicio de Administración Tributaria.'),
    ('UIF', 'Unidad de Inteligencia Financiera de la Secretaría de Hacienda y Crédito Público.'),
    ('UMA', 'Unidad de Medida y Actualización: $117.31 MXN diarios en 2026.'),
]
GLOSARIO_SECTOR = {
    'PROF': [('Momento de sujeción', 'distinción propia de la fracción XI entre la preparación del servicio (identificación) y la operación financiera en nombre y representación del cliente (Aviso).'),
             ('PM:Órgano de administración', 'Consejo de Administración, Administrador Único o, en sociedades civiles, los socios administradores de la persona moral.'),
             ('PM:Representante Encargado de Cumplimiento', 'persona física designada por la persona moral ante la Secretaría en términos del art. 20 de la Ley.')],
    'INM': [('PM:Órgano de administración', 'Consejo de Administración o Administrador Único de la persona moral.'),
            ('PM:Representante Encargado de Cumplimiento', 'persona física designada por la persona moral ante la Secretaría en términos del art. 20 de la Ley.')],
    'NOT': [('Encargado Operativo', 'persona de la Notaría que auxilia al Notario Titular en tareas operativas de cumplimiento, sin trasladarle su responsabilidad (D2-NOT).'),
            ('Notario Titular', 'fedatario a cargo de la Notaría, quien cumple personal y directamente las obligaciones de la Ley (art. 20, último párrafo, de la Ley).')],
    'COR': [('Corredor Público', 'particular habilitado por la Secretaría de Economía para ejercer, entre otras funciones, la fe pública en materia mercantil (Ley Federal de Correduría Pública).'),
            ('Encargado Operativo', 'persona de la Correduría que auxilia al Corredor Público Titular en tareas operativas de cumplimiento, sin trasladarle su responsabilidad (D2-COR).'),
            ('LFCP', 'Ley Federal de Correduría Pública.'),
            ('Póliza / Acta', 'instrumentos en que el Corredor hace constar, respectivamente, actos o contratos y hechos de naturaleza mercantil.')],
}

def glosario(ctx):
    capitulo(ctx, 23)
    items = GLOSARIO_COMUN + GLOSARIO_SECTOR[ctx.kit]
    items = sorted(items, key=lambda x: x[0].split(':')[-1].lower().replace('ó', 'o').replace('é', 'e'))
    for term, defi in items:
        if term.startswith('PM:'):
            p(ctx, '**%s:** %s' % (term[3:], defi), bullet=True, pm=True)
        else:
            p(ctx, '**%s:** %s' % (term, defi), bullet=True)

def anexo_a(ctx):
    anexo(ctx, 'A')
    p(ctx, 'Este Anexo relaciona cada rubro del contenido mínimo del Manual de Políticas Internas (art. 37 Bis RCG) y '
           'las demás disposiciones que deben constar en él con el capítulo de este Manual y el documento del Programa '
           'Certus PLD que los desarrollan (art. 37, segundo párrafo, RCG).')
    tabla(ctx, canon.anexo_a(ctx.kit), [1200, 4100, 2100, 2100])

def pie(ctx):
    ctx.body.append(escribir_par(Par.from_text(PIE), 'CertusPie'))

# ---------------------------------------------------------------- especificación por kit
def es_bc(t):
    return t.startswith(('Beneficiario Controlador', 'Orden de prelación', 'La evidencia del procedimiento'))
EXENCION = r'^(?:Incisos que no se prestan|Supuestos que la (?:Notaría|Correduría|empresa) no realiza) \([Aa]rt\. 37 Bis 2 RCG\)\.\s*'
TIT_EXENCION = 'Supuestos que no se realizan (art. 37 Bis 2 RCG)'
HONORABILIDAD = ('Si un funcionario o empleado deja de contar con calidad técnica u honorabilidad (por ejemplo, por '
                 'sentencia por delito patrimonial o inhabilitación), se le retira de inmediato de las funciones de '
                 'atención al público, identificación de clientes, envío de Avisos o administración de recursos; se '
                 'documenta el hecho y se resuelve su situación conforme a la legislación laboral (art. 39 Bis 2, último '
                 'párrafo, RCG). Cada empleado actualiza su declaración cuando cambie su situación y, en todo caso, una '
                 'vez al año (criterio del Programa).')

def intro(ctx, sujeto, destino):
    p(ctx, 'El presente Manual de Políticas Internas en materia de PLD/FT (“Manual”; art. 3, fr. XI Bis, y Capítulo X RCG) '
           'es adoptado por %s en cumplimiento de las obligaciones establecidas en la Ley Federal para la Prevención e '
           'Identificación de Operaciones con Recursos de Procedencia Ilícita (“Ley” o “LFPIORPI”), su Reglamento y las '
           'Reglas de Carácter General (“RCG”). Ha sido desarrollado por Certus Consultores y Espinosa-Salazar Abogados '
           'Corporativos a través del Programa Certus PLD, diseñado específicamente para %s.' % (sujeto, destino))

def construir_prof():
    ctx = Ctx('PROF')
    a = Fuente('PROF_pr12', PROF12_CH, SUB_PROF12)
    b = Fuente('PROF_pr2', PR2_CH, SUB_COMUN)
    R = REGLAS_PROF
    portada(ctx)
    capitulo(ctx, 1)
    intro(ctx, 'el profesional[PM: , mediante acuerdo de su órgano de administración,]',
          'prestadores de servicios profesionales sujetos a la fracción XI del artículo 17 de la Ley')
    emitir(ctx, a.body(1)[1:], a)
    sub(ctx, a, 1, '1.1'); sub(ctx, b, 1, '1.2', reglas_pfpm=R)
    capitulo(ctx, 2)
    cuerpo2 = a.body(2)
    emitir(ctx, [cuerpo2[0], cuerpo2[2]], a)
    seccion(ctx, TIT_EXENCION)
    emitir(ctx, [cuerpo2[1]], a, quitar_inicio=EXENCION)
    sub(ctx, b, 2, '2.1', reglas_pfpm=R); sub(ctx, b, 2, '2.2', reglas_pfpm=R)
    capitulo(ctx, 3)
    emitir(ctx, a.body(3), a)
    sub(ctx, b, 3, '3.1', reglas_pfpm=R)
    sub(ctx, b, 3, '3.2', titulo='Grupo Empresarial (art. 37 Bis 1 RCG)', pm=True, reglas_pfpm=[R[-1]])
    capitulo(ctx, 4)
    emitir(ctx, a.body(4), a); sub(ctx, a, 4, '4.1'); sub(ctx, b, 4, '4.2', reglas_pfpm=R)
    capitulo(ctx, 5)
    emitir(ctx, a.body(5), a, edits=[
        (r'siempre que este Manual documente los criterios que justifican la clasificación', 'conforme a los criterios de la sección 5.2'),
        (r'cuestionario de identificación reforzada y monitoreo', 'cuestionario de identificación reforzada (Anexo B) y monitoreo'),
        (r'si el cliente es además PEP, \[PF: el profesional en su calidad de\]\[PM: el\] Representante Encargado de Cumplimiento dejará constancia firmada de los motivos considerados \(art\. 23 Ter 5\)',
         'si el cliente es además PEP, se requiere la aprobación o constancia de la sección 5.3 (art. 23 Ter 5)')])
    for s in ('5.1', '5.2', '5.3', '5.4'):
        sub(ctx, b, 5, s, reglas_pfpm=R)
    capitulo(ctx, 6); emitir(ctx, a.body(6), a); sub(ctx, b, 6, '6.1', reglas_pfpm=R)
    capitulo(ctx, 7); emitir(ctx, a.body(7), a)
    capitulo(ctx, 8); emitir(ctx, a.body(8), a)
    capitulo(ctx, 9); emitir(ctx, a.body(9), a)
    capitulo(ctx, 10); emitir(ctx, a.body(10), a, edits=[(r'\(capítulo 7\)', '(capítulo 8)')])
    sub(ctx, b, 9, '9.1', reglas_pfpm=R)
    capitulo(ctx, 11); emitir(ctx, a.body(11), a); sub(ctx, b, 10, '10.1', reglas_pfpm=R)
    capitulo(ctx, 12); emitir(ctx, a.body(12), a); sub(ctx, b, 11, '11.1', reglas_pfpm=R)
    capitulo(ctx, 13); emitir(ctx, a.body(13), a, edits=[
        (r'Cuando el Riesgo del profesional sea bajo', 'Cuando el Riesgo del profesional, determinado conforme a la sección 5.1, sea bajo')])
    sub(ctx, b, 12, '12.1', reglas_pfpm=R)
    capitulo(ctx, 14); emitir(ctx, a.body(14), a)
    sub(ctx, b, 13, '13.1', reglas_pfpm=R); sub(ctx, b, 13, '13.2', reglas_pfpm=R)
    capitulo(ctx, 15); emitir(ctx, a.body(15)[:1], a)
    capitulo(ctx, 16); emitir(ctx, a.body(15)[1:], a); sub(ctx, b, 15, '15.1', reglas_pfpm=R)
    capitulo(ctx, 17); emitir(ctx, a.body(16), a); sub(ctx, b, 16, '16.1', reglas_pfpm=R)
    capitulo(ctx, 18); emitir(ctx, a.body(17), a)
    capitulo(ctx, 19); emitir(ctx, a.body(18), a)
    capitulo(ctx, 20)
    emitir(ctx, a.body(19), a, filtro=lambda t: not t.startswith('Las medidas ante la pérdida'))
    p(ctx, HONORABILIDAD)
    capitulo(ctx, 21); emitir(ctx, a.body(20), a)
    capitulo(ctx, 22); emitir(ctx, a.body(21), a); sub(ctx, b, 23, '23.1', reglas_pfpm=R)
    glosario(ctx)
    anexos_bcd(ctx, b, R)
    pie(ctx)
    return ctx

def anexos_bcd(ctx, f, R):
    anexo_a(ctx)
    anexo(ctx, 'B'); emitir(ctx, f.body('AB'), f, reglas_pfpm=R)
    anexo(ctx, 'C'); emitir(ctx, f.body('AC'), f, reglas_pfpm=R)
    if ctx.kit in ('PROF', 'INM'):
        extra = 'socio director o equivalente' if ctx.kit == 'PROF' else 'directivo o equivalente'
        tbl = ctx.body[-1]
        if tbl.tag == q('w:tbl'):
            filas = tbl.findall(q('w:tr'))
            nueva = copy.deepcopy(filas[-1])
            celdas = nueva.findall(q('w:tc'))
            for k, tc in enumerate(celdas):
                for pp in tc.findall(q('w:p')):
                    tc.remove(pp)
                texto = ('Persona moral: %s que aprueba actos u operaciones con PEP de Riesgo alto (art. 23 Ter 5 RCG)' % extra) if k == 0 else ''
                tc.append(escribir_par(Par.from_text(texto), 'CertusTabla'))
            # se inserta antes de la fila del flujo del acto u operación (última)
            filas[-1].addprevious(nueva)
        else:
            log('Anexo C: no se encontró la tabla para agregar la fila del directivo')
    anexo(ctx, 'D'); emitir(ctx, f.body('AD'), f, reglas_pfpm=R,
                            filtro=lambda t: not t.upper().startswith('CERTUS CONSULTORES  |'))

def construir_fe_publica(kit):
    ctx = Ctx(kit)
    f = Fuente(kit + '_pr2', FEP_CH, SUB_COR if kit == 'COR' else SUB_COMUN)
    nota, notario = ('Notaría', 'Notario Titular') if kit == 'NOT' else ('Correduría', 'Corredor Público Titular')
    portada(ctx)
    capitulo(ctx, 1)
    if kit == 'NOT':
        intro(ctx, 'el Notario Titular', 'notarías sujetas a la fracción XII, Apartado A, del artículo 17 de la Ley')
        p(ctx, 'Para efectos de este Manual, “la Notaría” designa a la notaría pública a cargo del Notario Titular, quien '
               'cumple personal y directamente las obligaciones que la Ley establece (art. 20, último párrafo, de la Ley); '
               '“el Notario” designa al Notario Titular.')
    else:
        intro(ctx, 'el Corredor Público Titular', 'corredores públicos sujetos a la fracción XII, Apartado B, del artículo 17 de la Ley')
        p(ctx, 'Para efectos de este Manual, “la Correduría” designa a la correduría pública a cargo del Corredor Público '
               'Titular, quien cumple personal y directamente las obligaciones que la Ley establece (art. 20, último párrafo, '
               'de la Ley); “el Corredor” designa al Corredor Público Titular.')
    emitir(ctx, f.body(1)[1:], f)
    sub(ctx, f, 1, '1.1')
    sub(ctx, f, 1, '1.3' if kit == 'COR' else '1.2')
    capitulo(ctx, 2)
    c2 = f.body(2)
    emitir(ctx, [c2[0]] + c2[2:], f)
    seccion(ctx, TIT_EXENCION)
    emitir(ctx, [c2[1]], f, quitar_inicio=EXENCION)
    sub(ctx, f, 2, '2.1'); sub(ctx, f, 2, '2.2')
    capitulo(ctx, 3); emitir(ctx, f.body(3), f)
    for s in ('3.1', '3.2', '3.3'):
        sub(ctx, f, 3, s)
    capitulo(ctx, 4); emitir(ctx, f.body(4), f); sub(ctx, f, 4, '4.1'); sub(ctx, f, 4, '4.2')
    capitulo(ctx, 5); emitir(ctx, f.body(5), f)
    for s in ('5.1', '5.2', '5.3', '5.4'):
        sub(ctx, f, 5, s)
    capitulo(ctx, 6); emitir(ctx, f.body(6), f, filtro=lambda t: not es_bc(t)); sub(ctx, f, 6, '6.1')
    capitulo(ctx, 7); emitir(ctx, f.body(6), f, filtro=es_bc)
    capitulo(ctx, 8); emitir(ctx, f.body(7), f)
    capitulo(ctx, 9); emitir(ctx, f.body(8), f)
    capitulo(ctx, 10); emitir(ctx, f.body(9), f); sub(ctx, f, 9, '9.1')
    capitulo(ctx, 11); emitir(ctx, f.body(10), f); sub(ctx, f, 10, '10.1')
    capitulo(ctx, 12); emitir(ctx, f.body(11), f); sub(ctx, f, 11, '11.1')
    capitulo(ctx, 13); emitir(ctx, f.body(12), f); sub(ctx, f, 12, '12.1')
    capitulo(ctx, 14); emitir(ctx, f.body(13), f); sub(ctx, f, 13, '13.1'); sub(ctx, f, 13, '13.2')
    capitulo(ctx, 15); emitir(ctx, f.body(14), f)
    capitulo(ctx, 16); emitir(ctx, f.body(15), f); sub(ctx, f, 15, '15.1')
    capitulo(ctx, 17); emitir(ctx, f.body(16), f); sub(ctx, f, 16, '16.1')
    capitulo(ctx, 18); emitir(ctx, f.body(19), f)
    capitulo(ctx, 19); emitir(ctx, f.body(20), f)
    capitulo(ctx, 20); emitir(ctx, f.body(21), f)
    capitulo(ctx, 21)
    acto = 'protocolización' if kit == 'NOT' else 'formalización'
    seccion(ctx, 'Identificación de quien solicita materialmente la %s (art. 13 RCG)' % acto)
    emitir(ctx, f.body(17), f)
    seccion(ctx, 'Fe pública prestada al Poder Judicial o a autoridades administrativas (art. 13 Bis RCG)')
    emitir(ctx, f.body(18), f)
    capitulo(ctx, 22); emitir(ctx, f.body(23), f); sub(ctx, f, 23, '23.1')
    glosario(ctx)
    anexos_bcd(ctx, f, ())
    pie(ctx)
    return ctx

def construir_inm():
    ctx = Ctx('INM')
    a = Fuente('INM_pr13', INM_CH, SUB_INM)
    b = Fuente('INM_pr2', INM_CH, SUB_INM)
    R = REGLAS_INM
    portada(ctx)
    capitulo(ctx, 1)
    intro(ctx, '[EMPRESA][PM: , mediante acuerdo de su órgano de administración,]',
          'inmobiliarias sujetas a las fracciones V, V Bis y XV del artículo 17 de la Ley')
    emitir(ctx, a.body(1)[1:], a)
    sub(ctx, a, 1, '1.1'); sub(ctx, b, 1, '1.2', reglas_pfpm=R)
    capitulo(ctx, 2)
    emitir(ctx, a.body(2), a)
    p(ctx, 'Este Manual es de observancia obligatoria para:')
    emitir(ctx, a.sub(2, '2.1')[1], a)
    seccion(ctx, TIT_EXENCION)
    emitir(ctx, b.sub(2, '2.2')[1], b, reglas_pfpm=R)
    sub(ctx, b, 2, '2.3', reglas_pfpm=R); sub(ctx, b, 2, '2.4', reglas_pfpm=R)
    capitulo(ctx, 3); emitir(ctx, a.body(3), a)
    sub(ctx, a, 3, '3.1')
    sub(ctx, b, 3, '3.1.1', reglas_pfpm=R)
    sub(ctx, b, 1, '1.3', titulo='Grupo Empresarial (art. 37 Bis 1 RCG)', pm=True)
    capitulo(ctx, 4)
    emitir(ctx, a.sub(4, '4.1')[1], a); emitir(ctx, a.sub(4, '4.3')[1], a)
    sub(ctx, a, 4, '4.2', titulo='Umbrales de identificación y Aviso (art. 17, fr. V, V Bis y XV, de la Ley)')
    sub(ctx, b, 4, '4.4', reglas_pfpm=R)
    capitulo(ctx, 5)
    emitir(ctx, a.body(5), a)
    emitir(ctx, a.body(6), a, edits=[
        (r'siempre que la empresa haya documentado en este Manual los criterios que justifican dicha clasificación', 'conforme a los criterios de la sección 5.2 de este Manual'),
        (r'\(formato Certus PLD\) para clientes', '(Anexo B de este Manual) para clientes')])
    sub(ctx, b, 5, '5.1', reglas_pfpm=R)
    for s in ('6.1', '6.2', '6.3'):
        sub(ctx, b, 6, s, reglas_pfpm=R)
    capitulo(ctx, 6); emitir(ctx, a.body(8), a); sub(ctx, b, 8, '8.1', reglas_pfpm=R)
    capitulo(ctx, 7); emitir(ctx, a.body(7), a)
    capitulo(ctx, 8); emitir(ctx, a.body(10), a); sub(ctx, a, 10, '10.1')
    capitulo(ctx, 9); emitir(ctx, a.body(11), a)
    capitulo(ctx, 10)
    vbis = lambda t: t.startswith('Para la fracción V Bis')
    emitir(ctx, a.body(12), a, filtro=lambda t: not vbis(t))
    sub(ctx, b, 12, '12.1', reglas_pfpm=R)
    capitulo(ctx, 11); emitir(ctx, a.body(14), a); sub(ctx, b, 14, '14.1', reglas_pfpm=R)
    capitulo(ctx, 12); emitir(ctx, a.body(17), a, edits=[
        (r'las medidas previstas en el Documento D11C', 'las medidas de la sección 12.1'),
        (r'Ver Documentos D11, D11B y D11C, y la Declaración de Selección de Personal \(capítulo 20\)\.',
         'Ver Documentos D11, D11B, D11C y D11D (Declaración de Selección de Personal, capítulo 20).')])
    sub(ctx, b, 17, '17.1', reglas_pfpm=R)
    capitulo(ctx, 13); emitir(ctx, a.body(19), a, edits=[
        (r'conforme a la metodología EBR,', 'conforme a la evaluación de la sección 5.1,')])
    sub(ctx, b, 19, '19.1', reglas_pfpm=R)
    capitulo(ctx, 14); emitir(ctx, a.body(9), a); sub(ctx, b, 9, '9.1', reglas_pfpm=R); sub(ctx, b, 9, '9.2', reglas_pfpm=R)
    capitulo(ctx, 15); emitir(ctx, a.body(13), a)
    capitulo(ctx, 16); emitir(ctx, a.body(15), a); sub(ctx, b, 15, '15.1', reglas_pfpm=R)
    capitulo(ctx, 17); emitir(ctx, a.body(20), a); sub(ctx, b, 20, '20.1', reglas_pfpm=R)
    capitulo(ctx, 18); emitir(ctx, a.body(16), a)
    capitulo(ctx, 19); emitir(ctx, a.body(18), a)
    capitulo(ctx, 20); emitir(ctx, b.sub(3, '3.2')[1], b, reglas_pfpm=R)
    capitulo(ctx, 21)
    p(ctx, 'El alcance de la sujeción se determina por fracción del artículo 17 de la Ley y por acto u operación. Los '
           'umbrales de identificación y de Aviso de cada fracción se señalan en la sección 4.1, y las fracciones o '
           'supuestos que la empresa no realiza se hacen constar en la sección 2.1.')
    seccion(ctx, 'Desarrollo e intermediación inmobiliaria (fr. V)')
    p(ctx, 'Para efectos del Aviso, la fecha del acto u operación es aquella en que se tenga por liquidado y se alcance '
           'el umbral de Aviso (art. 24 Bis, fr. III, RCG).')
    seccion(ctx, 'Recepción de recursos para Desarrollos Inmobiliarios (fr. V Bis)')
    p(ctx, 'La fecha del acto u operación que se considera para el Aviso es aquella en que se recibió y destinó la última '
           'aportación al Desarrollo Inmobiliario en el mes calendario (art. 24 Bis, fr. IV, RCG).')
    emitir(ctx, a.body(12), a, filtro=vbis)
    seccion(ctx, 'Arrendamiento (fr. XV)')
    p(ctx, 'Cuando el uso o goce de un inmueble se otorgue por copropietarios, cada uno cumple sus obligaciones conforme a '
           'la Ley y considera como monto del acto u operación el total de la renta mensual pactada; en el Aviso señala lo '
           'relativo al pago de la renta que recibe en lo individual (art. 24 Bis 6 RCG).')
    capitulo(ctx, 22); emitir(ctx, a.body(22), a); sub(ctx, b, 22, '22.1', reglas_pfpm=R)
    glosario(ctx)
    anexos_bcd(ctx, b, R)
    pie(ctx)
    return ctx

# ---------------------------------------------------------------- ensamblado
NOMBRES = {
    'PROF': 'profesionistas/D5-PROF_Manual_Politicas_Internas_PLD_v3.docx',
    'INM': 'inmobiliaria/D5_Manual_Politicas_Internas_PLD_v6.docx',
    'NOT': 'notarial/D5-NOT_Manual_Politicas_Internas_PLD_v4.docx',
    'COR': 'corredores/D5-COR_Manual_Politicas_Internas_PLD_v3.docx',
}

def guardar(ctx):
    doc = Document(os.path.join(SRC, 'PROF_pr12.docx'))
    instalar_estilos(doc)
    body = doc.element.body
    sect = body.find(q('w:sectPr'))
    for el in list(body):
        if el is not sect:
            body.remove(el)
    for el in ctx.body:
        sect.addprevious(el)
    cp = doc.core_properties
    cp.title = 'D5%s Manual de Políticas Internas PLD/FT — %s' % (canon.KITS[ctx.kit]['suf'], canon.KITS[ctx.kit]['variante'])
    cp.author = 'Certus Consultores · Espinosa-Salazar Abogados Corporativos'
    cp.subject = 'Programa Certus PLD — edición homologada noviembre 2026'
    destino = os.path.join(OUT, NOMBRES[ctx.kit])
    os.makedirs(os.path.dirname(destino), exist_ok=True)
    doc.save(destino)
    return destino

if __name__ == '__main__':
    kits = sys.argv[3:] or ['PROF', 'INM', 'NOT', 'COR']
    for k in kits:
        LOG.clear()
        ctx = {'PROF': construir_prof, 'INM': construir_inm}.get(k, lambda: construir_fe_publica(k))()
        ruta = guardar(ctx)
        with open(os.path.join(OUT, 'log_%s.txt' % k), 'w') as fh:
            fh.write('\n'.join(LOG))
        print(k, '->', ruta, len(LOG), 'entradas de registro')
