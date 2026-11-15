"""D13 — Dictamen de Auditoría PLD/FT, homologado para los cuatro kits con las secciones del art. 47 RCG.

Uso: python3 build_d13.py <dir_contenedor> <dir_salida>
"""
import os, sys
sys.path.insert(0, os.path.dirname(__file__))
import build_d5 as B
from build_d5 import Ctx, p, tabla, portada, pie, escribir_par, Par
from docxlib import instalar_estilos, q, espaciar_tras_tablas
from docx import Document
import canon

SECCIONES = [
    ('I', 'Presentación (art. 47, fr. I, RCG)', 'Presentación'),
    ('II', 'Objeto (art. 47, fr. II, RCG)', 'Objeto'),
    ('III', 'Volumen de Información y Muestreo (art. 47, fr. III, RCG)', 'Muestreo'),
    ('IV', 'Proceso de Auditoría (arts. 46 y 47, fr. IV, RCG)', 'Proceso'),
    ('V', 'Hallazgos (art. 47, fr. V, RCG)', 'Hallazgos'),
    ('VI', 'Resultados de Cumplimiento (arts. 47, fr. VI, y 48 RCG)', 'Resultados'),
    ('VII', 'Acciones Correctivas y Recomendaciones de Mejora (arts. 47, fr. VII, y 49 RCG)', 'Acciones correctivas'),
    ('VIII', 'Entrega, Seguimiento y Firmas (arts. 42, 50 y 51 RCG)', 'Firmas'),
]

def sec(ctx, rom):
    t = [s[1] for s in SECCIONES if s[0] == rom][0]
    ctx.body.append(escribir_par(Par.from_text('Sección %s: %s' % (rom, t)), 'CertusCapitulo'))

def sub(ctx, t):
    ctx.body.append(escribir_par(Par.from_text(t), 'CertusSeccion'))

def ficha(ctx, filas):
    tabla(ctx, [('**%s**' % a if not a.startswith('[') else a, b) for a, b in filas], [3600, 5900], encabezado=False)

def indice():
    return 'Índice: ' + ' │ '.join('%s. %s' % (r, c) for r, _, c in SECCIONES)

S = {
 'INM': dict(suf='', suj='la empresa', auditado='[PF: Titular de la Actividad Vulnerable][PM: Denominación de la empresa]',
    recibe='[PF: Titular][PM: Órgano de administración o administrador único]',
    rec='[PF: el titular][PM: el Representante Encargado de Cumplimiento]',
    actos='[ ] Fr. V — compraventa o intermediación   [ ] Fr. V Bis — recepción de recursos para desarrollos   [ ] Fr. XV — arrendamiento',
    d2='[PF: Cumplimiento personal del titular (art. 20, cuarto párrafo, de la Ley; D2-PF)][PM: Designación y vigencia del Representante Encargado de Cumplimiento (art. 20 de la Ley; art. 10 RCG; D2)]',
    sector=[('Desarrollos y arrendamiento: Aviso consolidado por desarrollo y copropietarios (arts. 24 Bis 1 y 24 Bis 6 RCG)', 'D12, sección I')],
    bc='D7, Sección E; D14', efectivo='art. 32, fr. I y VII, de la Ley'),
 'NOT': dict(suf='-NOT', suj='la Notaría', auditado='Notario Público No. [__] — [Nombre del Notario Titular]',
    recibe='Notario Titular', rec='el Notario Titular',
    actos='[ ] a) Inmuebles   [ ] b) Poderes irrevocables   [ ] c) Sociedades   [ ] d) Fideicomisos   [ ] e) Mutuos o créditos',
    d2='Cumplimiento personal del Notario Titular y, en su caso, Encargado Operativo (art. 20, cuarto párrafo, de la Ley; D2-NOT)',
    sector=[('Identificación de la persona que solicita materialmente la protocolización (art. 13 RCG)', 'D7-NOT, Sección A.4'),
            ('Identificación de la forma de pago en el instrumento (art. 33 de la Ley)', 'D7-NOT, Sección I')],
    bc='D7-NOT, Sección E', efectivo='art. 32, fr. I y VI, de la Ley'),
 'COR': dict(suf='-COR', suj='la Correduría', auditado='Corredor Público No. [__] — [Nombre del Corredor Público Titular]',
    recibe='Corredor Público Titular', rec='el Corredor Público Titular',
    actos='[ ] a) Avalúos   [ ] b) Sociedades mercantiles   [ ] c) Fideicomisos   [ ] d) Mutuos o créditos mercantiles',
    d2='Cumplimiento personal del Corredor Público Titular y, en su caso, Encargado Operativo (art. 20, cuarto párrafo, de la Ley; D2-COR)',
    sector=[('Identificación de la persona que solicita materialmente la formalización (art. 13 RCG)', 'D7-COR, Sección A.4'),
            ('Identificación de la forma de pago en el instrumento (art. 33 de la Ley)', 'D7-COR, Sección I')],
    bc='D7-COR, Sección E', efectivo='art. 32, fr. VI, de la Ley'),
 'PROF': dict(suf='-PROF', suj='[PF: el profesional][PM: el despacho]', auditado='[PF: Nombre del profesional][PM: Denominación del despacho]',
    recibe='[PF: Profesional Titular][PM: Órgano de administración del despacho]',
    rec='[PF: el profesional][PM: el Representante Encargado de Cumplimiento]',
    actos='[ ] a) Inmuebles   [ ] b) Administración de recursos   [ ] c) Cuentas   [ ] d) Aportaciones de capital   [ ] e) Personas morales o vehículos corporativos',
    d2='[PF: Cumplimiento personal del profesional (art. 20, cuarto párrafo, de la Ley; D2-PROF-PF)][PM: Designación y vigencia del Representante Encargado de Cumplimiento (art. 20 de la Ley; art. 10 RCG; D2-PROF-PM)]',
    sector=[('Distinción entre preparación y operación financiera en representación del cliente, y secreto profesional', 'D1-PROF, secciones 2.2 y IV')],
    bc='D7-PROF, Sección E', efectivo='art. 32 de la Ley, en los actos que se preparen o ejecuten'),
}

def obligaciones(d, D):
    filas = [
        ('Alta, registro y actualización en el padrón (art. 18, fr. IV Bis, de la Ley; arts. 4 y 7 RCG)', D('3')),
        (d['d2'], D('2') if not d['d2'].startswith('[') else 'D2'),
        ('Manual de Políticas Internas (art. 18, fr. VIII, de la Ley; arts. 37 a 37 Bis 3 RCG)', D('5')),
        ('Metodología de evaluación de Riesgos y su actualización (art. 18, fr. VII, de la Ley; arts. 10 Septies a 10 Septies 3 RCG)', D('6')),
        ('Clasificación del Grado de Riesgo de los clientes (arts. 23 Bis a 23 Bis 4 RCG)', '%s, sección XI; %s, Sección D' % (D('6'), D('7'))),
        ('Identificación y expediente único; verificación anual (art. 18, fr. I y II, de la Ley; arts. 12 y 21 RCG)', D('7')),
        ('Conocimiento del cliente, Perfil Transaccional y sistema de alertas (arts. 23 Ter a 23 Ter 2 RCG)', '%s, Sección G' % D('7')),
        ('Beneficiario Controlador (art. 18, fr. III, de la Ley; arts. 23 Quinquies a 23 Quinquies 2 RCG)', d['bc']),
        ("Personas Políticamente Expuestas (arts. 23 Ter 3 a 23 Ter 5, 23 Quáter y 23 Quáter 1 RCG)", D('8')),
    ] + d['sector'] + [
        ('Avisos mensuales e informes en ceros, en tiempo y forma (arts. 17, 23 y 24 de la Ley; arts. 24 a 25 RCG)', D('12')),
        ('Seguimiento y acumulación de actos u operaciones, cuando aplique (art. 17, penúltimo párrafo, de la Ley; art. 19 RCG)', D('12')),
        ('Avisos de 24 horas y consulta de listas (art. 18, fr. VI, de la Ley; arts. 26 Bis a 27 y 38 RCG)', '%s; %s, Sección H' % (D('9'), D('7'))),
        ('Restricción del uso de efectivo (%s)' % d['efectivo'], D('12')),
        ('Conservación de información y documentación (art. 18, fr. IV, de la Ley)', D('10')),
        ('Capacitación anual y su evidencia (art. 18, fr. IX, de la Ley; arts. 39 Bis y 39 Bis 1 RCG)', '%s, %s y %s' % (D('11'), D('11B'), D('11C'))),
        ('Selección de personal (art. 39 Bis 2 RCG)', D('11D')),
        ('Mecanismos automatizados (art. 18, fr. X, de la Ley; art. 41 RCG)', '%s, capítulo 19' % D('5')),
        ('Confidencialidad (art. 50 de la Ley; art. 31 RCG)', '%s, sección V' % D('9')),
        ('Atención de las acciones correctivas del dictamen anterior (art. 49, segundo párrafo, RCG)', 'Dictamen del año anterior'),
    ]
    return filas

def construir(kit):
    d = S[kit]; ctx = Ctx(kit); s = d['suf']
    D = lambda n: 'D%s%s' % (n, s)
    portada(ctx, codigo='D13', titulo='Dictamen de Auditoría PLD/FT', indice=indice())
    p(ctx, 'Este formato sigue las secciones mínimas del dictamen anual (art. 47, segundo párrafo, RCG). La auditoría evalúa y '
           'dictamina, del 1 de enero al 31 de diciembre, la efectividad del cumplimiento de la Ley, su Reglamento y las RCG '
           '(art. 18, fr. XI, de la Ley; art. 42 RCG). El primer periodo es 2028 (Transitorio Octavo del Acuerdo 115/2026); en el '
           'primer año de operaciones, el periodo va desde el inicio de la Actividad Vulnerable hasta el 31 de diciembre del año '
           'siguiente (art. 43 RCG). Si este formato se usa para la revisión de cierre de la implementación del Programa, se '
           'trata de un servicio contractual de verificación y no del dictamen del art. 42 RCG.')
    p(ctx, '**Quién emite el dictamen.** Si el Riesgo del auditado es bajo o medio, puede firmarlo su área de auditoría o '
           'control interno, independiente %s (art. 44 RCG); Certus PLD puede acompañarla con la metodología, estos papeles '
           'de trabajo, una revisión de calidad y la capacitación anual que exige el art. 44, sin firmar el dictamen. Si el '
           'Riesgo es alto, o si el auditado lo elige, lo emite una persona auditora externa independiente con certificación '
           'vigente de la UIF que no le haya prestado servicios con conflicto de interés durante el periodo (art. 45 RCG).' % (('de ' + d['rec']) if d['rec'].startswith('[') else d['rec'].replace('el ', 'del ', 1)))
    # I
    sec(ctx, 'I')
    ficha(ctx, [('Auditado:', d['auditado']), ('Periodo auditado:', '01/01/[aaaa] al 31/12/[aaaa]'),
                ('Riesgo del auditado según su Metodología EBR (%s):' % D('6'), '[ ] Bajo   [ ] Medio   [ ] Alto'),
                ('Tipo de auditor:', '[ ] Área de auditoría o control interno (art. 44 RCG; solo si el Riesgo es bajo o medio)   [ ] Persona auditora externa independiente (art. 45 RCG)'),
                ('Nombre del auditor o responsable del área:', ''),
                ('Título y cédula profesional:', ''),
                ('Certificación UIF vigente (art. 34 Bis RCG):', '[No. y vigencia — obligatoria para el auditor externo, art. 45, inciso b)]')])
    p(ctx, '**Manifestación bajo protesta de decir verdad.** El auditor manifiesta que cumple los requisitos y conocimientos '
           'para realizar esta auditoría y, en particular:')
    for t in ['Área interna (art. 44 RCG): es independiente %s, conoce los procesos internos de los actos u operaciones '
              'del art. 17 de la Ley y acredita al menos un programa de capacitación anual.' % d['rec'].replace('el ', 'del ', 1) if not d['rec'].startswith('[') else 'Área interna (art. 44 RCG): es independiente de [PF: el profesional][PM: el Representante Encargado de Cumplimiento], conoce los procesos internos de los actos u operaciones del art. 17 de la Ley y acredita al menos un programa de capacitación anual.' if kit=='PROF' else 'Área interna (art. 44 RCG): es independiente de %s, conoce los procesos internos de los actos u operaciones del art. 17 de la Ley y acredita al menos un programa de capacitación anual.' % d['rec'],
              'Auditor externo (art. 45 RCG): cuenta con título y cédula en derecho, contaduría, finanzas, administración, '
              'informática o áreas afines y tres años de experiencia en prevención; tiene certificación vigente de la UIF; no ha '
              'sido sentenciado por delitos patrimoniales; no prestó servicios al auditado durante el periodo que generen '
              'conflicto de interés, y no aceptó el cargo de Representante Encargado de Cumplimiento del auditado durante el '
              'periodo ni en los dos años posteriores. Se mantiene independiente hasta emitir el dictamen.']:
        p(ctx, t, bullet=True)
    p(ctx, '**Conocimiento del auditado y Riesgos identificados por el auditor** (cuestionario inicial y análisis de Riesgo, '
           'art. 47, primer párrafo, RCG): [conclusiones].')
    # II
    sec(ctx, 'II')
    ficha(ctx, [('Actos u operaciones evaluados:', d['actos']),
                ('Obligaciones evaluadas:', 'Todas las de la sección VI'),
                ('Sucursales u oficinas evaluadas:', '[ ] No aplica   [ ] [Relación]')])
    # III
    sec(ctx, 'III')
    ficha(ctx, [('Actos u operaciones realizados en el periodo:', '[Número]'),
                ('Avisos e informes presentados en el periodo:', '[Número]'),
                ('Criterio de selección de la muestra:', '[Aleatorio, por Riesgo, por monto, por tipo de acto…]'),
                ('Actos u operaciones revisados:', '[Número] — [__]% del total del periodo'),
                ('Expedientes de clientes revisados:', '[Número] — [__]% del total')])
    # IV
    sec(ctx, 'IV')
    for t in ['Cuestionario inicial de conocimiento del auditado y análisis de su Riesgo; programa de trabajo derivado de ambos (art. 47, primer párrafo, RCG).',
              'Revisión de la Metodología EBR y de su congruencia con los mecanismos automatizados (art. 10 Septies 2 RCG).',
              'Revisión documental de la muestra de expedientes (%s) y de Avisos, informes en ceros y acuses (%s).' % (D('7'), D('12')),
              'Revisión de los Reportes de Operación Inusual y de los Avisos de 24 horas (%s).' % D('9'),
              'Verificación de la evidencia de capacitación y de selección de personal (%s, %s).' % (D('11'), D('11D')),
              'Verificación del resguardo documental (%s) y de los mecanismos automatizados.' % D('10'),
              'Entrevistas con %s y con el personal clave; ninguno de ellos puede actuar como auditor.' % d['rec']]:
        p(ctx, t, bullet=True)
    p(ctx, '**Asuntos relevantes durante la auditoría:** [describir].')
    # V
    sec(ctx, 'V')
    tabla(ctx, [('No.', 'Obligación (sección VI)', 'Hallazgo: deficiencia, omisión o área de oportunidad', 'Evidencia')] +
          [('H%d' % i, '', '', '') for i in range(1, 5)], [700, 2300, 4600, 1900])
    p(ctx, '**Aspectos destacables:** [describir buenas prácticas observadas].')
    # VI
    sec(ctx, 'VI')
    p(ctx, 'Cada obligación se califica con uno de los cinco resultados del art. 48 RCG: **C** Cumple; **CM** Cumple '
           'mayoritariamente (al menos una recomendación); **CP** Cumple parcialmente (al menos un hallazgo); **NC** No cumple '
           '(al menos un hallazgo); **NA** No aplica (verificado y congruente con el análisis de Riesgo).')
    filas = [('No.', 'Obligación evaluada', 'Documento', 'Resultado', 'Hallazgo o recomendación')]
    for i, (o, doc) in enumerate(obligaciones(d, D), 1):
        filas.append((str(i), o, doc, '[ ] C  [ ] CM  [ ] CP  [ ] NC  [ ] NA', ''))
    tabla(ctx, filas, [500, 3600, 1500, 1900, 2000])
    # VII
    sec(ctx, 'VII')
    p(ctx, 'Para cada hallazgo se señala la acción correctiva, el plazo, el responsable y la proyección económica de la multa '
           'que correspondería si no se atiende (art. 49 RCG). Referencia: 200 a 2,000 UMA por incumplir obligaciones del art. 18 '
           'o presentar Avisos extemporáneos; 10,000 a 65,000 UMA, o del 10% al 100% del valor del acto, por omitir Avisos u '
           'operar con efectivo prohibido (arts. 53 y 54 de la Ley). Las recomendaciones de mejora se listan por separado.')
    tabla(ctx, [('Hallazgo', 'Acción correctiva', 'Responsable', 'Plazo', 'Multa proyectada (UMA y MXN)')] +
          [('H%d' % i, '', '', '[dd/mm/aaaa]', '') for i in range(1, 5)], [1000, 3400, 1800, 1300, 2000])
    sub(ctx, 'VII.1 Recomendaciones de mejora')
    tabla(ctx, [('No.', 'Obligación', 'Recomendación', 'Plazo sugerido')] + [('R%d' % i, '', '', '') for i in range(1, 4)],
          [700, 2600, 4600, 1600])
    sub(ctx, 'VII.2 Seguimiento del dictamen del año anterior (art. 49, segundo párrafo, RCG)')
    tabla(ctx, [('Hallazgo anterior', 'Acción comprometida', 'Estado', 'Comentario del auditor')] +
          [('', '', '[ ] Atendido  [ ] En proceso  [ ] Pendiente', '')] * 3, [2000, 2900, 2400, 2200])
    # VIII
    sec(ctx, 'VIII')
    p(ctx, 'El dictamen se entrega a más tardar el último día hábil de marzo del año siguiente al auditado (art. 50 RCG) y se '
           'presenta %s para evaluar la eficacia de las medidas y dar seguimiento a las acciones correctivas (art. 42, '
           'segundo párrafo, RCG). Las acciones se atienden antes del inicio del siguiente periodo de auditoría o en los plazos '
           'señalados; los incumplimientos pueden subsanarse de manera espontánea en esos plazos (art. 50 RCG). El dictamen y '
           'su soporte se conservan al menos cinco años desde su entrega y se ponen a disposición del SAT cuando lo requiera '
           '(art. 51 RCG; %s).' % ('[PF: al titular][PM: al órgano de administración o al administrador único]' if kit == 'INM'
                                   else '[PF: al profesional][PM: al órgano de administración del despacho]' if kit == 'PROF'
                                   else 'al %s' % d['recibe'], D('10')))
    ficha(ctx, [('Fecha de entrega del dictamen:', '[dd/mm/aaaa]'),
                ('Auditor (nombre y firma):', ''),
                ('%s (recibió y tomó nota):' % d['recibe'], ''),
                ('Fecha de recepción:', '[dd/mm/aaaa]')])
    pie(ctx)
    return ctx

NOMBRES = {'PROF': 'profesionistas/D13-PROF_Informe_Auditoria_PLD_v2.docx', 'INM': 'inmobiliaria/D13_Informe_Auditoria_PLD_v5.docx',
           'NOT': 'notarial/D13-NOT_Informe_Auditoria_PLD_v3.docx', 'COR': 'corredores/D13-COR_Informe_Auditoria_PLD_v2.docx'}

if __name__ == '__main__':
    B.SRC, OUT = sys.argv[1], sys.argv[2]
    for k in ('PROF', 'INM', 'NOT', 'COR'):
        ctx = construir(k)
        doc = Document(os.path.join(B.SRC, 'PROF_pr12.docx'))
        instalar_estilos(doc)
        body = doc.element.body; sect = body.find(q('w:sectPr'))
        for el in list(body):
            if el is not sect: body.remove(el)
        espaciar_tras_tablas(ctx.body)
        for el in ctx.body: sect.addprevious(el)
        cp = doc.core_properties
        cp.title = 'D13%s Dictamen de Auditoría PLD/FT — %s' % (canon.KITS[k]['suf'], canon.KITS[k]['variante'])
        cp.author = 'Certus Consultores · Espinosa-Salazar Abogados Corporativos'
        cp.subject = 'Programa Certus PLD — edición homologada noviembre 2026'
        dest = os.path.join(OUT, NOMBRES[k]); os.makedirs(os.path.dirname(dest), exist_ok=True)
        from docxlib import numero_de_pagina; numero_de_pagina(doc)
        doc.save(dest); print(k, dest)
