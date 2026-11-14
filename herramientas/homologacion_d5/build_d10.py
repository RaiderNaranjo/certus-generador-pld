"""D10 — Política de Resguardo Documental, homologada para los cuatro kits.

Uso: python3 build_d10.py <dir_contenedor> <dir_salida>
"""
import os, sys
sys.path.insert(0, os.path.dirname(__file__))
import build_d5 as B
from build_d5 import Ctx, p, tabla, portada, pie, escribir_par, Par
from docxlib import instalar_estilos, q, espaciar_tras_tablas
from docx import Document
import canon

SECCIONES = [
    ('I', 'Fundamento y Plazos de Conservación (art. 18, fr. IV, de la Ley; art. 37 Bis, fr. VII, RCG)', 'Fundamento y plazos'),
    ('II', 'Documentos Sujetos a Resguardo', 'Documentos'),
    ('III', 'Sistema de Organización (art. 12 RCG)', 'Organización'),
    ('IV', 'Control de Vencimientos', 'Vencimientos'),
    ('V', 'Procedimiento de Destrucción Documental', 'Destrucción'),
    ('VI', 'Atención de Requerimientos y Visitas de la Autoridad (arts. 25, 34 y 36 de la Ley; art. 9 RCG)', 'Requerimientos'),
]

def sec(ctx, rom):
    t = [s[1] for s in SECCIONES if s[0] == rom][0]
    ctx.body.append(escribir_par(Par.from_text('Sección %s: %s' % (rom, t)), 'CertusCapitulo'))

def sub(ctx, t):
    ctx.body.append(escribir_par(Par.from_text(t), 'CertusSeccion'))

def indice():
    return 'Índice: ' + ' │ '.join('%s. %s' % (r, c) for r, _, c in SECCIONES)

S = {
 'INM': dict(suf='', suj='la empresa', Suj='La empresa', resp='[PF: el titular][PM: el Representante Encargado de Cumplimiento]',
    Resp='[PF: Titular][PM: Representante Encargado de Cumplimiento]',
    d2='[PF: Constancia de Asunción Personal de Obligaciones del titular (D2-PF)][PM: Acta de Designación del Representante Encargado de Cumplimiento (D2)]',
    actos='Contratos celebrados con los clientes (compraventa, intermediación, preventa o aportación a desarrollos y arrendamiento) y su documentación de pago',
    extra=[], archivo='por año y por fracción (V, V Bis y XV)',
    acceso='[PF: el titular y el personal que éste autorice por escrito][PM: el Representante Encargado de Cumplimiento y el personal que autorice por escrito]',
    sin_perjuicio=None,
    destruye='[PF: El titular autoriza la destrucción mediante constancia firmada.][PM: El órgano de administración autoriza la destrucción mediante acuerdo o minuta, a propuesta del Representante Encargado de Cumplimiento.]',
    notifica='[PF: al titular][PM: al Representante Encargado de Cumplimiento y al órgano de administración]',
    req_extra=[]),
 'NOT': dict(suf='-NOT', suj='la Notaría', Suj='La Notaría', resp='el Notario Titular', Resp='Notario Titular',
    d2='Constancia de Asunción Personal de Obligaciones del Notario Titular (D2-NOT)',
    actos='Instrumentos notariales formalizados (transmisiones, poderes irrevocables, sociedades, fideicomisos y mutuos) y la identificación de la forma de pago (art. 33 de la Ley)',
    extra=['Datos de la persona que solicitó materialmente la protocolización (art. 13 RCG; D7-NOT, Sección A.4).'],
    archivo='por año y por inciso (a) inmuebles, b) poderes, c) sociedades, d) fideicomisos, e) mutuos)',
    acceso='el Notario Titular y, en su caso, el personal que designe por escrito',
    sin_perjuicio='Esta Política se aplica sin perjuicio de las obligaciones de conservación del protocolo y del apéndice que impone la ley del notariado aplicable.',
    destruye='El Notario Titular autoriza la destrucción mediante constancia firmada; no hay órgano colegiado que deba aprobarla.',
    notifica='al Notario Titular',
    req_extra=['Si el requerimiento proviene de la autoridad estatal que supervisa la fe pública, se aplica el mismo protocolo.']),
 'COR': dict(suf='-COR', suj='la Correduría', Suj='La Correduría', resp='el Corredor Público Titular', Resp='Corredor Público Titular',
    d2='Constancia de Asunción Personal de Obligaciones del Corredor Público Titular (D2-COR)',
    actos='Pólizas y actas formalizadas, sus asientos en el libro de registro y los avalúos emitidos con fe pública con su soporte técnico, así como la identificación de la forma de pago (art. 33 de la Ley)',
    extra=['Datos de la persona que solicitó materialmente la formalización (art. 13 RCG; D7-COR, Sección A.4).'],
    archivo='por año y por inciso (a) avalúos, b) sociedades, c) fideicomisos, d) mutuos y créditos)',
    acceso='el Corredor Público Titular y, en su caso, el personal que designe por escrito',
    sin_perjuicio='Esta Política se aplica sin perjuicio de las obligaciones de conservación del archivo y del libro de registro que impone la Ley Federal de Correduría Pública.',
    destruye='El Corredor Público Titular autoriza la destrucción mediante constancia firmada; no hay órgano colegiado que deba aprobarla.',
    notifica='al Corredor Público Titular',
    req_extra=['Si el requerimiento proviene de la Secretaría de Economía, en ejercicio de su supervisión de la correduría pública, se aplica el mismo protocolo.']),
 'PROF': dict(suf='-PROF', suj='[PF: el profesional][PM: el despacho]', Suj='[PF: El profesional][PM: El despacho]',
    resp='[PF: el profesional][PM: el Representante Encargado de Cumplimiento]',
    Resp='[PF: Profesional Titular][PM: Representante Encargado de Cumplimiento]',
    d2='[PF: Constancia de Asunción Personal de Obligaciones (D2-PROF-PF)][PM: Acta de Designación del Representante Encargado de Cumplimiento (D2-PROF-PM)]',
    actos='Instrumentos, contratos o documentación que evidencien el servicio prestado en cada inciso (a) a e)) y, en su caso, las operaciones financieras realizadas en nombre y representación del cliente',
    extra=[], archivo='por año y por inciso del servicio (a) a e))',
    acceso='[PF: el profesional y el personal que autorice por escrito][PM: el Representante Encargado de Cumplimiento y el personal que autorice por escrito]',
    sin_perjuicio='La entrega de información a la autoridad no transgrede el secreto profesional (art. 22 de la Ley), sin perjuicio del criterio sobre secreto profesional y garantía de defensa del D1-PROF, sección IV.',
    destruye='[PF: El profesional autoriza la destrucción mediante constancia firmada.][PM: El órgano de administración del despacho autoriza la destrucción mediante acuerdo, a propuesta del Representante Encargado de Cumplimiento.]',
    notifica='[PF: al profesional][PM: al Representante Encargado de Cumplimiento y al órgano de administración del despacho]',
    req_extra=['Se atiende el requerimiento con respeto al secreto profesional y la garantía de defensa aplicables a la fracción XI (D1-PROF, sección IV).']),
}

def construir(kit):
    d = S[kit]; ctx = Ctx(kit); s = d['suf']
    D = lambda n: 'D%s%s' % (n, s)
    portada(ctx, codigo='D10', titulo='Política de Resguardo Documental', indice=indice())
    # I
    sec(ctx, 'I')
    p(ctx, '%s custodia, protege y resguarda, evitando su destrucción u ocultamiento, la información y documentación que '
           'sirve de soporte a la Actividad Vulnerable: los registros que permiten reconstruir cada acto u operación, la '
           'correspondencia que las partes se hayan compartido, los resultados de los análisis previos y la identificación de '
           'los clientes. Se conserva de manera física o electrónica, en el domicilio registrado ante la Secretaría (art. 18, '
           'fr. IV, de la Ley). Esta Política desarrolla el %s, capítulo 11.' % (d['Suj'], D('5')))
    tabla(ctx, [('Información', 'Plazo mínimo', 'Fundamento'),
                ('Expedientes de identificación, soporte de cada acto u operación, Avisos y acuses', 'Diez años desde la fecha del acto u operación', 'art. 18, fr. IV, de la Ley'),
                ('Metodología EBR, sus revisiones y resultados', 'Diez años', 'art. 10 Septies 3 RCG'),
                ('Historial de cambios de Grado de Riesgo y de Perfil Transaccional', 'Diez años', 'art. 41, fr. IV, RCG'),
                ('Evidencia de capacitación (programas, listas, evaluaciones y constancias)', 'Diez años', 'art. 39 Bis 1 RCG'),
                ('Dictamen de auditoría y su soporte', 'Cinco años desde su entrega', 'art. 51 RCG')],
          [4600, 2700, 2200])
    p(ctx, 'Si respecto de la información se promueve un recurso o juicio, el plazo se interrumpe en la fecha de su '
           'presentación y se reinicia cuando la resolución definitiva quede firme (art. 18, fr. IV, de la Ley).')
    if d['sin_perjuicio']:
        p(ctx, d['sin_perjuicio'])
    # II
    sec(ctx, 'II')
    sub(ctx, 'II.1 Expedientes de clientes')
    for t in ['Expediente único de identificación de cada cliente (%s), con identificaciones, comprobantes, constancias y cotejos.' % D('7'),
              '%s.' % d['actos']] + d['extra'] + [
              'Soporte del origen y destino de los recursos o bienes.',
              'Documentación de identificación del Beneficiario Controlador y del procedimiento seguido para identificarlo (art. 23 Quinquies RCG).',
              'Fichas de evaluación EBR firmadas (%s, sección XI) y, en su caso, el Cuestionario de Identificación Reforzada.' % D('6'),
              "Constancias de consulta de listas y de PEP, y, en su caso, la aprobación o constancia del %s, sección VIII." % D('8')]:
        p(ctx, t, bullet=True)
    sub(ctx, 'II.2 Documentos del Programa Certus PLD')
    for t in ['Manual de Políticas Internas (%s) y todas sus versiones anteriores, con su control de versiones (Anexo D).' % D('5'),
              d['d2'] + '.',
              'Acuse de alta y registro en el padrón de quienes realizan Actividades Vulnerables y sus actualizaciones (%s).' % D('3'),
              'Avisos, Informes y sus acuses (%s).' % D('12'),
              'Reportes de Operación Inusual y, en su caso, acuses de los Avisos de 24 horas (%s).' % D('9'),
              "Registro de PEP's (%s, sección XI)." % D('8'),
              'Evidencia de capacitación (%s, %s y %s) y declaraciones de selección de personal (%s; art. 39 Bis 2 RCG).' % (D('11'), D('11B'), D('11C'), D('11D')),
              'Dictámenes de auditoría (%s) y su soporte.' % D('13')]:
        p(ctx, t, bullet=True)
    # III
    sec(ctx, 'III')
    p(ctx, 'Todos los datos y documentos de un cliente se integran en un expediente único, físico o electrónico, dentro de un '
           'mismo archivo (art. 12 RCG).')
    sub(ctx, 'III.1 Archivo físico')
    for t in ['Archiveros en el domicilio registrado ante la Secretaría, organizados %s.' % d['archivo'],
              'Una carpeta por cliente, identificada con el número de expediente (%s).' % D('7'),
              'Una carpeta de Avisos e Informes organizada por año y mes.',
              'Acceso restringido: %s.' % d['acceso']]:
        p(ctx, t, bullet=True)
    sub(ctx, 'III.2 Archivo electrónico')
    for t in ['Repositorio cifrado (servidor propio o nube) con respaldo automático diario en una ubicación alterna.',
              'Nomenclatura uniforme: [Documento]-[Año-Mes]-[Cliente o concepto], por ejemplo D7-2027-03-Gómez.',
              'Usuario y contraseña individuales para cada persona autorizada, y bitácora de accesos y modificaciones.',
              'Los registros deben permitir consultar y reconstruir cada acto u operación y ser verificables ante la autoridad (art. 41, fr. I, RCG; %s, capítulo 19).' % D('5')]:
        p(ctx, t, bullet=True)
    # IV
    sec(ctx, 'IV')
    filas = [('Año del acto u operación', 'Se conserva al menos hasta', 'Verifica', 'Estado')]
    for a in range(2025, 2032):
        filas.append((str(a), '%d (según la fecha de cada acto u operación más diez años)' % (a + 10), d['Resp'], '[ ] Vigente   [ ] Venció'))
    tabla(ctx, filas, [1700, 3800, 2300, 1700])
    p(ctx, 'La revisión de vencimientos se hace una vez al año, junto con la verificación anual de expedientes (art. 21 RCG).')
    # V
    sec(ctx, 'V')
    for t in ['Verificar que el plazo de conservación haya vencido y que no exista investigación, auditoría, requerimiento, visita, recurso o juicio pendiente relacionado con los documentos.',
              '%s elabora un inventario de los documentos por destruir, con sus fechas originales.' % d['resp'].replace('[PF: el', '[PF: El').replace('[PM: el', '[PM: El').replace('el ', 'El ', 1 if not d['resp'].startswith('[') else 0),
              d['destruye'],
              'Los documentos físicos se destruyen por trituración y los electrónicos por borrado seguro, con constancia del proveedor o de quien la ejecute.',
              'El acta de destrucción y el inventario se conservan cinco años adicionales.']:
        p(ctx, t, bullet=True)
    # VI
    sec(ctx, 'VI')
    for t in ['Notificar de inmediato %s y al asesor legal (Certus Consultores · Espinosa-Salazar Abogados Corporativos │ contacto@certusconsultores.com.mx │ 999 163 1363 · contacto@espinosasalazarabogados.com │ 999 575 3129).' % d['notifica'],
              'Atender el requerimiento en el plazo que fije la autoridad; los requerimientos sobre el alta y registro se atienden en diez días hábiles, prorrogables hasta por cinco a solicitud del interesado (art. 9 RCG). La Secretaría puede verificar los actos u operaciones de los cinco años anteriores al inicio de la visita (art. 36 de la Ley).',
              'Proporcionar solo la información relacionada con las Actividades Vulnerables (art. 34, segundo párrafo, de la Ley) y conservar copia y acuse de lo entregado.',
              'No destruir, alterar ni ocultar documento alguno, ni alertar al cliente o a terceros sobre el requerimiento (art. 31 RCG; %s, sección V).' % D('9')] + d['req_extra']:
        p(ctx, t, bullet=True)
    p(ctx, '')
    tabla(ctx, [('**Aprobó la Política (nombre y cargo):**', ''), ('**Fecha de adopción:**', '[dd/mm/aaaa]'), ('**Firma:**', '')],
          [3600, 5900], encabezado=False)
    pie(ctx)
    return ctx

NOMBRES = {'PROF': 'profesionistas/D10-PROF_Politica_Resguardo_Documental_v2.docx', 'INM': 'inmobiliaria/D10_Politica_Resguardo_Documental_v5.docx',
           'NOT': 'notarial/D10-NOT_Politica_Resguardo_Documental_v3.docx', 'COR': 'corredores/D10-COR_Politica_Resguardo_Documental_v2.docx'}

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
        cp.title = 'D10%s Política de Resguardo Documental — %s' % (canon.KITS[k]['suf'], canon.KITS[k]['variante'])
        cp.author = 'Certus Consultores · Espinosa-Salazar Abogados Corporativos'
        cp.subject = 'Programa Certus PLD — edición homologada noviembre 2026'
        dest = os.path.join(OUT, NOMBRES[k]); os.makedirs(os.path.dirname(dest), exist_ok=True)
        from docxlib import numero_de_pagina; numero_de_pagina(doc)
        doc.save(dest); print(k, dest)
