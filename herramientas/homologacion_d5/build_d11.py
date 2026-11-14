"""D11 — Programa Anual de Capacitación, homologado para los cuatro kits.

Uso: python3 build_d11.py <dir_contenedor> <dir_salida>
"""
import os, sys
sys.path.insert(0, os.path.dirname(__file__))
import build_d5 as B
from build_d5 import Ctx, p, tabla, portada, pie, escribir_par, Par
from docxlib import instalar_estilos, q, espaciar_tras_tablas
from docx import Document
import canon

SECCIONES = [
    ('I', 'Objetivo y Fundamento (art. 18, fr. IX, de la Ley; arts. 39 Bis a 39 Bis 2 RCG)', 'Objetivo'),
    ('II', 'Grupos de Capacitación', 'Grupos'),
    ('III', 'Módulos de Capacitación Certus PLD', 'Módulos'),
    ('IV', 'Requisitos de Quien Imparte la Capacitación (art. 39 Bis, fr. III, RCG)', 'Capacitador'),
    ('V', 'Calendario Anual', 'Calendario'),
    ('VI', 'Capacitación de Nuevo Ingreso (art. 39 Bis 1, tercer párrafo, RCG)', 'Nuevo ingreso'),
    ('VII', 'Registro de Asistencia', 'Asistencia'),
    ('VIII', 'Evaluación y Constancias (art. 39 Bis 1, segundo párrafo, RCG)', 'Evaluación'),
    ('IX', 'Selección de Personal (art. 39 Bis 2 RCG)', 'Selección'),
    ('X', 'Resguardo de la Evidencia (art. 39 Bis 1, primer párrafo, RCG)', 'Resguardo'),
]

def sec(ctx, rom):
    t = [s[1] for s in SECCIONES if s[0] == rom][0]
    ctx.body.append(escribir_par(Par.from_text('Sección %s: %s' % (rom, t)), 'CertusCapitulo'))

def indice():
    return 'Índice: ' + ' │ '.join('%s. %s' % (r, c) for r, _, c in SECCIONES)

CERTUS = 'Certus Consultores · Espinosa-Salazar Abogados Corporativos'
ACRED = 'Instructor acreditado (art. 39 Bis, fr. III, RCG)'

S = {
 'INM': dict(suf='', suj='la empresa', Suj='La empresa', frac='fracciones V, V Bis y XV',
    resp='[PF: el titular][PM: el Representante Encargado de Cumplimiento]',
    Resp='[PF: Titular][PM: Representante Encargado de Cumplimiento]',
    grupos=[('[PF: G1 – Colaboradores directivos][PM: G1 – Alta dirección]', '[PF: Colaboradores con funciones de dirección, en su caso][PM: Consejo de administración o administrador único y directivos]', 'Mod. 1, 2, 6 y 7', 'Presencial', 'Anual', 'Mayo', CERTUS),
            ('[PF: G2 – Titular][PM: G2 – Representante Encargado de Cumplimiento]', '[PF: Persona física titular, que cumple personal y directamente (art. 20 de la Ley)][PM: Representante Encargado de Cumplimiento y su suplente]', 'Todos (1 a 7)', 'Presencial y virtual', 'Semestral', 'Enero y julio', CERTUS),
            ('G3 – Área comercial', 'Vendedores, asesores y ejecutivos de preventa', 'Mod. 1, 2, 3, 4 y 7', 'Presencial', 'Anual y al ingresar', 'Febrero', ACRED),
            ('G4 – Arrendamiento', 'Administradores de inmuebles y gestores de contratos', 'Mod. 1, 2, 3, 5 y 7', 'Presencial', 'Anual y al ingresar', 'Marzo', ACRED),
            ('G5 – Administración', 'Contabilidad, tesorería, jurídico y, en su caso, auditoría interna', 'Mod. 1, 2, 6 y 7', 'Virtual', 'Anual', 'Abril', ACRED)],
    nota_grupos=None,
    m1='Ley, Reglamento y RCG (Acuerdo 115/2026); fracciones V, V Bis y XV y sus umbrales; UMA vigente; contenido del Manual de Políticas Internas (D5)',
    m2='Tipologías y señales de alerta del sector inmobiliario (D9, sección II); casos de estudio',
    m5=('Desarrollos y arrendamiento: fracciones V Bis y XV', 'Recepción de recursos para desarrollos (Aviso mensual por desarrollo, art. 24 Bis 1 RCG), arrendamiento (umbral mensual y copropietarios, art. 24 Bis 6 RCG) y prohibición de efectivo (art. 32, fr. I y VII, de la Ley)')),
 'NOT': dict(suf='-NOT', suj='la Notaría', Suj='La Notaría', frac='fracción XII, Apartado A',
    resp='el Notario Titular', Resp='Notario Titular',
    grupos=[('G1 – Notario Titular', 'Notario Titular, que cumple personal y directamente (art. 20 de la Ley)', 'Todos (1 a 7)', 'Presencial y virtual', 'Semestral', 'Enero y julio', CERTUS),
            ('G2 – Encargado Operativo', 'Encargado Operativo interno, si fue designado (D2-NOT, sección III)', 'Todos (1 a 7)', 'Presencial y virtual', 'Semestral', 'Enero y julio', CERTUS),
            ('G3 – Atención al público', 'Oficiales, pasantes y gestores que atienden e identifican a los comparecientes', 'Mod. 1, 2, 3, 4, 5 y 7', 'Presencial', 'Anual y al ingresar', 'Febrero', ACRED),
            ('G4 – Administración y archivo', 'Contabilidad, archivo, área jurídica interna y, en su caso, auditoría interna', 'Mod. 1, 2, 6 y 7', 'Virtual', 'Anual', 'Abril', ACRED)],
    nota_grupos='No hay grupos de alta dirección ni de área comercial: el Notario Titular responde personalmente del cumplimiento (art. 20 de la Ley) y recibe la capacitación completa.',
    m1='Ley, Reglamento y RCG (Acuerdo 115/2026); fracción XII, Apartado A (incisos a) a e)); régimen sancionador y aviso a la autoridad de la fe pública (art. 58 de la Ley); UMA vigente; contenido del Manual (D5-NOT)',
    m2='Tipologías y señales de alerta por inciso del acto notarial (D9-NOT, sección II); casos de estudio',
    m5=('Actos siempre objeto de Aviso: poderes, sociedades y mutuos', 'Por qué los incisos b), c) y e) no tienen umbral; umbrales de los incisos a) (8,000 UMA) y d) (4,000 UMA); forma de pago (art. 33 de la Ley); expediente por tipo de acto (D7-NOT, Sección I)')),
 'COR': dict(suf='-COR', suj='la Correduría', Suj='La Correduría', frac='fracción XII, Apartado B',
    resp='el Corredor Público Titular', Resp='Corredor Público Titular',
    grupos=[('G1 – Corredor Público Titular', 'Corredor Público Titular, que cumple personal y directamente (art. 20 de la Ley)', 'Todos (1 a 7)', 'Presencial y virtual', 'Semestral', 'Enero y julio', CERTUS),
            ('G2 – Encargado Operativo', 'Encargado Operativo interno, si fue designado (D2-COR, sección III)', 'Todos (1 a 7)', 'Presencial y virtual', 'Semestral', 'Enero y julio', CERTUS),
            ('G3 – Atención al público', 'Auxiliares y pasantes que reciben e identifican a los comparecientes y a los solicitantes de avalúos', 'Mod. 1, 2, 3, 4, 5 y 7', 'Presencial', 'Anual y al ingresar', 'Febrero', ACRED),
            ('G4 – Administración, archivo y avalúos', 'Contabilidad, archivo y libro de registro, área técnica de avalúos y, en su caso, auditoría interna', 'Mod. 1, 2, 6 y 7', 'Virtual', 'Anual', 'Abril', ACRED)],
    nota_grupos='No hay grupos de alta dirección ni de área comercial: el Corredor Público Titular responde personalmente del cumplimiento (art. 20 de la Ley) y recibe la capacitación completa.',
    m1='Ley, Reglamento y RCG (Acuerdo 115/2026); fracción XII, Apartado B (incisos a) a d)); régimen sancionador ante la Secretaría de Economía (art. 57 de la Ley); UMA vigente; contenido del Manual (D5-COR)',
    m2='Tipologías y señales de alerta por inciso del acto ante el Corredor (D9-COR, sección II); casos de estudio',
    m5=('Sociedades, fideicomisos y créditos; avalúos con fe pública', 'Por qué los incisos b), c) y d) se avisan siempre; umbral del inciso a) (8,025 UMA); efectivo en acciones (art. 32, fr. VI, de la Ley) y forma de pago (art. 33); expediente por tipo de acto (D7-COR, Sección I)')),
 'PROF': dict(suf='-PROF', suj='[PF: el profesional][PM: el despacho]', Suj='[PF: El profesional][PM: El despacho]', frac='fracción XI',
    resp='[PF: el profesional][PM: el Representante Encargado de Cumplimiento]',
    Resp='[PF: Profesional Titular][PM: Representante Encargado de Cumplimiento]',
    grupos=[('[PF: G1 – Colaboradores directivos][PM: G1 – Dirección del despacho]', '[PF: Colaboradores con funciones de dirección, en su caso][PM: Órgano de administración y socios directores]', 'Mod. 1, 2, 6 y 7', 'Presencial', 'Anual', 'Mayo', CERTUS),
            ('[PF: G2 – Profesional Titular][PM: G2 – Representante Encargado de Cumplimiento]', '[PF: Profesional independiente, que cumple personal y directamente (art. 20 de la Ley)][PM: Representante Encargado de Cumplimiento del despacho y su suplente]', 'Todos (1 a 7)', 'Presencial y virtual', 'Semestral', 'Enero y julio', CERTUS),
            ('[PF: G3 – Colaboradores profesionales][PM: G3 – Socios y asociados]', '[PF: Colaboradores que prestan servicios de los incisos a) a e) bajo supervisión del profesional][PM: Socios y asociados que prestan servicios de los incisos a) a e)]', 'Mod. 1, 2, 3, 4, 5 y 7', 'Presencial', 'Anual y al ingresar', 'Febrero', ACRED),
            ('G4 – Personal de apoyo', 'Asistentes, pasantes y personal administrativo con contacto con clientes', 'Mod. 1, 2, 6 y 7', 'Virtual', 'Anual y al ingresar', 'Abril', ACRED)],
    nota_grupos='La capacitación se organiza por el grado de participación en los servicios de los incisos a) a e), no por funciones de venta.',
    m1='Ley, Reglamento y RCG (Acuerdo 115/2026); fracción XI (incisos a) a e)), sin umbral monetario; UMA vigente; contenido del Manual (D5-PROF)',
    m2='Tipologías y señales de alerta por inciso del servicio (D9-PROF, sección II); casos de estudio',
    m5=('Preparación frente a representación: cuándo procede el Aviso', 'D1-PROF, sección 2.2; Aviso mensual solo por operaciones financieras en nombre y representación del cliente; Aviso de 24 horas aun en preparación; secreto profesional y garantía de defensa (D1-PROF, sección IV)')),
}

def construir(kit):
    d = S[kit]; ctx = Ctx(kit); s = d['suf']
    D = lambda n: 'D%s%s' % (n, s)
    portada(ctx, codigo='D11', titulo='Programa Anual de Capacitación', indice=indice())
    # I
    sec(ctx, 'I')
    p(ctx, 'Este Programa desarrolla el %s, capítulo 12. Se dirige, según corresponda, a quienes integran el órgano de '
           'administración o al administrador único, a los directivos y funcionarios, a quien cumple las funciones de '
           'Representante Encargado de Cumplimiento y a todo el personal que atiende al público, participa en la '
           'identificación o el conocimiento del cliente, en el envío de Avisos o en funciones de auditoría (art. 39 Bis, '
           'primer párrafo, RCG).' % D('5'))
    p(ctx, 'Contenido mínimo, al menos una vez al año: (a) la Ley, su Reglamento, las RCG y las Resoluciones de formatos '
           'oficiales; (b) el Manual de Políticas Internas; (c) los actos u operaciones del art. 17 de la Ley que realiza %s '
           '(%s); (d) los Riesgos a que está expuesto, coherentes con su Metodología EBR (%s); y técnicas, métodos y '
           'tendencias para prevenir y detectar operaciones del art. 400 Bis del Código Penal Federal (art. 39 Bis, fr. I y '
           'II, RCG).' % (d['suj'], d['frac'], D('6')))
    p(ctx, 'Vigencia: el primer periodo anual de capacitación comprende del 1 de enero al 31 de diciembre de 2027 '
           '(Transitorio Séptimo del Acuerdo 115/2026), y los procedimientos de selección de personal se aplican a las '
           'contrataciones realizadas a partir del 1 de marzo de 2027 (Transitorio Sexto).')
    # II
    sec(ctx, 'II')
    tabla(ctx, [('Grupo', 'Integrantes', 'Módulos', 'Modalidad', 'Frecuencia')] + [g[:5] for g in d['grupos']],
          [2100, 3200, 1500, 1300, 1400])
    if d['nota_grupos']:
        p(ctx, d['nota_grupos'])
    # III
    sec(ctx, 'III')
    tabla(ctx, [('Módulo', 'Tema', 'Duración', 'Contenido principal'),
                ('Mod. 1', 'Marco normativo PLD/FT', '60 min', d['m1']),
                ('Mod. 2', 'Señales de alerta del sector', '45 min', d['m2']),
                ('Mod. 3', 'Identificación y conocimiento del cliente; Beneficiario Controlador', '60 min', 'Expediente único (%s), cotejo, origen y destino de los recursos, Perfil Transaccional y orden de prelación del Beneficiario Controlador (art. 23 Quinquies RCG)' % D('7')),
                ('Mod. 4', "PEP's y debida diligencia reforzada", '45 min', 'Definición y personas asimiladas (art. 23 Quáter RCG), roles a evaluar, procedimiento de identificación, Consulta PEP 2.0, medidas reforzadas y aprobación o constancia del art. 23 Ter 5 RCG (%s)' % D('8')),
                ('Mod. 5', d['m5'][0], '45 min', d['m5'][1]),
                ('Mod. 6', 'Avisos, Informes y operaciones inusuales', '45 min', 'Aviso mensual e informe en ceros (%s), Aviso de 24 horas por sospecha o indicios (%s), coincidencias en listas y confidencialidad' % (D('12'), D('9'))),
                ('Mod. 7', 'Reforma RCG 2026: EBR, Grado de Riesgo, mecanismos automatizados y auditoría', '45 min', 'Metodología EBR y escala de 7 a 21 (%s), reevaluación semestral, mecanismos automatizados (art. 41 RCG) y auditoría anual (%s)' % (D('6'), D('13')))],
          [900, 2300, 1050, 5250])
    p(ctx, 'El Módulo 7 se imparte a todos los grupos durante el primer periodo de capacitación (2027). Los casos prácticos '
           '(%s) y la evaluación (%s) se aplican al cierre de cada sesión.' % (D('11B'), D('11C')))
    # IV
    sec(ctx, 'IV')
    p(ctx, 'Quien imparta la capacitación debe contar y acreditar experiencia de por lo menos cinco años en prevención de '
           'lavado de dinero y delitos relacionados (art. 39 Bis, fr. III, RCG). Si %s no reúne ese requisito, puede '
           'coordinar las sesiones, pero no impartirlas.' % d['resp'])
    tabla(ctx, [('Capacitador', 'Experiencia acreditada', 'Credencial', 'Evidencia conservada en'),
                ('[Nombre]', '[Años]', '[Certificación UIF (art. 34 Bis RCG) o cédula, si aplica]', '[Ubicación]')],
          [2400, 1800, 2900, 2400])
    # V
    sec(ctx, 'V')
    filas = [('Mes', 'Grupo', 'Módulos', 'Fecha', 'Lugar', 'Instructor')]
    for g in d['grupos']:
        filas.append((g[5], g[0], g[2], '[dd/mm/aaaa]', '[Lugar o virtual]', g[6]))
    filas.append(('Octubre', 'Todos los grupos', 'Recordatorio y actualización normativa', '[dd/mm/aaaa]', '[Lugar o virtual]', ACRED))
    filas.append(('Diciembre', d['Resp'], 'Evaluación anual del Programa y ajustes para el año siguiente', '[dd/mm/aaaa]', '[Lugar]', d['Resp']))
    tabla(ctx, filas, [1100, 2300, 1600, 1200, 1300, 2000])
    # VI
    sec(ctx, 'VI')
    p(ctx, 'Quien vaya a operar o laborar en áreas de atención al público o de administración de recursos recibe la '
           'capacitación de manera previa o simultánea a su ingreso o al inicio de sus actividades en esas áreas.')
    tabla(ctx, [('Nombre', 'Puesto o área', 'Fecha de ingreso', 'Fecha de capacitación de inducción'),
                ('', '', '', '')], [2800, 2200, 1900, 2600])
    # VII
    sec(ctx, 'VII')
    tabla(ctx, [('Nombre del participante', 'Cargo', 'Grupo', 'Fecha', 'Módulos', 'Calificación', 'Firma')] + [('', '', '', '', '', '', '')] * 4,
          [2200, 1300, 900, 1100, 1300, 1200, 1500])
    # VIII
    sec(ctx, 'VIII')
    p(ctx, 'Al terminar la capacitación, cada participante presenta la Evaluación de Comprensión (%s). La calificación '
           'mínima aprobatoria es del 70%% de aciertos. La constancia se expide solo a quien aprueba.' % D('11C'))
    p(ctx, 'Las medidas ante resultados no satisfactorios son las del %s, sección 12.1: tras la primera reprobación, '
           'capacitación de refuerzo dentro de los quince días hábiles siguientes y nueva evaluación; tras la segunda, retiro '
           'temporal de las funciones de atención al público, identificación de clientes, envío de Avisos o administración de '
           'recursos hasta aprobar, con supervisión directa mientras tanto.' % D('5'))
    # IX
    sec(ctx, 'IX')
    p(ctx, '%s aplica procedimientos de selección que garantizan que su personal tenga la calidad técnica, la experiencia '
           'necesaria y la honorabilidad para desempeñar sus funciones. De cada funcionario o empleado obtiene una '
           'declaración firmada en la que conste, en su caso, su experiencia previa en otros sectores sujetos a la Ley, y que '
           'no ha sido sentenciado por delitos patrimoniales ni inhabilitado para ejercer el comercio, el servicio público o '
           'funciones en el sistema financiero (%s; %s, capítulo 20).' % (d['Suj'], D('11D'), D('5')))
    # X
    sec(ctx, 'X')
    p(ctx, 'Los programas, materiales, listas de asistencia, evaluaciones y constancias se conservan, de manera física o '
           'electrónica, por al menos diez años (%s).' % D('10'))
    p(ctx, '')
    tabla(ctx, [('**Aprobó el Programa (nombre y cargo):**', ''), ('**Fecha de aprobación:**', '[dd/mm/aaaa]'), ('**Firma:**', '')],
          [3600, 5900], encabezado=False)
    pie(ctx)
    return ctx

NOMBRES = {'PROF': 'profesionistas/D11-PROF_Programa_Anual_Capacitacion_v2.docx', 'INM': 'inmobiliaria/D11_Programa_Anual_Capacitacion_v6.docx',
           'NOT': 'notarial/D11-NOT_Programa_Anual_Capacitacion_v3.docx', 'COR': 'corredores/D11-COR_Programa_Anual_Capacitacion_v2.docx'}

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
        cp.title = 'D11%s Programa Anual de Capacitación — %s' % (canon.KITS[k]['suf'], canon.KITS[k]['variante'])
        cp.author = 'Certus Consultores · Espinosa-Salazar Abogados Corporativos'
        cp.subject = 'Programa Certus PLD — edición homologada noviembre 2026'
        dest = os.path.join(OUT, NOMBRES[k]); os.makedirs(os.path.dirname(dest), exist_ok=True)
        from docxlib import numero_de_pagina; numero_de_pagina(doc)
        doc.save(dest); print(k, dest)
