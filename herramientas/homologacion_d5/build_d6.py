"""D6 — Metodología EBR homologada para los cuatro kits (4 elementos, 7 indicadores, suma 7–21).

Uso: python3 build_d6.py <dir_contenedor> <dir_salida>   (contenedor = carpeta con PROF_pr12.docx)
"""
import os, sys
sys.path.insert(0, os.path.dirname(__file__))
import build_d5 as B
from build_d5 import Ctx, p, seccion, tabla, portada, pie, escribir_par, Par, aplicar_globales
from docxlib import instalar_estilos, q
from docx import Document
import canon

SECCIONES = [
    ('I', 'Objetivo y Marco Normativo (art. 18, fr. VII, de la Ley; Capítulos II Quáter y III Bis RCG)', 'Objetivo y marco normativo'),
    ('II', 'Fuente y Periodo de los Datos (art. 10 Septies 2 RCG)', 'Fuente y periodo de los datos'),
    ('III', 'Elementos e Indicadores de Riesgo (art. 10 Septies 1, fr. I, RCG)', 'Elementos e indicadores'),
    ('IV', 'Indicadores de Financiamiento al Terrorismo y Lavado de Dinero (art. 10 Septies 1, último párrafo, RCG)', 'Indicadores FT y LD'),
    ('V', 'Mitigantes (art. 10 Septies 1, fr. III, RCG)', 'Mitigantes'),
    ('VI', 'Cálculo y Clasificación del Grado de Riesgo (arts. 23 Bis y 23 Bis 2 RCG)', 'Cálculo y clasificación'),
    ('VII', 'Clasificación Obligatoria de Riesgo Alto (art. 23 Bis 4 RCG)', 'Riesgo alto obligatorio'),
    ('VIII', 'Factores Adicionales para PEP Nacionales (art. 23 Bis 3 RCG)', 'PEP nacionales'),
    ('IX', 'Reevaluación y Revisión de la Metodología (arts. 10 Septies 2, 10 Septies 3 y 23 Bis 1 RCG)', 'Reevaluación y revisión'),
    ('X', 'Resguardo y Disponibilidad ante la Autoridad (arts. 10 Septies 3 y 41, fr. IV, RCG)', 'Resguardo'),
    ('XI', 'Ficha de Evaluación EBR por Cliente', 'Ficha de evaluación'),
]

def sec(ctx, rom):
    for r, t, _ in SECCIONES:
        if r == rom:
            ctx.body.append(escribir_par(Par.from_text('Sección %s: %s' % (r, t)), 'CertusCapitulo'))
            ctx.cap = rom

def indice():
    return 'Índice: ' + ' │ '.join('%s. %s' % (r, c) for r, _, c in SECCIONES)

# ---------------------------------------------------------------- datos por sector
S = {
 'INM': dict(
    suj='la empresa', Suj='La empresa', suf='', ent='[EMPRESA]',
    objeto='los Clientes o Usuarias de la empresa en los actos u operaciones de las fracciones V, V Bis y XV del artículo 17 de la Ley',
    A1=('A1. Tipo de acto u operación', 'Compraventa de inmueble residencial para uso propio', 'Inversión inmobiliaria habitual o arrendamiento', 'Múltiples adquisiciones, permutas u operaciones de preventa de alto valor'),
    A2=('A2. Monto del acto u operación', 'Menor a 4,000 UMA (~$469,240 MXN)', 'Entre 4,000 y 8,025 UMA (~$469,240 a $941,413 MXN)', 'Igual o superior a 8,025 UMA (≥$941,413 MXN), umbral de Aviso de las fracciones V y V Bis; en arrendamiento, renta mensual de 3,210 UMA o más'),
    A3=('A3. Origen y destino de los recursos', 'Recursos propios comprobables (nómina, actividad empresarial documentada)', 'Recursos mixtos o préstamo familiar documentado', 'Efectivo, activos virtuales o procedencia no acreditada'),
    C1=('C1. Zona geográfica del cliente o del inmueble', 'Zona residencial o comercial convencional, en entidad sin alerta', 'Zona fronteriza, turística de alta densidad o con alerta moderada', 'Zona con alta incidencia delictiva, o vínculo con jurisdicción de régimen fiscal preferente o de alto Riesgo (UIF, GAFI)'),
    D1=('D1. Canal de captación y formalización', 'Operación formalizada de manera presencial, directamente con el cliente', 'Participación parcial de medios remotos o digitales, con el cliente plenamente identificado', 'Operación a distancia, a través de apoderado sin relación previa documentada o de intermediarios no identificados plenamente'),
    ftA=('Fraccionamiento de pagos para evitar umbrales de identificación o Aviso; pagos de terceros sin relación aparente con el cliente', 'Aportaciones a un desarrollo por debajo del valor de mercado, seguidas de solicitud de reembolso o de cesión a un tercero no relacionado'),
    alto_extra='',
    nota=None,
    aviso_ficha='[ ] Alcanza el umbral de Aviso (D12)   [ ] No alcanza el umbral   [ ] Aviso de 24 horas por sospecha',
    operacion='Acto u operación (fracción):',
 ),
 'NOT': dict(
    suj='la Notaría', Suj='La Notaría', suf='-NOT', ent='la Notaría',
    objeto='los clientes de la Notaría en los actos de la fracción XII, Apartado A, del artículo 17 de la Ley',
    A1=('A1. Tipo de acto notarial (fr. XII, Apartado A)', 'a) Transmisión de inmueble residencial para uso propio del comprador identificado', 'c) Constitución de sociedad con socios identificables y giro lícito habitual; d) fideicomiso de garantía simple con acreedor institucional', 'b) Poder irrevocable a favor de tercero sin vínculo aparente; c) sociedad con estructura de capital opaca o extranjera; d) fideicomiso traslativo complejo; e) mutuo sin garantía entre partes sin vínculo aparente'),
    A2=('A2. Monto del acto', 'Menor a 4,000 UMA (~$469,240 MXN)', 'Entre 4,000 y 8,000 UMA (~$469,240 a $938,480 MXN)', 'Igual o superior a 8,000 UMA (≥$938,480 MXN), umbral de Aviso del inciso a); los incisos b), c) y e) son siempre objeto de Aviso'),
    A3=('A3. Origen de los recursos o del bien aportado', 'Recursos propios comprobables (nómina, actividad empresarial documentada)', 'Recursos mixtos, préstamo familiar o bien aportado con antigüedad acreditable', 'Efectivo, activos virtuales o procedencia no acreditada'),
    C1=('C1. Zona geográfica del inmueble o domicilio social', 'Zona residencial o comercial convencional, en entidad sin alerta', 'Zona fronteriza, turística de alta densidad o con alerta moderada', 'Zona con alta incidencia delictiva, o jurisdicción de régimen fiscal preferente o de alto Riesgo según las listas que dé a conocer la UIF'),
    D1=('D1. Canal de solicitud y comparecencia', 'Las partes solicitan el acto y comparecen personalmente ante el Notario', 'Solicitud por gestor, apoderado o despacho identificado, con comparecencia de las partes', 'Solicitud por intermediarios no identificados plenamente, apoderado sin relación previa documentada o instrucciones a distancia sin contacto con las partes'),
    ftA=('Actos reiterados del mismo cliente en poco tiempo sin justificación económica; pagos de terceros sin relación aparente con las partes', 'Precio notoriamente distinto al valor de mercado; aportaciones o mutuos que se revierten o ceden a terceros poco después de formalizados'),
    nota=('Nota: en los incisos b), c) y e) el Aviso es obligatorio siempre, sin importar el Grado de Riesgo asignado. '
          'La clasificación EBR no exime ni condiciona esa obligación: determina la intensidad de la debida diligencia '
          'que se aplica antes de formalizar el acto.'),
    aviso_ficha='[ ] Objeto de Aviso (inciso ___)   [ ] No objeto de Aviso   [ ] Aviso de 24 horas por sospecha',
    operacion='Acto notarial (inciso):',
 ),
 'COR': dict(
    suj='la Correduría', Suj='La Correduría', suf='-COR', ent='la Correduría',
    objeto='los clientes de la Correduría en los actos de la fracción XII, Apartado B, del artículo 17 de la Ley',
    A1=('A1. Tipo de acto ante el Corredor (fr. XII, Apartado B)', 'a) Avalúo con fe pública para fines contables o de crédito, solicitado por cliente identificado; b) constitución de sociedad con socios personas físicas identificables y giro habitual', 'b) Aumento de capital o compraventa de acciones entre partes relacionadas con soporte documental; c) fideicomiso de administración o garantía con fiduciaria institucional; d) crédito mercantil con garantía entre empresas con relación comercial acreditada', 'b) Socios o adquirentes extranjeros o en cascada, o aportaciones de valor dudoso; c) cesión de derechos fideicomisarios a terceros sin vínculo aparente; d) mutuo mercantil sin garantía entre partes sin vínculo aparente; a) valor de avalúo notoriamente distinto al de mercado'),
    A2=('A2. Monto del acto', 'Menor a 3,210 UMA (~$376,565 MXN)', 'Entre 3,210 y 8,025 UMA (~$376,565 a $941,413 MXN)', 'Igual o superior a 8,025 UMA (≥$941,413 MXN), umbral del inciso a); los incisos b), c) y d) son siempre objeto de Aviso'),
    A3=('A3. Origen de los recursos o de los bienes aportados', 'Recursos propios comprobables (nómina, actividad empresarial documentada, estados financieros)', 'Recursos mixtos, préstamos de terceros relacionados o bienes aportados con antigüedad acreditable', 'Efectivo, activos virtuales o procedencia no acreditada'),
    C1=('C1. Zona geográfica', 'Zona residencial o comercial convencional, en entidad sin alerta', 'Zona fronteriza, turística de alta densidad o con alerta moderada', 'Zona con alta incidencia delictiva, o jurisdicción de régimen fiscal preferente o de alto Riesgo según las listas que dé a conocer la UIF'),
    D1=('D1. Canal de solicitud y comparecencia', 'Las partes solicitan el acto y comparecen personalmente ante el Corredor', 'Solicitud por gestor, apoderado o despacho identificado, con comparecencia de las partes', 'Solicitud por intermediarios no identificados plenamente, apoderado sin relación previa documentada o instrucciones a distancia sin contacto con las partes'),
    ftA=('Actos reiterados del mismo cliente en poco tiempo sin justificación económica; aportaciones de terceros sin relación aparente con la sociedad', 'Avalúos o aportaciones con valor notoriamente distinto al de mercado; créditos o derechos fideicomisarios cedidos a terceros poco después de formalizados'),
    nota=('Nota: en los incisos b), c) y d) el Aviso es obligatorio siempre, sin importar el Grado de Riesgo asignado; en el '
          'inciso a) depende únicamente del umbral de 8,025 UMA. La clasificación EBR no exime ni condiciona esas '
          'obligaciones: determina la intensidad de la debida diligencia que se aplica antes de formalizar el acto.'),
    aviso_ficha='[ ] Objeto de Aviso (inciso ___)   [ ] No objeto de Aviso   [ ] Aviso de 24 horas por sospecha',
    operacion='Acto ante el Corredor (inciso):',
 ),
 'PROF': dict(
    suj='el profesional', Suj='El profesional', suf='-PROF', ent='el profesional',
    objeto='los clientes o asuntos del profesional en los servicios de los incisos a) a e) de la fracción XI del artículo 17 de la Ley',
    A1=('A1. Tipo de servicio (incisos a) a e) de la fr. XI)', 'a) Compraventa de inmueble residencial para uso propio del cliente identificado; c) manejo de cuenta bancaria ordinaria y trazable', 'b) Administración de activos de valor moderado; d) organización de aportaciones de capital con socios identificables y giro lícito habitual', 'b) Administración o manejo de recursos de terceros sin trazabilidad clara; e) constitución o compraventa de personas morales o vehículos corporativos con estructura opaca o multijurisdiccional'),
    A2=('A2. Monto o cuantía involucrada (indicador de Riesgo, no umbral legal)', 'Cuantía baja o no significativa respecto del perfil habitual del cliente', 'Cuantía moderada, superior al perfil habitual pero justificada', 'Cuantía elevada o desproporcionada respecto del perfil económico conocido del cliente'),
    A3=('A3. Origen de los recursos o activos involucrados', 'Recursos propios comprobables (nómina, actividad empresarial documentada)', 'Recursos mixtos, préstamo familiar o activos con antigüedad acreditable', 'Efectivo, activos virtuales o procedencia no acreditada'),
    C1=('C1. Zona geográfica del cliente, del bien o de la sociedad', 'Zona residencial o comercial convencional, en entidad sin alerta', 'Zona fronteriza, turística de alta densidad o con alerta moderada', 'Zona con alta incidencia delictiva, o jurisdicción de baja o nula tributación o de alto Riesgo según la UIF o el GAFI'),
    D1=('D1. Canal de contratación e instrucciones', 'Cliente atendido personalmente, que contrata y da instrucciones directamente', 'Instrucciones por medios electrónicos de un cliente plenamente identificado', 'Instrucciones de terceros o intermediarios no identificados plenamente, o cliente nunca atendido directamente'),
    ftA=('Instrucciones de fraccionar pagos o aportaciones; recursos de terceros sin relación aparente con el cliente', 'Constitución o compraventa de sociedades sin actividad real; administración de recursos que se mueven entre cuentas de terceros sin justificación'),
    nota=('Nota: en la fracción XI el Grado de Riesgo no determina si hay Aviso. El Aviso procede cuando el profesional '
          'lleva a cabo, en nombre y representación del cliente, una operación financiera relacionada con los incisos '
          'a) a e) (D1-PROF, sección 2.2), y siempre ante sospecha o indicios (art. 18, fr. VI, de la Ley). La '
          'clasificación EBR determina la intensidad de la debida diligencia y es especialmente relevante para decidir '
          'si se acepta o continúa un asunto con un cliente de Riesgo alto.'),
    aviso_ficha='[ ] Sí: operación financiera en nombre y representación del cliente (inciso ___)   [ ] No: solo preparación o asesoría   [ ] Aviso de 24 horas por sospecha',
    operacion='Servicio o asunto (inciso):',
 ),
}
COMUN = dict(
    B1=('B1. Perfil del cliente', 'Persona física residente en México con identificación verificada; persona moral con estructura simple y Beneficiario Controlador identificado', 'Persona física con recursos de fuente no habitual, o persona moral de reciente constitución', 'Persona moral con estructura compleja o en cascada, Beneficiario Controlador difícil de identificar, o cliente no residente'),
    B2=('B2. Vínculo con PEP', 'Sin vínculo con PEP, declarado y verificado', 'Vínculo indirecto con una PEP que no lo asimila como tal (conocido o colaborador sin vínculos patrimoniales)', 'El cliente, socio, apoderado o Beneficiario Controlador es PEP, nacional o extranjera, o está asimilado a una (art. 23 Quáter RCG)'),
)

def matriz(ctx, filas):
    tabla(ctx, [('Indicador', 'Bajo (1)', 'Medio (2)', 'Alto (3)')] + list(filas), [2200, 2400, 2400, 2500])

def construir(kit):
    d = S[kit]; ctx = Ctx(kit); suf = d['suf']
    D = lambda n: 'D%s%s' % (n, suf)
    D5 = 'D5' + suf
    portada(ctx, codigo='D6', titulo='Metodología de Enfoque Basado en Riesgo (EBR)', indice=indice())
    # I
    sec(ctx, 'I')
    p(ctx, 'La Metodología de Enfoque Basado en Riesgo (“Metodología EBR”) establece el procedimiento para identificar, '
           'analizar, medir y mitigar los Riesgos de %s, en cumplimiento del artículo 18, fracción VII, de la Ley y de los '
           'Capítulos II Quáter y III Bis RCG. Cada cliente se evalúa con siete indicadores agrupados en los cuatro '
           'elementos de Riesgo que exige el art. 10 Septies 1 RCG; la suma de sus calificaciones determina su Grado de '
           'Riesgo individual.' % d['objeto'])
    p(ctx, 'Esta Metodología forma parte del Manual de Políticas Internas (%s; art. 37, tercer párrafo, RCG). La evaluación '
           'de los Riesgos de %s como sujeto obligado, que determina su nivel de Riesgo y la ruta de la auditoría (arts. 44 y '
           '45 RCG), se documenta en la sección 5.1 del %s con los mismos cuatro elementos y con igual ponderación.' % (D5, d['suj'], D5))
    # II
    sec(ctx, 'II')
    p(ctx, 'La Metodología se alimenta con la información de un periodo no menor a doce meses: total de Clientes o Usuarias, '
           'número de actos u operaciones y monto operado (art. 10 Septies 2, fr. II, RCG). Cuando %s no cuente con '
           'operaciones en ese periodo, la Metodología se implementa con datos proyectados y se actualiza al cumplir los '
           'primeros doce meses de operación.' % d['suj'])
    p(ctx, 'La información debe ser consistente con la que obre en los mecanismos automatizados (%s, capítulo 19; art. 10 '
           'Septies 2, fr. I, RCG); cualquier inconsistencia se corrige de inmediato. Se considera además la Evaluación '
           'Nacional de Riesgos y sus actualizaciones que dé a conocer la UIF (art. 10 Septies RCG).' % D5)
    # III
    sec(ctx, 'III')
    p(ctx, 'Los indicadores se agrupan en los cuatro elementos de Riesgo del art. 10 Septies 1, fr. I, RCG. Cada indicador se '
           'califica de 1 (bajo) a 3 (alto). Los cuatro elementos tienen la misma ponderación: la puntuación total es la '
           'suma simple de los siete indicadores.')
    for titulo, filas in (('Elemento A — Actos u operaciones', (d['A1'], d['A2'], d['A3'])),
                          ('Elemento B — Tipo de Cliente o Usuaria', (COMUN['B1'], COMUN['B2'])),
                          ('Elemento C — Países y áreas geográficas', (d['C1'],)),
                          ('Elemento D — Transacciones y canales de envío o distribución', (d['D1'],))):
        p(ctx, '**%s**' % titulo)
        matriz(ctx, filas)
    # IV
    sec(ctx, 'IV')
    p(ctx, 'La Metodología incorpora, en cada elemento, indicadores específicos de los delitos de los artículos 139 Quáter '
           '(financiamiento al terrorismo) y 400 Bis (operaciones con recursos de procedencia ilícita) del Código Penal '
           'Federal. Su presencia califica el indicador correspondiente en nivel alto y se analiza conforme al %s.' % D('9'))
    tabla(ctx, [('Elemento', 'Financiamiento al terrorismo (art. 139 Quáter CPF)', 'Lavado de dinero (art. 400 Bis CPF)'),
                ('A. Actos u operaciones', d['ftA'][0], d['ftA'][1]),
                ('B. Tipo de cliente', 'Cliente o Beneficiario Controlador con vínculos declarados o presuntos con organizaciones catalogadas como terroristas por autoridad nacional o internacional', 'Estructuras corporativas opacas o multijurisdiccionales sin justificación de negocio aparente'),
                ('C. Geografía', 'Recursos u operaciones vinculados con zonas de conflicto o jurisdicciones señaladas por la ONU o el GAFI en materia de financiamiento al terrorismo', 'Recursos u operaciones vinculados con jurisdicciones de alto Riesgo conforme al art. 23 Bis 4 RCG'),
                ('D. Canales', 'Intermediarios anónimos o plataformas sin trazabilidad para el envío de recursos', 'Cadenas de transferencias sucesivas entre cuentas de terceros antes del pago')],
          [1900, 3800, 3800])
    # V
    sec(ctx, 'V')
    p(ctx, '%s cuenta con los siguientes Mitigantes, cuya efectiva aplicación se documenta en el Manual de Políticas '
           'Internas (art. 10 Septies 1, fr. III, y art. 3, fr. XI Quinquies, RCG):' % d['Suj'])
    for t in ('Integración del expediente único de identificación antes del acto u operación (%s, capítulo 4).' % D5,
              'Medidas simplificadas, completas o reforzadas según el Grado de Riesgo (%s, capítulo 5).' % D5,
              'Identificación y seguimiento reforzado de PEP (%s, capítulo 6) y consulta de listas de autoridades (%s, capítulo 16).' % (D5, D5),
              'Sistema de alertas sobre desviaciones del Perfil Transaccional (%s, capítulo 14).' % D5,
              'Aprobación de un directivo o constancia del titular para actos u operaciones con PEP de Riesgo alto (%s, sección 5.3).' % D5,
              'Capacitación anual del personal que participa en la identificación de clientes y el envío de Avisos (%s, capítulo 12).' % D5,
              'Control interno, supervisión y auditoría anual (%s, capítulo 13).' % D5):
        p(ctx, t, bullet=True)
    p(ctx, 'Si la implementación de la Metodología revela Riesgos mayores o nuevos, %s modifica sus políticas, criterios, '
           'medidas o procedimientos en un plazo no mayor a doce meses contados desde que cuente con los resultados, e '
           'indica el año y mes en que los obtuvo (art. 10 Septies 2 RCG).' % d['suj'])
    # VI
    sec(ctx, 'VI')
    p(ctx, 'La puntuación total del cliente es la suma de las calificaciones de los siete indicadores, con un rango de 7 a 21 '
           'puntos. Los cuatro elementos tienen la misma ponderación (criterio del Programa).')
    tabla(ctx, [('Puntuación total', 'Grado de Riesgo', 'Medidas'),
                ('7 a 11 puntos', 'BAJO', 'Medidas simplificadas (art. 17 RCG), siempre que el cliente cumpla los criterios de la sección 5.2 del %s. Reevaluación al menos cada seis meses (art. 23 Bis 1 RCG).' % D5),
                ('12 a 16 puntos', 'MEDIO', 'Expediente de identificación completo (art. 17 Bis RCG) y verificación adicional del origen de los recursos. Reevaluación al menos cada seis meses.'),
                ('17 a 21 puntos', 'ALTO', 'Debida diligencia reforzada (arts. 23 Ter 3 y 23 Ter 4 RCG): cuestionario reforzado (%s, Anexo B), mayor información sobre el origen y destino de los recursos y monitoreo más estricto; si además es PEP, aprobación o constancia del art. 23 Ter 5 RCG (%s, sección 5.3). Valorar no aceptar o terminar la relación (%s, capítulo 18). Reevaluación con mayor frecuencia.' % (D5, D5, D5))],
          [1700, 1300, 6500])
    if d['nota']:
        p(ctx, d['nota'])
    # VII
    sec(ctx, 'VII')
    p(ctx, 'Con independencia de la puntuación obtenida, %s clasifica como de Riesgo alto, cuando menos, a los siguientes '
           'Clientes o Usuarias (art. 23 Bis 4 RCG):' % d['suj'])
    p(ctx, 'Personas no residentes en territorio mexicano vinculadas con países o jurisdicciones que la legislación mexicana '
           'considera de régimen fiscal preferente, o que autoridades mexicanas u organismos internacionales en materia de '
           'PLD/FT de los que México sea miembro determinen que no cuentan con medidas para prevenir, detectar y combatir '
           'estas operaciones, o cuya aplicación sea deficiente.', bullet=True)
    p(ctx, 'Personas Políticamente Expuestas extranjeras.', bullet=True)
    p(ctx, 'En estos casos, %s recaba y documenta en el expediente (%s) las razones por las que el cliente celebra el acto u '
           'operación en territorio mexicano. La lista de países y jurisdicciones la pone a disposición la UIF en el Portal '
           'en Internet.' % (d['suj'], D('7')))
    # VIII
    sec(ctx, 'VIII')
    p(ctx, 'Para los clientes que sean PEP de nacionalidad mexicana, %s evalúa factores de Riesgo adicionales a fin de '
           'determinar si su comportamiento transaccional corresponde razonablemente con sus ingresos, funciones, nivel y '
           'responsabilidad, conforme a la información de que disponga (art. 23 Bis 3 RCG; %s).' % (d['suj'], D('8')))
    # IX
    sec(ctx, 'IX')
    tabla(ctx, [('Elemento', 'Periodicidad', 'Fundamento'),
                ('Grado de Riesgo de cada cliente (ficha de la sección XI)', 'Al menos cada seis meses, con mayor frecuencia a mayor Riesgo, y siempre que cambie el perfil del cliente', 'art. 23 Bis 1 RCG'),
                ('Metodología EBR en su conjunto', 'Cada doce meses, o antes si se detectan Riesgos nuevos o la UIF actualiza la Evaluación Nacional de Riesgos', 'art. 10 Septies 3 RCG'),
                ('Ajuste de Mitigantes tras detectar Riesgos mayores o nuevos', 'Dentro de los doce meses siguientes a contar con los resultados de la implementación', 'art. 10 Septies 2 RCG')],
          [3300, 4300, 1900])
    p(ctx, 'Toda revisión y actualización consta por escrito y está a disposición del SAT o de la UIF cuando la requieran.')
    # X
    sec(ctx, 'X')
    p(ctx, '%s conserva la información y documentación generada con motivo de esta Metodología, incluidas las fichas de '
           'evaluación y el historial de cambios de Grado de Riesgo y de Perfil Transaccional de cada cliente, por un plazo '
           'no menor a diez años, y la proporciona al SAT o a la UIF cuando la requieran (arts. 10 Septies 3 y 41, fr. IV, '
           'RCG; %s). La Metodología, como parte del Manual de Políticas Internas, debe estar a disposición de la autoridad '
           'a partir del 1 de marzo de 2027 (Transitorios Segundo y Tercero del Acuerdo 115/2026).' % (d['Suj'], D('10')))
    # XI
    sec(ctx, 'XI')
    tabla(ctx, [('**Cliente:**', '[Nombre o denominación del cliente]'),
                ('**Tipo:**', '[ ] Persona física   [ ] Persona moral   [ ] Fideicomiso u otra figura'),
                ('**%s**' % d['operacion'], '[Descripción]'),
                ('**Fecha de evaluación:**', '[dd/mm/aaaa]'),
                ('**Evaluador:**', '[Nombre y cargo del evaluador]'),
                ('**Próxima reevaluación (6 meses):**', '[dd/mm/aaaa]')], [3200, 6300], encabezado=False)
    p(ctx, '')
    filas = [('Indicador', 'Calificación', 'Observaciones')]
    for k in ('A1', 'A2', 'A3'):
        filas.append((d[k][0], '[ ] 1   [ ] 2   [ ] 3', ''))
    for k in ('B1', 'B2'):
        filas.append((COMUN[k][0], '[ ] 1   [ ] 2   [ ] 3', ''))
    filas.append((d['C1'][0], '[ ] 1   [ ] 2   [ ] 3', ''))
    filas.append((d['D1'][0], '[ ] 1   [ ] 2   [ ] 3', ''))
    filas.append(('**PUNTUACIÓN TOTAL (7 a 21)**', '_____', '**Grado de Riesgo:** [ ] BAJO   [ ] MEDIO   [ ] ALTO'))
    tabla(ctx, filas, [4300, 2000, 3200])
    p(ctx, '')
    tabla(ctx, [('**Riesgo alto obligatorio (sección VII):**', '[ ] No aplica   [ ] No residente vinculado con jurisdicción de Riesgo   [ ] PEP extranjera'),
                ('**PEP nacional (sección VIII):**', '[ ] No   [ ] Sí: congruencia con ingresos y funciones verificada'),
                ('**Aviso:**', d['aviso_ficha']),
                ('**Decisión:**', '[ ] Operar con monitoreo ordinario   [ ] Operar con medidas reforzadas   [ ] Aprobación o constancia del art. 23 Ter 5 RCG   [ ] No operar o terminar la relación'),
                ('**Firma del evaluador:**', ''),
                ('**Apoyo Certus PLD:**', 'contacto@certusconsultores.com.mx │ 999 163 1363 · contacto@espinosasalazarabogados.com │ 999 575 3129')],
          [3200, 6300], encabezado=False)
    pie(ctx)
    return ctx

NOMBRES = {'PROF': 'profesionistas/D6-PROF_Metodologia_EBR_v2.docx', 'INM': 'inmobiliaria/D6_Metodologia_EBR_v5.docx',
           'NOT': 'notarial/D6-NOT_Metodologia_EBR_v3.docx', 'COR': 'corredores/D6-COR_Metodologia_EBR_v2.docx'}

if __name__ == '__main__':
    B.SRC, OUT = sys.argv[1], sys.argv[2]
    for k in ('PROF', 'INM', 'NOT', 'COR'):
        ctx = construir(k)
        doc = Document(os.path.join(B.SRC, 'PROF_pr12.docx'))
        instalar_estilos(doc)
        body = doc.element.body; sect = body.find(q('w:sectPr'))
        for el in list(body):
            if el is not sect: body.remove(el)
        from docxlib import espaciar_tras_tablas
        espaciar_tras_tablas(ctx.body)
        for el in ctx.body: sect.addprevious(el)
        cp = doc.core_properties
        cp.title = 'D6%s Metodología de Enfoque Basado en Riesgo (EBR) — %s' % (canon.KITS[k]['suf'], canon.KITS[k]['variante'])
        cp.author = 'Certus Consultores · Espinosa-Salazar Abogados Corporativos'
        cp.subject = 'Programa Certus PLD — edición homologada noviembre 2026'
        dest = os.path.join(OUT, NOMBRES[k]); os.makedirs(os.path.dirname(dest), exist_ok=True)
        from docxlib import numero_de_pagina; numero_de_pagina(doc)
        doc.save(dest); print(k, dest)
