"""D9 — Reporte de Operación Inusual (ROI), homologado para los cuatro kits.

Uso: python3 build_d9.py <dir_contenedor> <dir_salida>
Secciones I–V iguales en los cuatro kits; cambian el catálogo sectorial (II.2), el tipo de acto (IV.2) y quién valida.
"""
import os, sys
sys.path.insert(0, os.path.dirname(__file__))
import build_d5 as B
from build_d5 import Ctx, p, tabla, portada, pie, escribir_par, Par
from docxlib import instalar_estilos, q, espaciar_tras_tablas
from docx import Document
import canon

SECCIONES = [
    ('I', 'Objetivo y Fundamento (art. 18, fr. VI, de la Ley; arts. 26 Bis a 27 RCG)', 'Objetivo'),
    ('II', 'Catálogo de Señales de Alerta (art. 26 Bis RCG)', 'Señales de alerta'),
    ('III', 'Procedimiento de Reporte Interno', 'Procedimiento'),
    ('IV', 'Formato de Reporte de Operación Inusual', 'Formato'),
    ('V', 'Confidencialidad (arts. 38, 50 y 63, fr. II, de la Ley; art. 31 RCG)', 'Confidencialidad'),
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

GENERALES = [
    ('Efectivo', 'Pretensión de pagar en efectivo un acto u operación cuyo valor alcanza el umbral prohibido por el art. 32 de la Ley, o de fraccionar pagos o declarar un valor menor para eludir la prohibición o los umbrales de Aviso.', 'ALTO'),
    ('Identidad', 'La persona no acredita su identidad, presenta documentos alterados o se niega a proporcionarlos.', 'ALTO'),
    ('Recursos', 'No puede justificar el origen o el destino de los recursos, o los justifica con documentos inconsistentes.', 'ALTO'),
    ('Terceros', 'Pagos de terceros sin relación aparente, uso de prestanombres o indicios de que actúa por cuenta de otra persona sin declararlo (art. 22 RCG).', 'ALTO'),
    ('Beneficiario Controlador', 'Se niega a identificar al Beneficiario Controlador o la información es inconsistente (D7, Sección E).', 'ALTO'),
    ('PEP', 'Es PEP o persona asimilada y no acredita recursos congruentes con sus ingresos y funciones (D8).', 'ALTO'),
    ('Geografía', 'Vínculo sin justificación con países o jurisdicciones de régimen fiscal preferente o de Riesgo alto según la lista de la UIF.', 'ALTO'),
    ('Listas', 'Coincidencia confirmada con la lista de personas a que se refiere el art. 38 RCG: no se opera y el Aviso es obligatorio (art. 27 RCG).', 'ALTO'),
    ('Perfil Transaccional', 'El acto u operación se aparta del Perfil Transaccional del cliente (D7, Sección G) sin explicación razonable.', 'MEDIO'),
    ('Conducta', 'Prisa injustificada, desinterés por el precio o las condiciones, o instrucciones para cambiar datos ya asentados.', 'MEDIO'),
    ('Medios de pago', 'Propuesta de pago con activos virtuales o con instrumentos difíciles de rastrear, sin justificación.', 'MEDIO'),
]

S = {
 'INM': dict(suf='', suj='la empresa', Suj='La empresa', valida='[PF: Titular][PM: Representante Encargado de Cumplimiento]',
    pm_extra=True,
    tipo='[ ] Compraventa   [ ] Intermediación   [ ] Preventa o aportación a desarrollo (V Bis)   [ ] Arrendamiento (XV)',
    bien=('Inmueble o desarrollo relacionado:', '[Descripción y ubicación]'),
    catalogo=[
     ('Fracción V: compraventa o intermediación', [
        ('Precio pactado significativamente inferior o superior al valor de mercado, sin justificación económica.', 'MEDIO'),
        ('Adquisición de varios inmuebles en poco tiempo por la misma persona o grupo relacionado, sin justificación.', 'MEDIO'),
        ('Reventa inmediata del inmueble con una diferencia de precio que no se explica por el mercado.', 'MEDIO'),
        ('Adquirente cuyo perfil económico declarado es incongruente con el valor del inmueble.', 'ALTO')]),
     ('Fracción V Bis: recepción de recursos para un desarrollo inmobiliario', [
        ('Aportaciones desde cuentas de terceros no identificados o sin relación con el aportante.', 'ALTO'),
        ('Aportaciones por debajo del valor de mercado seguidas de una solicitud de reembolso o de cesión a un tercero no relacionado.', 'ALTO'),
        ('Cancelación anticipada con solicitud de devolver los recursos a una cuenta distinta de la de origen.', 'ALTO')]),
     ('Fracción XV: arrendamiento', [
        ('Pago anticipado de varios meses o años de renta sin justificación.', 'ALTO'),
        ('Arrendatario cuyo perfil no corresponde con el monto de la renta.', 'MEDIO'),
        ('Inmueble sin uso aparente o destinado a un fin distinto del declarado.', 'MEDIO')])],
    aviso_mensual='el Aviso mensual por alcanzar el umbral (D12)'),
 'NOT': dict(suf='-NOT', suj='la Notaría', Suj='La Notaría', valida='Notario Titular', pm_extra=False,
    tipo='[ ] a) Inmuebles   [ ] b) Poder irrevocable   [ ] c) Sociedades   [ ] d) Fideicomiso   [ ] e) Mutuo o crédito',
    bien=('No. de instrumento (si se formalizó):', '[Número]'),
    catalogo=[
     ('Inciso a): transmisión o constitución de derechos reales sobre inmuebles', [
        ('Precio pactado significativamente inferior al valor catastral o comercial, sin justificación económica.', 'MEDIO'),
        ('Adquirente que no explica el origen de los recursos o cuyo perfil es incongruente con el valor del inmueble.', 'ALTO'),
        ('Varias adquisiciones por la misma persona o grupo relacionado en poco tiempo, sin justificación de negocio.', 'MEDIO')]),
     ('Inciso b): poderes irrevocables', [
        ('Poder irrevocable de dominio a favor de una persona sin vínculo familiar, societario o comercial aparente con el poderdante.', 'ALTO'),
        ('Poderdante de edad avanzada o en aparente condición de vulnerabilidad, sin acompañamiento de familiares.', 'ALTO'),
        ('Facultades de dominio amplísimas sin correspondencia con el patrimonio declarado del poderdante.', 'MEDIO')]),
     ('Inciso c): sociedades, capital, fusión, escisión y acciones', [
        ('Capital social desproporcionado para el giro declarado, aportado en efectivo o sin soporte documental.', 'ALTO'),
        ('Estructura de socios en cascada que dificulta identificar al Beneficiario Controlador.', 'ALTO'),
        ('Aumento súbito de capital poco después de la constitución, sin correspondencia con la operación del negocio.', 'MEDIO'),
        ('Compraventa de acciones o partes sociales entre partes sin vínculo aparente, a un precio muy alejado del valor contable.', 'MEDIO')]),
     ('Inciso d): fideicomisos traslativos de dominio o de garantía', [
        ('Fideicomitente que aporta bienes de valor muy superior a su perfil económico declarado.', 'ALTO'),
        ('Fideicomisarios múltiples sin relación familiar o societaria clara con el fideicomitente.', 'MEDIO'),
        ('Modificaciones frecuentes que cambian a los fideicomisarios en periodos cortos.', 'MEDIO')]),
     ('Inciso e): mutuo o crédito', [
        ('Mutuo sin garantía entre partes sin vínculo aparente y por montos elevados.', 'ALTO'),
        ('Mutuante cuyo perfil de ingresos no es congruente con el monto prestado.', 'ALTO'),
        ('Condiciones inusuales: tasa nula, plazo indefinido o sin exigibilidad real.', 'MEDIO')])],
    aviso_mensual='el Aviso mensual del acto (D12-NOT)'),
 'COR': dict(suf='-COR', suj='la Correduría', Suj='La Correduría', valida='Corredor Público Titular', pm_extra=False,
    tipo='[ ] a) Avalúo con fe pública   [ ] b) Sociedades   [ ] c) Fideicomiso   [ ] d) Mutuo o crédito mercantil',
    bien=('No. de póliza o acta (si se formalizó):', '[Número]'),
    catalogo=[
     ('Inciso a): avalúos con fe pública', [
        ('Presión del solicitante para obtener un valor predeterminado, notoriamente distinto del de mercado, sin soporte técnico.', 'ALTO'),
        ('Avalúos fraccionados (partes de una misma unidad económica) que quedan justo por debajo de 8,025 UMA.', 'ALTO'),
        ('Solicitante que no acredita la titularidad o legítima posesión del bien, o cuyo interés en el avalúo no es claro.', 'MEDIO'),
        ('Avalúo de bienes o acciones que se aportarán a capital, con un valor desproporcionado respecto de la información financiera de la sociedad.', 'MEDIO')]),
     ('Inciso b): sociedades mercantiles, capital, fusión, escisión y acciones', [
        ('Capital social desproporcionado para el giro declarado, aportado en efectivo o sin soporte documental.', 'ALTO'),
        ('Estructura de socios en cascada que dificulta identificar al Beneficiario Controlador.', 'ALTO'),
        ('Aumento súbito de capital poco después de la constitución, sin correspondencia con la operación del negocio.', 'MEDIO'),
        ('Compraventa de acciones o partes sociales entre partes sin vínculo aparente, a un precio muy alejado del valor contable.', 'MEDIO')]),
     ('Inciso c): fideicomisos', [
        ('Fideicomitente que aporta bienes de valor muy superior a su perfil económico declarado.', 'ALTO'),
        ('Cesión de derechos fideicomisarios a terceros sin vínculo aparente, o a un precio muy alejado de su valor.', 'ALTO'),
        ('Modificaciones frecuentes que cambian a los fideicomisarios o el destino del patrimonio en periodos cortos.', 'MEDIO')]),
     ('Inciso d): mutuo o crédito mercantil', [
        ('Mutuo o crédito sin garantía entre partes sin vínculo comercial aparente y por montos elevados.', 'ALTO'),
        ('Acreditante cuyo giro o capacidad financiera no es congruente con el monto prestado.', 'ALTO'),
        ('Condiciones inusuales (tasa nula, plazo indefinido) o liquidación anticipada inmediata.', 'MEDIO'),
        ('Créditos sucesivos entre las mismas partes que parecen reciclar recursos.', 'MEDIO')])],
    aviso_mensual='el Aviso mensual del acto (D12-COR)'),
 'PROF': dict(suf='-PROF', suj='[PF: el profesional][PM: el despacho]', Suj='[PF: El profesional][PM: El despacho]',
    valida='[PF: Profesional Titular][PM: Representante Encargado de Cumplimiento]', pm_extra=True,
    tipo='[ ] a) Inmuebles   [ ] b) Administración de recursos   [ ] c) Cuentas   [ ] d) Aportaciones de capital   [ ] e) Personas morales o vehículos corporativos',
    bien=('Momento de sujeción (D1-PROF, sección 2.2):', '[ ] Solo preparación o asesoría   [ ] Operación financiera en nombre y representación del cliente'),
    catalogo=[
     ('Inciso a): compraventa de inmuebles o cesión de derechos', [
        ('Precio pactado significativamente inferior al valor catastral o comercial, sin justificación económica.', 'MEDIO'),
        ('Adquirente que no explica el origen de los recursos o cuyo perfil es incongruente con el valor del inmueble.', 'ALTO')]),
     ('Inciso b): administración y manejo de recursos, valores o activos', [
        ('Instrucciones de transferir los recursos administrados a jurisdicciones de Riesgo alto o de baja o nula tributación.', 'ALTO'),
        ('Movimientos sin relación aparente con la actividad económica declarada del cliente.', 'ALTO'),
        ('Solicitud de retiros fraccionados en efectivo, sin justificación.', 'ALTO')]),
     ('Inciso c): manejo de cuentas bancarias, de ahorro o de valores', [
        ('Depósitos múltiples y fraccionados por montos que evitan los umbrales de reporte bancario.', 'ALTO'),
        ('La cuenta se usa como paso de fondos, sin relación con la actividad del cliente.', 'ALTO'),
        ('Solicitud de facultades sobre la cuenta a favor de terceros no identificados o no relacionados.', 'MEDIO')]),
     ('Inciso d): organización de aportaciones de capital', [
        ('Aportantes que no acreditan el origen de los recursos aportados.', 'ALTO'),
        ('Aportaciones desproporcionadas respecto del perfil económico del aportante.', 'ALTO')]),
     ('Inciso e): personas morales, vehículos corporativos y compraventa de entidades mercantiles', [
        ('Estructura de socios en cascada que dificulta identificar al Beneficiario Controlador.', 'ALTO'),
        ('Capital social desproporcionado para el giro declarado, aportado en efectivo o sin soporte documental.', 'ALTO'),
        ('Compraventa de la entidad entre partes sin vínculo aparente, a un precio muy alejado del valor contable.', 'MEDIO')])],
    aviso_mensual='el Aviso mensual que procede solo cuando el profesional realiza la operación financiera en nombre y representación del cliente (D12-PROF)'),
}

def construir(kit):
    d = S[kit]; ctx = Ctx(kit); s = d['suf']
    D = lambda n: 'D%s%s' % (n, s)
    portada(ctx, codigo='D9', titulo='Reporte de Operación Inusual (ROI)', indice=indice())
    # I
    sec(ctx, 'I')
    p(ctx, 'Este documento guía la detección, el análisis y el reporte interno de los actos u operaciones inusuales, y '
           'documenta la decisión de presentar el Aviso de 24 horas ante la Secretaría. Forma parte del Manual de Políticas '
           'Internas (%s, capítulo 8).' % D('5'))
    p(ctx, 'Procede presentar Aviso dentro de las 24 horas siguientes en dos supuestos (art. 18, fr. VI, segundo párrafo, de '
           'la Ley):')
    p(ctx, '**Aviso por sospecha:** cuando, con la información del expediente y las características habituales de los actos '
           'u operaciones, %s reconoce en el cliente una actividad, conducta o comportamiento inusitado que podría estar '
           'vinculado con operaciones con recursos de procedencia ilícita (art. 26 Bis RCG).' % d['suj'], bullet=True)
    p(ctx, '**Aviso por hechos o indicios:** cuando conoce, por fuentes públicas o privadas, un hecho relacionado con el '
           'cliente que permita considerar que los recursos podrían estar vinculados con esos delitos (art. 26 Bis 1 RCG).', bullet=True)
    p(ctx, 'Estos Avisos proceden aunque el acto u operación no alcance el umbral de Aviso o no se haya celebrado, siempre que '
           'se cuente con datos para identificar a la persona (art. 26 Bis 2 RCG), y son independientes de %s. También es '
           'obligatorio cuando la persona figure en la lista del art. 38 RCG (art. 27 RCG).' % d['aviso_mensual'])
    if kit == 'PROF':
        p(ctx, 'En la fracción XI el Aviso de 24 horas procede aun en la fase de preparación del servicio. Presentarlo no '
               'transgrede el secreto profesional (art. 22 de la Ley), sin perjuicio del criterio sobre secreto profesional y '
               'garantía de defensa del D1-PROF, sección IV.')
    p(ctx, 'Envío: el Portal habilitará estos Avisos seis meses después de que entre en vigor la Resolución que modifique los '
           'formatos oficiales y los identifique expresamente (Transitorio Quinto del Acuerdo 115/2026). Mientras tanto, cada '
           'reporte se integra y conserva con este formato, y Certus PLD informará la fecha y el medio de envío en cuanto la '
           'autoridad los publique.')
    # II
    sec(ctx, 'II')
    p(ctx, 'Para valorar si un acto u operación es inusual se consideran, al menos, los montos, la frecuencia, las formas, '
           'periodos e instrumentos de pago, el área geográfica y el tipo o naturaleza del acto (art. 26 Bis, segundo párrafo, '
           'RCG). Las señales son enunciativas: una sola no obliga a presentar Aviso, pero toda señal se documenta y analiza.')
    sub(ctx, 'II.1 Señales generales (todos los actos u operaciones)')
    filas = [('No.', 'Categoría', 'Señal de alerta', 'Nivel')]
    for i, (c, t, n) in enumerate(GENERALES, 1):
        filas.append(('G%d' % i, c, t, n))
    tabla(ctx, filas, [700, 1800, 6000, 1000])
    sub(ctx, 'II.2 Señales propias del sector')
    n = 0
    for titulo, items in d['catalogo']:
        p(ctx, '**%s**' % titulo)
        filas = [('No.', 'Señal de alerta', 'Nivel')]
        for t, niv in items:
            n += 1
            filas.append(('S%d' % n, t, niv))
        tabla(ctx, filas, [700, 7800, 1000])
    # III
    sec(ctx, 'III')
    R = d['valida']
    tabla(ctx, [('Paso', 'Acción', 'Responsable', 'Plazo'),
                ('1', 'Detectar la señal de alerta durante la integración del expediente (%s) o la atención del cliente.' % D('7'),
                      'Personal que atiende al cliente', 'Al detectarla'),
                ('2', 'Comunicarla de inmediato y registrar la fecha y hora en que se reconoció la conducta o el indicio (sección IV.1); desde ese momento corre el plazo de 24 horas.',
                      'Personal que atiende al cliente', 'Inmediatamente'),
                ('3', 'Analizar la señal con el expediente, el Grado de Riesgo y el Perfil Transaccional, y decidir si existe sospecha o indicio.', R, 'Dentro del plazo'),
                ('4', 'Si el cliente no proporciona la información o documentación requerida, abstenerse de realizar el acto u operación (art. 21, segundo párrafo, de la Ley).', R, 'Antes de operar'),
                ('5', 'Si existe sospecha o indicio: presentar el Aviso, aun si el acto no se celebró o no alcanza el umbral. Si se descarta: documentar los motivos.', R, '24 horas'),
                ('6', 'Decidir si se continúa, se limita o se termina la relación (%s, capítulo 18) y, si el cliente es PEP de Riesgo alto, aplicar el %s, sección VIII.' % (D('5'), D('8')), R, 'Mismo día'),
                ('7', 'Conservar el reporte y, en su caso, el acuse del Aviso en el expediente del cliente, por al menos diez años (art. 18, fr. IV, de la Ley).', R, 'Mismo día')],
          [700, 5500, 2000, 1300])
    # IV
    sec(ctx, 'IV')
    sub(ctx, 'IV.1 Datos del reporte')
    ficha(ctx, [('No. de reporte interno:', '[Número secuencial]'),
                ('Tipo de Aviso:', '[ ] Por sospecha (art. 26 Bis RCG)   [ ] Por hechos o indicios (art. 26 Bis 1 RCG)   [ ] Coincidencia en lista (art. 27 RCG)'),
                ('Fecha y hora en que se reconoció la conducta o el indicio:', '[dd/mm/aaaa] — [hh:mm]'),
                ('Plazo límite de envío (24 horas después):', '[dd/mm/aaaa] — [hh:mm]'),
                ('Detectó (nombre y cargo):', '')])
    sub(ctx, 'IV.2 Acto u operación')
    ficha(ctx, [('Tipo de acto u operación:', d['tipo']),
                d['bien'],
                ('Fecha del acto u operación (si aplica):', '[dd/mm/aaaa]'),
                ('Monto involucrado:', '$ ______________ MXN'),
                ('Umbral de Aviso (art. 26 Bis 2 RCG):', '[ ] Lo alcanza   [ ] No lo alcanza   [ ] Acto u operación no celebrado'),
                ('Forma de pago:', '[ ] Transferencia   [ ] Cheque   [ ] Efectivo   [ ] Activos virtuales   [ ] Mixto   [ ] Otro: ____')])
    sub(ctx, 'IV.3 Persona involucrada')
    ficha(ctx, [('Nombre o denominación:', ''), ('RFC o CURP:', ''),
                ('Tipo:', '[ ] Persona física   [ ] Persona moral   [ ] Fideicomiso u otra figura'),
                ('Relación con %s:' % ('[PF: el profesional][PM: el despacho]' if kit == 'PROF' else d['suj']),
                 '[ ] Cliente   [ ] Otra parte del acto   [ ] Beneficiario Controlador   [ ] Tercero que paga   [ ] Otro: ____'),
                ('No. de expediente (%s):' % D('7'), ''),
                ('Grado de Riesgo (%s):' % D('6'), '[ ] BAJO   [ ] MEDIO   [ ] ALTO — Puntuación: _____ / 21'),
                ('¿Es PEP o persona asimilada? (%s):' % D('8'), '[ ] Sí   [ ] No'),
                ('Beneficiario Controlador identificado (%s, Sección E):' % D('7'), '[ ] Sí   [ ] No   [ ] No aplica')])
    sub(ctx, 'IV.4 Señales identificadas y descripción')
    ficha(ctx, [('Señales de la sección II (números):', 'G__, G__, S__, S__'),
                ('Nivel más alto identificado:', '[ ] MEDIO   [ ] ALTO')])
    p(ctx, 'Descripción narrativa: [hechos, antecedentes, personas, fechas, montos, conversaciones relevantes y razones por '
           'las que la situación se considera inusual].')
    sub(ctx, 'IV.5 Resolución y seguimiento')
    ficha(ctx, [('Decisión:', '[ ] Se presenta Aviso de 24 horas   [ ] Se descarta la sospecha (motivos abajo)   [ ] Abstención por falta de información (art. 21 de la Ley)'),
                ('Motivos de la decisión:', ''),
                ('Fecha y hora de envío del Aviso:', '[dd/mm/aaaa] — [hh:mm]'),
                ('¿Se cumplió el plazo de 24 horas?:', '[ ] Sí   [ ] No — justificar: ____'),
                ('Folio o acuse del Aviso:', ''),
                ('¿Se celebró el acto u operación?:', '[ ] Sí   [ ] No'),
                ('Relación con el cliente:', '[ ] Continúa con monitoreo reforzado   [ ] Se limita   [ ] Se termina (%s, capítulo 18)' % D('5')),
                ('Decidió (%s), nombre y firma:' % R, '')])
    if d['pm_extra']:
        p(ctx, 'El Representante Encargado de Cumplimiento informa la decisión al órgano de administración, sin '
               'revelar el reporte a personas no autorizadas.', pm=True)
    # V
    sec(ctx, 'V')
    p(ctx, 'Este reporte y, en su caso, el Aviso son confidenciales y reservados (art. 38 de la Ley). Está prohibido alertar '
           'o dar aviso al cliente o a terceros sobre su existencia, sobre requerimientos de información de la autoridad o '
           'sobre órdenes de aseguramiento antes de su ejecución (art. 50 de la Ley; art. 31 RCG; %s, capítulo 17). Revelar '
           'sin autorización información que vincule a una persona con un Aviso o requerimiento es delito sancionado con '
           'prisión de cuatro a diez años (art. 63, fr. II, de la Ley). Informar al cliente sobre las obligaciones que '
           'derivan de la Ley no viola este deber (art. 31, último párrafo, RCG).' % D('5'))
    pie(ctx)
    return ctx

NOMBRES = {'PROF': 'profesionistas/D9-PROF_Reporte_Operacion_Inusual_v2.docx', 'INM': 'inmobiliaria/D9_Reporte_Operacion_Inusual_v5.docx',
           'NOT': 'notarial/D9-NOT_Reporte_Operacion_Inusual_v3.docx', 'COR': 'corredores/D9-COR_Reporte_Operacion_Inusual_v2.docx'}

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
        cp.title = 'D9%s Reporte de Operación Inusual (ROI) — %s' % (canon.KITS[k]['suf'], canon.KITS[k]['variante'])
        cp.author = 'Certus Consultores · Espinosa-Salazar Abogados Corporativos'
        cp.subject = 'Programa Certus PLD — edición homologada noviembre 2026'
        dest = os.path.join(OUT, NOMBRES[k]); os.makedirs(os.path.dirname(dest), exist_ok=True)
        from docxlib import numero_de_pagina; numero_de_pagina(doc)
        doc.save(dest); print(k, dest)
