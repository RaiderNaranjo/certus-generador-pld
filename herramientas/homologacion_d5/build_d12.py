"""D12 — Control de Avisos e Informes, homologado para los cuatro kits.

Uso: python3 build_d12.py <dir_contenedor> <dir_salida>
"""
import os, sys
sys.path.insert(0, os.path.dirname(__file__))
import build_d5 as B
from build_d5 import Ctx, p, tabla, portada, pie, escribir_par, Par
from docxlib import instalar_estilos, q, espaciar_tras_tablas
from docx import Document
import canon

SECCIONES = [
    ('I', 'Marco de la Obligación (arts. 17, 23 y 24 de la Ley)', 'Obligación'),
    ('II', 'Fecha del Acto u Operación (art. 24 Bis RCG)', 'Fecha del acto'),
    ('III', 'Bitácora de Actos u Operaciones', 'Bitácora'),
    ('IV', 'Libro de Control Mensual e Informe en Ceros (art. 25 RCG)', 'Control mensual'),
    ('V', 'Control de Avisos de 24 Horas (arts. 26 Bis a 27 RCG)', 'Avisos de 24 horas'),
    ('VI', 'Presentación y Verificación Mensual', 'Presentación'),
]

def sec(ctx, rom):
    t = [s[1] for s in SECCIONES if s[0] == rom][0]
    ctx.body.append(escribir_par(Par.from_text('Sección %s: %s' % (rom, t)), 'CertusCapitulo'))

def indice():
    return 'Índice: ' + ' │ '.join('%s. %s' % (r, c) for r, _, c in SECCIONES)

S = {
 'INM': dict(suf='', suj='la empresa', Resp='[PF: Titular][PM: Representante Encargado de Cumplimiento]',
    umbrales=[('V', 'Compraventa o intermediación de inmuebles', '8,025 UMA por acto u operación', '$941,413'),
              ('V Bis', 'Recepción de recursos para un desarrollo inmobiliario', '8,025 UMA por acto u operación', '$941,413'),
              ('XV', 'Arrendamiento (uso o goce de inmuebles)', '3,210 UMA de renta mensual', '$376,565 al mes')],
    col='Fracción',
    fechas=[('V', 'Aquella en que se tenga por liquidado el acto u operación y se alcance el umbral de Aviso (art. 24 Bis, fr. III, RCG).'),
            ('V Bis', 'Aquella en que se recibió y destinó la última aportación al desarrollo inmobiliario en el mes calendario (art. 24 Bis, fr. IV, RCG).'),
            ('XV', 'El art. 24 Bis RCG no fija una regla propia; se aplican los arts. 5 y 24 del Reglamento. Criterio del Programa: la fecha en que se liquida la renta mensual que alcanza el umbral.')],
    acum='Los actos u operaciones de las fracciones V, V Bis y XV con un mismo cliente se acumulan en periodos de hasta seis meses, contados desde el primero; si la suma alcanza el umbral, procede el Aviso (art. 17, penúltimo párrafo, de la Ley; art. 19 RCG).',
    especiales=['**Desarrollos (V Bis):** puede presentarse un solo Aviso por todas las aportaciones del mes aplicadas al mismo desarrollo, siempre que no se rebase el plazo del art. 23 de la Ley respecto de cada aportación (art. 24 Bis 1, tercer párrafo, RCG). Registro en la tabla IV.1.',
                 '**Copropietarios (XV):** cada copropietario cumple sus obligaciones; el monto del acto es la renta mensual total pactada y el Aviso señala el pago recibido en lo individual (art. 24 Bis 6 RCG).'],
    consolidado=True, fiscal=False),
 'NOT': dict(suf='-NOT', suj='la Notaría', Resp='Notario Titular',
    umbrales=[('a)', 'Transmisión o constitución de derechos reales sobre inmuebles (el mayor entre precio, valor catastral, valor comercial o monto garantizado)', '8,000 UMA', '$938,480'),
              ('b)', 'Poderes irrevocables para actos de administración o dominio', 'Siempre', '—'),
              ('c)', 'Constitución de personas morales, modificación patrimonial, fusión, escisión y compraventa de acciones o partes sociales', 'Siempre', '—'),
              ('d)', 'Constitución o modificación de fideicomisos traslativos de dominio o de garantía', '4,000 UMA', '$469,240'),
              ('e)', 'Mutuo o crédito, con o sin garantía, con acreedor fuera del sistema financiero', 'Siempre', '—')],
    col='Inciso (fr. XII-A)',
    fechas=[('Todos los incisos', 'El art. 24 Bis RCG no fija una regla propia para la fracción XII; se aplican los arts. 5 y 24 del Reglamento. Criterio del Programa: la fecha de firma del instrumento.')],
    acum='Los actos de los incisos a) y d) con un mismo cliente se acumulan en periodos de hasta seis meses; los incisos b) y e) se excluyen de la acumulación (art. 19 RCG) y el inciso c) se avisa siempre (art. 17, penúltimo párrafo, de la Ley).',
    especiales=['**Inciso a), vía fiscal:** el Notario puede cumplir los Avisos del inciso a) mediante los medios de las disposiciones fiscales federales para declaraciones y avisos, siempre que contengan la información del art. 24 de la Ley (art. 24, último párrafo, de la Ley). Medio que usa la Notaría: [ ] Portal en Internet del SAT   [ ] Vía fiscal.'],
    consolidado=False, fiscal=True),
 'COR': dict(suf='-COR', suj='la Correduría', Resp='Corredor Público Titular',
    umbrales=[('a)', 'Avalúos sobre bienes', '8,025 UMA de valor del bien', '$941,413'),
              ('b)', 'Constitución de personas morales mercantiles, modificación patrimonial, fusión, escisión y compraventa de acciones o partes sociales', 'Siempre', '—'),
              ('c)', 'Constitución, modificación o cesión de derechos de fideicomisos', 'Siempre', '—'),
              ('d)', 'Mutuo o crédito mercantil con acreedor fuera del sistema financiero', 'Siempre', '—')],
    col='Inciso (fr. XII-B)',
    fechas=[('Todos los incisos', 'El art. 24 Bis RCG no fija una regla propia para la fracción XII; se aplican los arts. 5 y 24 del Reglamento. Criterio del Programa: la fecha de firma de la póliza o acta, o de emisión del avalúo.')],
    acum='Los incisos del Apartado B no se acumulan (art. 19 RCG): cada acto se evalúa por separado, sin perjuicio del Aviso de 24 horas si se detecta un fraccionamiento artificial (D9-COR).',
    especiales=[], consolidado=False, fiscal=False),
 'PROF': dict(suf='-PROF', suj='[PF: el profesional][PM: el despacho]', Resp='[PF: Profesional Titular][PM: Representante Encargado de Cumplimiento]',
    umbrales=[('a)', 'Compraventa de inmuebles o cesión de derechos sobre estos', 'Operación financiera en representación', '—'),
              ('b)', 'Administración y manejo de recursos, valores o activos', 'Operación financiera en representación', '—'),
              ('c)', 'Manejo de cuentas bancarias, de ahorro o de valores', 'Operación financiera en representación', '—'),
              ('d)', 'Organización de aportaciones de capital', 'Operación financiera en representación', '—'),
              ('e)', 'Constitución, escisión, fusión, operación y administración de personas morales o vehículos corporativos', 'Operación financiera en representación', '—')],
    col='Inciso (fr. XI)',
    fechas=[('Todos los incisos', 'Aquella en la que el prestador del servicio realice la operación financiera en representación o por instrucción del cliente (art. 24 Bis, fr. VII, RCG).')],
    acum='La fracción XI no tiene umbral monetario ni está comprendida en la acumulación del art. 19 RCG: procede un Aviso por cada operación financiera realizada en nombre y representación del cliente. Preparar o asesorar, sin representar, no genera Aviso (D1-PROF, sección 2.2), con respeto al secreto profesional y la garantía de defensa (D1-PROF, sección IV).',
    especiales=[], consolidado=False, fiscal=False),
}

def construir(kit):
    d = S[kit]; ctx = Ctx(kit); s = d['suf']
    D = lambda n: 'D%s%s' % (n, s)
    R = d['Resp']
    portada(ctx, codigo='D12', titulo='Control de Avisos e Informes', indice=indice())
    # I
    sec(ctx, 'I')
    p(ctx, 'Los Avisos se presentan ante la Secretaría, por conducto del SAT, a más tardar el día 17 del mes inmediato '
           'siguiente a aquel en que se realizó el acto u operación que les da origen (art. 23 de la Ley), a través del Portal '
           'en Internet y en el formato oficial (art. 24 de la Ley; art. 24 RCG). Se presenta un Aviso por cada acto u '
           'operación que, individualmente o por acumulación, alcance el umbral o la condición de Aviso (art. 24 Bis 1 RCG). '
           'Este documento desarrolla el %s, capítulo 10.' % D('5'))
    tabla(ctx, [(d['col'], 'Acto u operación', 'Umbral de Aviso', 'Equivalente 2026')] + d['umbrales'],
          [1500, 4600, 2000, 1400])
    p(ctx, 'UMA 2026: $117.31. Los montos se actualizan cada año con la UMA vigente.')
    p(ctx, '**Acumulación.** %s' % d['acum'])
    for e in d['especiales']:
        p(ctx, e)
    p(ctx, '**Consecuencias.** Presentar el Aviso dentro de los 30 días siguientes al plazo se sanciona con multa de 200 a '
           '2,000 UMA (arts. 53, fr. III, y 54, fr. I, de la Ley); pasado ese plazo, se sanciona como omisión, con multa de '
           '10,000 a 65,000 UMA o del 10% al 100% del valor del acto, la que resulte mayor (arts. 53, fr. VI, y 54, fr. III).')
    # II
    sec(ctx, 'II')
    p(ctx, 'El plazo del art. 23 de la Ley se cuenta desde la fecha del acto u operación, que se determina así:')
    tabla(ctx, [(d['col'], 'Fecha del acto u operación')] + d['fechas'], [2200, 7300])
    # III
    sec(ctx, 'III')
    p(ctx, 'Se registra cada acto u operación del periodo, sea o no objeto de Aviso, para verificar umbrales y acumulación.')
    filas = [('No.', 'Exp. (%s)' % D('7'), 'Fecha del acto', d['col'].split(' ')[0], 'Cliente', 'Monto', '¿Aviso?', 'Fecha límite', 'Folio')]
    filas += [(str(i), '', '', '', '', '$', '[ ] Sí [ ] No', '', '') for i in range(1, 7)]
    tabla(ctx, filas, [500, 1000, 1100, 900, 1600, 1100, 1200, 1100, 1000])
    # IV
    sec(ctx, 'IV')
    p(ctx, 'En los meses en que no se realicen actos u operaciones objeto de Aviso se presenta, en el mismo formato, un '
           'informe en el que solo se llenan los datos de identificación, el periodo y la manifestación de que no hubo actos '
           'u operaciones objeto de Aviso. Una vez enviado no puede modificarse ni eliminarse; si después se identifica un '
           'acto que sí era objeto de Aviso, este se presenta de inmediato (art. 25 RCG).')
    meses = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']
    tabla(ctx, [('Mes', 'Avisos', 'Tipo', 'Fecha de presentación', 'Folio o acuse', 'Oportunidad', 'Firma')] +
          [(m, '', '[ ] Aviso   [ ] Informe en ceros', '', '', '[ ] En plazo   [ ] Extemporáneo', '') for m in meses],
          [1200, 800, 2000, 1400, 1300, 1700, 1100])
    if d['consolidado']:
        ctx.body.append(escribir_par(Par.from_text('IV.1 Avisos mensuales consolidados por desarrollo (art. 24 Bis 1 RCG)'), 'CertusSeccion'))
        tabla(ctx, [('Desarrollo inmobiliario', 'Mes', 'Aportaciones incluidas', 'Monto consolidado', 'Folio'),
                    ('[Nombre del desarrollo]', '', '', '$', '')], [2800, 1200, 1900, 2000, 1600])
    # V
    sec(ctx, 'V')
    p(ctx, 'Los Avisos por sospecha y por hechos o indicios se presentan dentro de las 24 horas siguientes a que se reconoce '
           'la conducta o el indicio, aunque el acto no alcance el umbral o no se haya celebrado (art. 18, fr. VI, de la Ley; '
           'arts. 26 Bis a 26 Bis 2 RCG); también cuando la persona figure en la lista del art. 38 RCG (art. 27 RCG). Son '
           'independientes del Aviso mensual. El detalle de cada caso se documenta en el %s. El envío por el Portal se '
           'habilita seis meses después de la Resolución que modifique los formatos oficiales (Transitorio Quinto del Acuerdo '
           '115/2026).' % D('9'))
    tabla(ctx, [('No. de ROI (%s)' % D('9'), 'Tipo', 'Fecha y hora de detección', 'Plazo límite', 'Fecha y hora de envío', '¿En plazo?')] +
          [('', '[ ] Sospecha [ ] Indicios [ ] Lista', '', '', '', '[ ] Sí [ ] No')] * 3,
          [1400, 2200, 1700, 1400, 1700, 1100])
    # VI
    sec(ctx, 'VI')
    pasos = ['Ingresar al Portal en Internet del SAT (https://sppld.sat.gob.mx/pld/index.html) con la e.firma: [PF: la del titular (art. 24, primer párrafo, RCG)][PM: la asociada al RFC de la persona moral, no la del representante legal (art. 24, segundo párrafo, RCG)].' if kit in ('INM', 'PROF')
             else 'Ingresar al Portal en Internet del SAT (https://sppld.sat.gob.mx/pld/index.html) con la e.firma del %s (art. 24, primer párrafo, RCG).' % R,
             'Seleccionar el tipo de envío: Aviso o Informe en ceros (y, cuando se habiliten, Aviso por sospecha o por hechos o indicios).',
             'Capturar los datos del cliente, del Beneficiario Controlador y del acto u operación conforme al expediente (%s) y a la Guía de llenado vigente.' % D('7'),
             'Obtener y conservar el acuse con folio; registrarlo en las secciones III y IV y en el expediente del cliente (%s).' % D('10')]
    if d['consolidado']:
        pasos.insert(3, 'En desarrollos inmobiliarios, decidir si procede el Aviso mensual consolidado (tabla IV.1) o Avisos individuales.')
    if d['fiscal']:
        pasos.insert(1, 'Para el inciso a), si la Notaría optó por la vía fiscal, presentar la información en ese medio y conservar el acuse (sección I).')
    for t in pasos:
        p(ctx, t, bullet=True)
    p(ctx, '**Verificación mensual.** A más tardar el día 15 de cada mes, %s verifica que todos los actos u operaciones del '
           'mes anterior objeto de Aviso estén identificados en la sección III y presentados, o que se haya presentado el '
           'informe en ceros.' % ('[PF: el titular][PM: el Representante Encargado de Cumplimiento]' if kit == 'INM'
                                  else '[PF: el profesional][PM: el Representante Encargado de Cumplimiento]' if kit == 'PROF'
                                  else 'el %s, con apoyo del Encargado Operativo (D2%s, sección III),' % (R, s)))
    tabla(ctx, [('**Responsable:**', R), ('**Firma:**', '')], [3600, 5900], encabezado=False)
    pie(ctx)
    return ctx

NOMBRES = {'PROF': 'profesionistas/D12-PROF_Control_Avisos_SAT_v2.docx', 'INM': 'inmobiliaria/D12_Control_Avisos_SAT_v5.docx',
           'NOT': 'notarial/D12-NOT_Control_Avisos_SAT_v3.docx', 'COR': 'corredores/D12-COR_Control_Avisos_SAT_v2.docx'}

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
        cp.title = 'D12%s Control de Avisos e Informes — %s' % (canon.KITS[k]['suf'], canon.KITS[k]['variante'])
        cp.author = 'Certus Consultores · Espinosa-Salazar Abogados Corporativos'
        cp.subject = 'Programa Certus PLD — edición homologada noviembre 2026'
        dest = os.path.join(OUT, NOMBRES[k]); os.makedirs(os.path.dirname(dest), exist_ok=True)
        from docxlib import numero_de_pagina; numero_de_pagina(doc)
        doc.save(dest); print(k, dest)
