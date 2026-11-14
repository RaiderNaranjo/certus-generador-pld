"""D8 — Política de Personas Políticamente Expuestas (PEP's), homologada para los cuatro kits.

Uso: python3 build_d8.py <dir_contenedor> <dir_salida>   (contenedor = carpeta con PROF_pr12.docx)
Once secciones iguales en los cuatro kits; cambian los roles a evaluar (IV), los momentos, quién aprueba (VIII)
y la regla de Aviso (IX).
"""
import os, sys
sys.path.insert(0, os.path.dirname(__file__))
import build_d5 as B
from build_d5 import Ctx, p, tabla, portada, pie, escribir_par, Par
from docxlib import instalar_estilos, q, espaciar_tras_tablas
from docx import Document
import canon

SECCIONES = [
    ('I', 'Objetivo y Fundamento (art. 37 Bis, fr. IV, RCG)', 'Objetivo'),
    ('II', 'Definición de Persona Políticamente Expuesta (art. 23 Quáter RCG)', 'Definición'),
    ('III', 'Personas Asimiladas a PEP (art. 23 Quáter, tercer párrafo, RCG)', 'Asimiladas'),
    ('IV', None, 'Roles a evaluar'),
    ('V', 'Procedimiento de Identificación (arts. 23 Quáter 1, 27 y 38 RCG)', 'Procedimiento'),
    ('VI', 'Grado de Riesgo de las PEP (arts. 23 Bis 3 y 23 Bis 4 RCG)', 'Grado de Riesgo'),
    ('VII', 'Medidas de Debida Diligencia Reforzada (arts. 23 Ter 3 y 23 Ter 4 RCG)', 'Debida diligencia reforzada'),
    ('VIII', 'Aprobación o Constancia para Operar con una PEP de Riesgo Alto (art. 23 Ter 5 RCG)', 'Aprobación o constancia'),
    ('IX', 'Avisos (arts. 17 y 18, fr. VI, de la Ley)', 'Avisos'),
    ('X', 'Cuándo No Operar o Terminar la Relación (art. 37, tercer párrafo, RCG)', 'Cuándo no operar'),
    ('XI', "Registro de PEP's", 'Registro'),
]

def sec(ctx, rom, titulo=None):
    t = titulo or [s[1] for s in SECCIONES if s[0] == rom][0]
    ctx.body.append(escribir_par(Par.from_text('Sección %s: %s' % (rom, t)), 'CertusCapitulo'))

def sub(ctx, t):
    ctx.body.append(escribir_par(Par.from_text(t), 'CertusSeccion'))

def indice():
    return 'Índice: ' + ' │ '.join('%s. %s' % (r, c) for r, _, c in SECCIONES)

RESP_PFPM = {
    'INM': '[PF: el titular][PM: el Representante Encargado de Cumplimiento]',
    'PROF': '[PF: el profesional][PM: el Representante Encargado de Cumplimiento del despacho]',
}

S = {
 'INM': dict(suf='', suj='la empresa', Suj='La empresa',
    resp=RESP_PFPM['INM'], Resp='[PF: Titular][PM: Representante Encargado de Cumplimiento]',
    momento='Antes de firmar el contrato', acto='el acto u operación',
    titIV='Roles a Evaluar en el Acto u Operación (art. 17, fr. V, V Bis y XV, de la Ley)',
    introIV='En una operación inmobiliaria la verificación no se limita al cliente: se evalúa de forma independiente a cada persona que interviene en ella.',
    roles=['Fracción V (compraventa o intermediación): adquirente y enajenante del inmueble; en la intermediación, ambas partes.',
           'Fracción V Bis (desarrollo inmobiliario): cada persona que aporta recursos al desarrollo o adquiere en preventa.',
           'Fracción XV (arrendamiento): arrendatario y, en su caso, su obligado solidario o fiador.',
           'Si el cliente es persona moral o fideicomiso: el representante o apoderado que comparece y el Beneficiario Controlador (D7, Sección E).'],
    monitoreo='durante la vigencia del contrato (por ejemplo, en arrendamientos o en preventas con aportaciones periódicas)',
    aviso='El Aviso procede cuando el acto u operación alcanza el umbral de la fracción correspondiente (V y V Bis: 8,025 UMA; XV: renta mensual de 3,210 UMA; D12).',
    firma_pf='el titular', aprueba_pm='un directivo de la empresa (Director General, Administrador Único o equivalente) que consienta el acto u operación',
    pfpm=True),
 'NOT': dict(suf='-NOT', suj='la Notaría', Suj='La Notaría',
    resp='el Notario Titular', Resp='Notario Titular',
    momento='Antes de formalizar el acto', acto='el acto notarial',
    titIV='Roles a Evaluar en el Acto Notarial (art. 17, fr. XII, Apartado A, de la Ley)',
    introIV='Un acto ante la fe pública del Notario puede involucrar varios roles, más amplios que el de “cliente”; cada uno se evalúa de forma independiente.',
    roles=['Inciso a): enajenante y adquirente del inmueble, o titulares de los derechos reales que se constituyan o transmitan.',
           'Inciso b): poderdante y apoderado.',
           'Inciso c): socios fundadores, adquirentes o enajenantes de acciones o partes sociales, representante que comparece y Beneficiario Controlador (D7-NOT, Sección E).',
           'Inciso d): fideicomitente(s) y fideicomisario(s).',
           'Inciso e): mutuante o acreedor, mutuatario o deudor y, en su caso, garantes.',
           'En todos los casos: la persona que solicita materialmente la protocolización (art. 13 RCG; D7-NOT, Sección A.4).'],
    monitoreo='durante la vigencia del acto (por ejemplo, en poderes irrevocables o fideicomisos de larga duración)',
    aviso='Los actos de los incisos b), c) y e) son siempre objeto de Aviso; el inciso a), desde 8,000 UMA, y el inciso d), desde 4,000 UMA (D12-NOT).',
    pfpm=False),
 'COR': dict(suf='-COR', suj='la Correduría', Suj='La Correduría',
    resp='el Corredor Público Titular', Resp='Corredor Público Titular',
    momento='Antes de formalizar el acto', acto='el acto ante el Corredor',
    titIV='Roles a Evaluar en el Acto ante el Corredor Público (art. 17, fr. XII, Apartado B, de la Ley)',
    introIV='Un acto ante la fe pública del Corredor puede involucrar varios roles, más amplios que el de “cliente”; cada uno se evalúa de forma independiente.',
    roles=['Inciso a): solicitante del avalúo y propietario del bien avaluado.',
           'Inciso b): socios o accionistas fundadores, adquirentes y enajenantes de acciones o partes sociales, representante que comparece y Beneficiario Controlador (D7-COR, Sección E).',
           'Inciso c): fideicomitente(s), fideicomisario(s), cedente y cesionario de derechos y, en su caso, miembros del comité técnico.',
           'Inciso d): mutuante o acreditante, mutuatario o acreditado y, en su caso, garantes.',
           'En todos los casos: la persona que solicita materialmente la formalización (art. 13 RCG; D7-COR, Sección A.4).'],
    monitoreo='durante la vigencia del acto (por ejemplo, en fideicomisos y créditos de larga duración)',
    aviso='Los actos de los incisos b), c) y d) son siempre objeto de Aviso; el inciso a), desde 8,025 UMA (D12-COR).',
    pfpm=False),
 'PROF': dict(suf='-PROF', suj='[PF: el profesional][PM: el despacho]', Suj='[PF: El profesional][PM: El despacho]',
    resp=RESP_PFPM['PROF'], Resp='[PF: Profesional Titular][PM: Representante Encargado de Cumplimiento]',
    momento='Antes de preparar o ejecutar el servicio', acto='el servicio',
    titIV='Roles a Evaluar según el Servicio Prestado (art. 17, fr. XI, de la Ley)',
    introIV='Los servicios de la fracción XI pueden involucrar personas distintas del cliente que contrata; cada rol se evalúa de forma independiente.',
    roles=['Inciso a): enajenante y adquirente del inmueble o de los derechos cedidos.',
           'Inciso b): el cliente titular de los recursos, valores o activos administrados y cualquier tercero beneficiario de esa administración.',
           'Inciso c): el titular de la cuenta bancaria, de ahorro o de valores y cualquier cotitular o autorizado.',
           'Inciso d): los socios que aportan capital.',
           'Inciso e): socios, fideicomitentes, fideicomisarios y Beneficiario Controlador de la persona moral o vehículo corporativo (D7-PROF, Sección E).'],
    monitoreo='durante la vigencia del servicio (por ejemplo, en la administración continua de recursos o el manejo de cuentas)',
    aviso='El Aviso procede solo cuando el profesional realiza, en nombre y representación del cliente, una operación financiera relacionada con los incisos a) a e) (D1-PROF, sección 2.2; D12-PROF); preparar o asesorar no genera Aviso por sí mismo.',
    firma_pf='el profesional', aprueba_pm='un directivo del despacho (socio director, órgano de administración o equivalente) que consienta el servicio',
    pfpm=True),
}

CARGOS = ['Titulares del Poder Ejecutivo federal, estatal o municipal.',
          'Legisladores federales y locales (senadores y diputados).',
          'Ministros, magistrados y jueces del Poder Judicial federal y de las entidades federativas.',
          'Secretarios de Estado y sus equivalentes en los gobiernos estatales.',
          'Directores generales de organismos descentralizados y de empresas de participación estatal.',
          'Embajadores y cónsules.',
          'Altos mandos de las Fuerzas Armadas.',
          'Titulares de órganos constitucionales autónomos.',
          'Funcionarios o miembros importantes de partidos políticos y de organizaciones internacionales.']

def construir(kit):
    d = S[kit]; ctx = Ctx(kit); s = d['suf']
    D = lambda n: 'D%s%s' % (n, s)
    portada(ctx, codigo='D8', titulo="Política de Personas Políticamente Expuestas (PEP's)", indice=indice())
    # I
    sec(ctx, 'I')
    p(ctx, 'Esta Política establece los procedimientos para identificar a las Personas Políticamente Expuestas (PEP) y a '
           'las personas asimiladas a ellas, y para darles el seguimiento reforzado que exige el art. 37 Bis, fr. IV, RCG. '
           'Forma parte del Manual de Políticas Internas (%s, capítulo 6) y es obligatoria para todo el personal que atiende '
           'clientes o integra expedientes.' % D('5'))
    p(ctx, 'Vigencia de las nuevas obligaciones: la clasificación del Grado de Riesgo y las medidas de los Capítulos III Bis '
           'y III Ter RCG (arts. 23 Bis 3, 23 Bis 4 y 23 Ter 3 a 23 Ter 5) son exigibles respecto de los actos u operaciones '
           'realizados a partir del 1 de marzo de 2027 (Transitorio Cuarto del Acuerdo 115/2026); Certus PLD recomienda '
           'aplicarlas desde la adopción de esta Política.')
    # II
    sec(ctx, 'II')
    p(ctx, 'Se considera PEP a la persona física que desempeña o ha desempeñado funciones públicas en territorio nacional o '
           'en un país extranjero, entre otras: jefes de Estado o de gobierno, líderes políticos, funcionarios '
           'gubernamentales, judiciales o militares de alta jerarquía, altos ejecutivos de empresas estatales y funcionarios '
           'o miembros importantes de partidos políticos y organizaciones internacionales (art. 23 Quáter, primer párrafo, '
           'RCG). Son organizaciones internacionales las establecidas mediante acuerdos políticos oficiales entre Estados, con '
           'estatus de tratados internacionales (art. 23 Quáter, segundo párrafo, RCG).')
    p(ctx, 'La calidad de PEP nacional se conserva durante el año siguiente a aquel en que la persona dejó el cargo. Si la '
           'persona dejó de reunir esas características dentro del año inmediato anterior a la fecha del acto u operación, '
           '%s la cataloga como PEP durante el año siguiente a aquel en que se realice (art. 23 Quáter, cuarto y quinto '
           'párrafos, RCG).' % d['suj'])
    sub(ctx, 'II.1 Cargos que típicamente califican como PEP en México')
    p(ctx, 'La relación es enunciativa; ante la duda, se consulta la información que ponga a disposición la UIF (sección V).')
    for c in CARGOS:
        p(ctx, c, bullet=True)
    # III
    sec(ctx, 'III')
    p(ctx, 'Se asimilan a PEP y se tratan conforme a esta Política:')
    for c in ['El cónyuge, la concubina o el concubinario de la PEP.',
              'Las personas con quienes la PEP mantenga parentesco por consanguinidad o afinidad hasta el segundo grado (padres, hijos, abuelos, nietos y hermanos; suegros, yernos, nueras y cuñados).',
              'Los asociados o socios de personas morales con las que la PEP mantenga vínculos patrimoniales.']:
        p(ctx, c, bullet=True)
    # IV
    sec(ctx, 'IV', d['titIV'])
    p(ctx, d['introIV'])
    for r in d['roles']:
        p(ctx, r, bullet=True)
    # V
    sec(ctx, 'V')
    R = d['Resp']
    tabla(ctx, [('Paso', 'Acción', 'Responsable', 'Plazo'),
                ('1', 'Preguntar a cada persona de la sección IV si es PEP o si tiene un vínculo de los señalados en la sección III; '
                      'la respuesta se asienta en el expediente.', 'Personal que atiende al cliente', 'Primer contacto'),
                ('2', 'Para PEP nacionales, consultar la aplicación Consulta PEP 2.0 de la UIF con la e.firma utilizada en el alta '
                      '(art. 23 Quáter 1 RCG) y, mientras no esté disponible, fuentes públicas oficiales (DOF, portales de '
                      'gobierno, declaraciones patrimoniales públicas).', '%s o personal designado' % R, d['momento']),
                ('3', 'Para PEP extranjeras, consultar fuentes públicas y listas de organismos internacionales reconocidos.',
                      '%s o personal designado' % R, d['momento']),
                ('4', 'Consultar las listas que dé a conocer la UIF y las de autoridades u organismos internacionales (arts. 27 y '
                      '38 RCG; %s, capítulo 16; %s, Sección H).' % (D('5'), D('7')), '%s o personal designado' % R, d['momento']),
                ('5', 'Si alguna persona es PEP o asimilada: determinar su Grado de Riesgo (sección VI) y activar las medidas de '
                      'la sección VII.', R, 'Inmediatamente'),
                ('6', 'Documentar el resultado en el expediente KYC (%s, Sección F).' % D('7'), '%s o personal designado' % R, 'Mismo día'),
                ('7', 'Si la PEP es además de Grado de Riesgo alto: %s de la sección VIII.' % (
                      '[PF: firmar la constancia][PM: obtener la aprobación del directivo]' if d['pfpm'] else 'firmar la constancia'),
                      R, 'Antes de operar'),
                ('8', "Registrar en el Registro de PEP's (sección XI).", R, 'Mismo día')],
          [800, 5000, 2200, 1500])
    p(ctx, 'La consulta a Consulta PEP 2.0 podrá realizarse nueve meses después de la entrada en vigor del Acuerdo 115/2026 '
           '(Transitorio Décimo); hasta entonces, la verificación de PEP nacionales se apoya en fuentes públicas oficiales y '
           'se documenta qué fuente se consultó.')
    # VI
    sec(ctx, 'VI')
    p(ctx, '**PEP extranjeras.** Se clasifican siempre como de Grado de Riesgo alto, con independencia de la puntuación EBR, '
           'y %s recaba y asienta las razones por las que celebran el acto u operación en territorio mexicano (art. 23 Bis 4 '
           'RCG; %s, sección VII).' % (d['suj'], D('6')))
    p(ctx, '**PEP nacionales y asimiladas.** Se evalúan con la Metodología EBR y con factores de Riesgo adicionales, a fin '
           'de verificar que su comportamiento transaccional corresponda razonablemente con sus ingresos, funciones, nivel y '
           'responsabilidad (art. 23 Bis 3 RCG; %s, sección VIII). Si no corresponde, el Grado de Riesgo se eleva a alto.' % D('6'))
    p(ctx, 'Las PEP de Grado de Riesgo alto se reevalúan con mayor frecuencia que el resto de los clientes (art. 23 Bis 1 '
           'RCG) y se integran al sistema de alertas de los mecanismos automatizados (art. 41, fr. V, RCG; %s, capítulo 19).' % D('5'))
    # VII
    sec(ctx, 'VII')
    sub(ctx, 'VII.1 Medidas para toda PEP o persona asimilada')
    for c in ['Verificar que el valor del acto u operación sea congruente con los ingresos, funciones y nivel de responsabilidad de la PEP.',
              'Solicitar documentación adicional sobre el origen de los recursos: declaración anual de los dos últimos ejercicios y estados de cuenta de los últimos seis meses, o la que resulte pertinente al acto u operación.',
              'Monitoreo reforzado %s.' % d['monitoreo'],
              'Registrar en el expediente KYC (%s) todas las medidas aplicadas y su resultado.' % D('7')]:
        p(ctx, c, bullet=True)
    sub(ctx, 'VII.2 Medidas adicionales cuando la PEP sea de Grado de Riesgo alto')
    p(ctx, 'Además de lo anterior, %s aplica las medidas que la RCG exige para todo cliente de Grado de Riesgo alto:' % d['suj'])
    p(ctx, 'Requerir mayor información sobre su actividad preponderante, realizar un monitoreo más estricto de su '
           'comportamiento transaccional y aplicar el Cuestionario de Identificación Reforzada sobre el origen y destino de los '
           'recursos (art. 23 Ter 3 RCG; %s, Anexo B). El cuestionario puede aplicarse a distancia, con la Firma Electrónica '
           'de quien lo suscribe.' % D('5'), bullet=True)
    p(ctx, 'Persona física: adoptar medidas reforzadas para conocer el origen y destino de los recursos y obtener los datos '
           'del Capítulo III RCG de su cónyuge y dependientes económicos, así como de las sociedades y asociaciones con las '
           'que mantenga vínculos patrimoniales (art. 23 Ter 4, fr. I, RCG).', bullet=True)
    p(ctx, 'Persona moral: obtener mayor información de sus principales accionistas o socios y confirmarla en los registros '
           'electrónicos de la Secretaría de Economía (art. 23 Ter 4, fr. II, RCG).', bullet=True)
    p(ctx, 'PEP extranjera: obtener, además, la documentación del Capítulo III RCG de su cónyuge y dependientes económicos '
           '(art. 23 Ter 4, fr. III, RCG).', bullet=True)
    # VIII
    sec(ctx, 'VIII')
    if d['pfpm']:
        p(ctx, 'Cuando la persona que pretende ser cliente, o que ya lo es, sea PEP y además de Grado de Riesgo alto, el acto u '
               'operación requiere, antes de llevarse a cabo, [PF: una constancia firmada por %s en la que señale los '
               'motivos que consideró para realizarlo, documentada en el expediente KYC][PM: la aprobación de %s] (art. 23 Ter 5 '
               'RCG). Si la condición se detecta después del acto u operación, la [PF: constancia][PM: aprobación] se recaba en '
               'cuanto se detecte.' % (d['firma_pf'], d['aprueba_pm']))
        etiqueta = '[PF: Titular que firma la constancia:][PM: Directivo que aprueba (nombre y cargo):]'
    else:
        p(ctx, 'Cuando la persona que pretende ser cliente, o que ya lo es, sea PEP y además de Grado de Riesgo alto, el %s, '
               'como persona física, firma antes de formalizar el acto una constancia en la que señala los motivos que '
               'consideró para realizarlo, y la documenta en el expediente KYC (art. 23 Ter 5, segundo párrafo, RCG). Si la '
               'condición se detecta después del acto, la constancia se firma en cuanto se detecte.' % d['Resp'])
        etiqueta = '%s que firma la constancia:' % d['Resp']
    tabla(ctx, [('**Persona PEP y rol en el acto u operación:**', ''),
                ('**Motivos considerados para realizarlo:**', ''),
                ('**%s**' % etiqueta if not etiqueta.startswith('[') else etiqueta, ''),
                ('**Fecha:**', '[dd/mm/aaaa]'),
                ('**Firma:**', '')], [3600, 5900], encabezado=False)
    # IX
    sec(ctx, 'IX')
    p(ctx, 'La condición de PEP no crea, por sí sola, la obligación de presentar Aviso. %s' % d['aviso'])
    p(ctx, 'Con independencia del monto, %s presenta Aviso dentro de las 24 horas siguientes cuando advierta indicios o '
           'hechos de que los recursos podrían estar vinculados con operaciones con recursos de procedencia ilícita (art. 18, '
           'fr. VI, segundo párrafo, de la Ley; arts. 26 Bis y 26 Bis 1 RCG; %s), o cuando la persona figure en la lista a que '
           'se refiere el art. 38 RCG (art. 27 RCG).' % (d['suj'], D('9')))
    # X
    sec(ctx, 'X')
    p(ctx, '%s no realiza %s, o termina la relación, cuando (%s, capítulo 18):' % (d['Suj'], d['acto'], D('5')))
    for c in ['No sea posible verificar el origen de los recursos.',
              'El acto u operación presente múltiples señales de alerta sin justificación.',
              'El valor del acto u operación sea notoriamente desproporcionado respecto de los ingresos, funciones o nivel de la PEP.',
              'La persona se niegue a proporcionar la información o documentación adicional solicitada.',
              ('No se haya obtenido la [PF: constancia][PM: aprobación] de la sección VIII para una PEP de Grado de Riesgo alto.'
               if d['pfpm'] else 'No se haya firmado la constancia de la sección VIII para una PEP de Grado de Riesgo alto.')]:
        p(ctx, c, bullet=True)
    # XI
    sec(ctx, 'XI')
    filas = [('No.', 'Nombre', 'Rol en el acto', 'Cargo o vínculo PEP', 'Fecha', 'Grado de Riesgo', 'Resultado')]
    for i in range(1, 4):
        filas.append((str(i), '', '', '', '', '[ ] Bajo [ ] Medio [ ] Alto', '[ ] Operó [ ] No operó'))
    tabla(ctx, filas, [500, 1700, 1300, 1600, 1000, 1700, 1700])
    p(ctx, '')
    tabla(ctx, [('**Aprobó la Política (nombre y cargo):**', ''), ('**Fecha de adopción:**', '[dd/mm/aaaa]'), ('**Firma:**', '')],
          [3600, 5900], encabezado=False)
    pie(ctx)
    return ctx

NOMBRES = {'PROF': 'profesionistas/D8-PROF_Politica_PEPs_v2.docx', 'INM': 'inmobiliaria/D8_Politica_PEPs_v5.docx',
           'NOT': 'notarial/D8-NOT_Politica_PEPs_v3.docx', 'COR': 'corredores/D8-COR_Politica_PEPs_v2.docx'}

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
        cp.title = "D8%s Política de Personas Políticamente Expuestas (PEP's) — %s" % (canon.KITS[k]['suf'], canon.KITS[k]['variante'])
        cp.author = 'Certus Consultores · Espinosa-Salazar Abogados Corporativos'
        cp.subject = 'Programa Certus PLD — edición homologada noviembre 2026'
        dest = os.path.join(OUT, NOMBRES[k]); os.makedirs(os.path.dirname(dest), exist_ok=True)
        from docxlib import numero_de_pagina; numero_de_pagina(doc)
        doc.save(dest); print(k, dest)
