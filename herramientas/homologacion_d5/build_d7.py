"""D7 — Expediente de Identificación y Conocimiento del Cliente (KYC), homologado para los cuatro kits.

Uso: python3 build_d7.py <dir_contenedor> <dir_salida>
Secciones A–J iguales en los cuatro; solo cambian A.4 (fedatarios), I (acto u operación) y una fila de J.
"""
import os, sys
sys.path.insert(0, os.path.dirname(__file__))
from lxml import etree
import build_d5 as B
from build_d5 import Ctx, p, tabla, pie, escribir_par, Par
from docxlib import instalar_estilos, q, W, espaciar_tras_tablas, tabla_simple, normalizar_tabla
from build_d5 import aplicar_globales
from docx import Document
import canon

SECC = [('A', 'Identificación del Cliente o Usuaria (art. 12 y Anexos 3, 4 y 8 RCG)', 'Identificación del cliente'),
        ('B', 'Identificación Oficial y Cotejo (art. 12 RCG)', 'Identificación oficial y cotejo'),
        ('C', 'Origen y Destino de los Recursos', 'Origen y destino de los recursos'),
        ('D', 'Evaluación EBR y Grado de Riesgo (D6; arts. 23 Bis a 23 Bis 4 RCG)', 'Evaluación EBR'),
        ('E', 'Beneficiario Controlador (art. 18, fr. III, de la Ley; arts. 23 Quinquies a 23 Quinquies 2 RCG)', 'Beneficiario Controlador'),
        ('F', 'Personas Políticamente Expuestas (arts. 23 Quáter y 23 Ter 5 RCG)', 'PEP'),
        ('G', 'Perfil Transaccional (arts. 23 Ter 1 y 23 Ter 2 RCG)', 'Perfil transaccional'),
        ('H', 'Consulta de Listas de Autoridades (arts. 27 y 38 RCG)', 'Listas de autoridades'),
        ('I', None, 'Acto u operación'),
        ('J', 'Documentos Integrados y Validación (arts. 12 y 21 RCG)', 'Documentos y validación')]

def sec(ctx, letra, titulo=None):
    t = titulo or [s[1] for s in SECC if s[0] == letra][0]
    ctx.body.append(escribir_par(Par.from_text('Sección %s: %s' % (letra, t)), 'CertusCapitulo'))

def sub(ctx, t):
    ctx.body.append(escribir_par(Par.from_text(t), 'CertusSeccion'))

def ficha(ctx, filas):
    tabla(ctx, [('**%s**' % a if a else '', b) for a, b in filas], [3600, 5900], encabezado=False)

SI_NO = '[ ] Sí   [ ] No'
S = {
 'INM': dict(suf='', variante='Inmobiliarias', banda='Empresa: [EMPRESA]',
    actos='[ ] Compraventa o intermediación (fr. V)   [ ] Preventa o aportación a desarrollo (fr. V Bis)   [ ] Arrendamiento (fr. XV)',
    titI='Acto u Operación Celebrada (art. 17, fr. V, V Bis y XV, de la Ley)',
    I=[('Tipo de acto u operación:', '[ ] Compraventa   [ ] Intermediación   [ ] Preventa o aportación (V Bis)   [ ] Arrendamiento'),
       ('Inmueble o desarrollo (descripción y ubicación):', ''),
       ('Valor de la operación o renta mensual:', '$ ______________ MXN'),
       ('Forma de pago:', '[ ] Transferencia   [ ] Cheque   [ ] Crédito de Entidad Financiera   [ ] Efectivo (prohibido si el inmueble vale 8,025 UMA o más, o si la renta mensual es de 3,210 UMA o más; art. 32, fr. I y VII, de la Ley)'),
       ('Objeto de Aviso (D12):', '[ ] Sí: alcanza el umbral (V y V Bis: 8,025 UMA; XV: 3,210 UMA mensuales) — Fecha del Aviso: ____   [ ] No')],
    subsI=[], doc='Contrato celebrado', A4=None, valida='Representante Encargado de Cumplimiento o titular'),
 'NOT': dict(suf='-NOT', variante='Notarías', banda='Notaría: [Denominación completa]',
    actos='[ ] a) Inmuebles   [ ] b) Poder irrevocable   [ ] c) Sociedades   [ ] d) Fideicomiso   [ ] e) Mutuo o crédito',
    titI='Acto Notarial Celebrado (art. 17, fr. XII, Apartado A, de la Ley)',
    I=[('No. de instrumento:', '[Número]'),
       ('Inciso aplicable:', '[ ] a) Transmisión o constitución de derechos reales sobre inmuebles   [ ] b) Poder irrevocable   [ ] c) Sociedad, capital, fusión, escisión o acciones   [ ] d) Fideicomiso traslativo o de garantía   [ ] e) Mutuo o crédito'),
       ('Valor del acto:', '$ ______________ MXN'),
       ('Forma de pago (art. 33 de la Ley):', '[ ] Transferencia   [ ] Cheque   [ ] Efectivo (prohibido en inmuebles de 8,025 UMA o más; art. 32, fr. I)   [ ] Pago previo, con declaración bajo protesta (art. 33, segundo párrafo)'),
       ('Objeto de Aviso (D12-NOT):', '[ ] Sí — Fecha del Aviso: ____   [ ] No (los incisos b), c) y e) son siempre objeto de Aviso; a) desde 8,000 UMA; d) desde 4,000 UMA)')],
    subsI=[('I.1 Inciso a): transmisión o constitución de derechos reales sobre inmuebles',
            [('Inmueble (descripción y ubicación):', ''), ('Valor catastral, comercial o precio pactado (el mayor):', '$ ______________ MXN'), ('Enajenante:', ''), ('Adquirente:', '')]),
           ('I.2 Inciso b): otorgamiento de poder con carácter irrevocable',
            [('Poderdante:', ''), ('Apoderado:', ''), ('Facultades otorgadas:', '[ ] Administración   [ ] Dominio   [ ] Pleitos y cobranzas'), ('Justificación de la irrevocabilidad:', '')]),
           ('I.3 Inciso c): constitución, modificación patrimonial, fusión, escisión o compraventa de acciones o partes sociales',
            [('Denominación de la sociedad:', ''), ('Socios fundadores o adquirentes (relacionar todos):', ''), ('Monto de capital social (constitución o variación):', '$ ______________ MXN'), ('Beneficiario Controlador identificado:', '[ ] Sí (Sección E)   [ ] Pendiente')]),
           ('I.4 Inciso d): fideicomiso traslativo de dominio o de garantía',
            [('Fideicomitente:', ''), ('Institución fiduciaria:', ''), ('Fideicomisario(s):', ''), ('Bien(es) fideicomitido(s) y valor:', '$ ______________ MXN')]),
           ('I.5 Inciso e): contrato de mutuo o crédito',
            [('Mutuante o acreedor:', ''), ('Mutuatario o deudor:', ''), ('Monto del mutuo o crédito:', '$ ______________ MXN'), ('Con garantía:', '[ ] Sí — Tipo: ____   [ ] No')])],
    doc='Instrumento notarial (escritura o acta)', A4='protocolización', valida='Notario Titular'),
 'COR': dict(suf='-COR', variante='Corredurías Públicas', banda='Correduría: [Denominación de la Correduría]',
    actos='[ ] a) Avalúo con fe pública   [ ] b) Sociedades   [ ] c) Fideicomiso   [ ] d) Mutuo o crédito mercantil',
    titI='Acto Formalizado ante el Corredor (art. 17, fr. XII, Apartado B, de la Ley)',
    I=[('No. de póliza o acta:', '[Número]'),
       ('Inciso aplicable:', '[ ] a) Avalúo con fe pública   [ ] b) Sociedad, capital, fusión, escisión o acciones   [ ] c) Fideicomiso   [ ] d) Mutuo o crédito mercantil'),
       ('Valor del acto:', '$ ______________ MXN'),
       ('Forma de pago (art. 33 de la Ley):', '[ ] Transferencia   [ ] Cheque   [ ] Efectivo (prohibido en acciones o partes sociales de 3,210 UMA o más; art. 32, fr. VI)   [ ] Pago previo, con declaración bajo protesta (art. 33, segundo párrafo)'),
       ('Objeto de Aviso (D12-COR):', '[ ] Sí — Fecha del Aviso: ____   [ ] No (los incisos b), c) y d) son siempre objeto de Aviso; a) desde 8,025 UMA)')],
    subsI=[('I.1 Inciso a): avalúo realizado con fe pública',
            [('Solicitante del avalúo:', ''), ('Bien avaluado (descripción y ubicación):', ''), ('Valor conclusivo del avalúo:', '$ ______________ MXN'), ('¿Se utilizó la fe pública? (art. 29 del Reglamento):', '[ ] Sí — inciso a)   [ ] No — avalúo sin fe pública (fuera del inciso a)'), ('Finalidad del avalúo:', '')]),
           ('I.2 Inciso b): constitución, aumento o disminución de capital, fusión, escisión o compraventa de acciones o partes sociales',
            [('Denominación de la sociedad mercantil:', ''), ('Tipo de acto:', '[ ] Constitución   [ ] Aumento de capital   [ ] Disminución de capital   [ ] Fusión   [ ] Escisión   [ ] Compraventa de acciones o partes sociales'), ('Socios, accionistas, adquirentes o enajenantes (relacionar todos):', ''), ('Monto (capital social, variación o precio):', '$ ______________ MXN'), ('Beneficiario Controlador identificado:', '[ ] Sí (Sección E)   [ ] Pendiente')]),
           ('I.3 Inciso c): constitución, modificación o cesión de derechos de fideicomiso',
            [('Tipo de acto:', '[ ] Constitución   [ ] Modificación   [ ] Cesión de derechos'), ('Fideicomitente(s):', ''), ('Institución fiduciaria:', ''), ('Fideicomisario(s) o cesionario(s):', ''), ('Bienes o derechos fideicomitidos y valor:', '$ ______________ MXN'), ('¿Garantiza un crédito a favor de una institución del sistema financiero?:', '[ ] Sí — excluido del inciso c)   [ ] No')]),
           ('I.4 Inciso d): mutuo o crédito mercantil',
            [('Mutuante o acreditante:', ''), ('Mutuatario o acreditado:', ''), ('Monto del mutuo o crédito:', '$ ______________ MXN'), ('Con garantía:', '[ ] Sí — Tipo: ____   [ ] No'), ('¿El acreedor forma parte del sistema financiero?:', '[ ] Sí — excluido del inciso d)   [ ] No')])],
    doc='Póliza o acta (asiento en el libro de registro)', A4='formalización', valida='Corredor Público Titular'),
 'PROF': dict(suf='-PROF', variante='Servicios Profesionales', banda='Profesional / Despacho: [Nombre]',
    actos='[ ] a) Inmuebles   [ ] b) Administración de recursos   [ ] c) Cuentas   [ ] d) Aportaciones de capital   [ ] e) Personas morales o vehículos corporativos',
    titI='Servicio Prestado (art. 17, fr. XI, incisos a) a e), de la Ley)',
    I=[('Inciso aplicable:', '[ ] a) Compraventa o cesión de derechos sobre inmuebles   [ ] b) Administración y manejo de recursos o activos   [ ] c) Manejo de cuentas bancarias, de ahorro o de valores   [ ] d) Organización de aportaciones de capital   [ ] e) Constitución, escisión, fusión, operación o administración de personas morales o vehículos corporativos'),
       ('Momento de sujeción (D1-PROF, sección 2.2):', '[ ] Solo preparación o asesoría (identificación obligatoria)   [ ] Operación financiera en nombre y representación del cliente (identificación y Aviso)'),
       ('Valor de la operación o del activo:', '$ ______________ MXN (indicador de Riesgo; la fracción XI no tiene umbral)'),
       ('Forma de pago:', '[ ] Transferencia   [ ] Cheque   [ ] Efectivo (prohibido en inmuebles de 8,025 UMA o más y en acciones o partes sociales de 3,210 UMA o más; art. 32, fr. I y VI)'),
       ('Objeto de Aviso (D12-PROF):', '[ ] Sí: operación financiera en nombre y representación del cliente — Fecha del Aviso: ____   [ ] No (el Aviso de 24 horas por sospecha se registra en el D9-PROF)')],
    subsI=[('I.1 Inciso a): compraventa de inmuebles o cesión de derechos sobre estos',
            [('Inmueble (descripción y ubicación):', ''), ('Enajenante:', ''), ('Adquirente:', '')]),
           ('I.2 Inciso b): administración y manejo de recursos, valores o activos del cliente',
            [('Recursos o activos administrados (descripción):', ''), ('Institución(es) donde se depositan:', ''), ('Facultades del profesional sobre los recursos:', '')]),
           ('I.3 Inciso c): manejo de cuentas bancarias, de ahorro o de valores',
            [('Institución financiera:', ''), ('Cuenta (últimos 4 dígitos):', ''), ('Facultades otorgadas al profesional:', '[ ] Firma mancomunada   [ ] Firma individual   [ ] Solo consulta')]),
           ('I.4 Inciso d): organización de aportaciones de capital para sociedades',
            [('Denominación de la sociedad:', ''), ('Socios aportantes (relacionar todos):', ''), ('Monto total de las aportaciones organizadas:', '$ ______________ MXN')]),
           ('I.5 Inciso e): personas morales o vehículos corporativos (incluido el fideicomiso) y compraventa de entidades mercantiles',
            [('Denominación de la persona moral o vehículo:', ''), ('Tipo de operación:', '[ ] Constitución   [ ] Escisión   [ ] Fusión   [ ] Operación o administración   [ ] Fideicomiso   [ ] Compraventa de entidad mercantil'), ('Socios, fideicomitentes o fideicomisarios (relacionar todos):', ''), ('Beneficiario Controlador identificado:', '[ ] Sí (Sección E)   [ ] Pendiente'), ('Monto de capital social o valor de la operación:', '$ ______________ MXN')])],
    doc='Instrumento o contrato que documenta el servicio', A4=None, valida='Representante Encargado de Cumplimiento o Profesional Titular'),
}

def banda(ctx, info, d):
    tbl = etree.fromstring(
        '<w:tbl xmlns:w="%s"><w:tblPr><w:tblW w:w="9500" w:type="dxa"/><w:jc w:val="center"/>'
        '<w:tblBorders><w:top w:val="single" w:sz="12" w:space="0" w:color="1F3864"/><w:left w:val="nil"/>'
        '<w:bottom w:val="single" w:sz="12" w:space="0" w:color="1F3864"/><w:right w:val="nil"/>'
        '<w:insideH w:val="nil"/><w:insideV w:val="nil"/></w:tblBorders><w:tblLayout w:type="fixed"/>'
        '<w:tblCellMar><w:top w:w="100" w:type="dxa"/><w:left w:w="120" w:type="dxa"/><w:bottom w:w="100" w:type="dxa"/>'
        '<w:right w:w="120" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="9500"/></w:tblGrid>'
        '<w:tr><w:tc><w:tcPr><w:tcW w:w="9500" w:type="dxa"/></w:tcPr></w:tc></w:tr></w:tbl>' % W)
    tc = tbl.find('.//' + q('w:tc'))
    e1 = escribir_par(Par.from_text('**CERTUS PLD   │   D7%s Expediente de Identificación y Conocimiento del Cliente (KYC) — %s**' % (info['suf'], info['variante'])), 'CertusTabla')
    for r in e1.iter(q('w:r')):
        rpr = r.find(q('w:rPr'))
        etree.SubElement(rpr, q('w:color')).set(q('w:val'), '1F3864')
        etree.SubElement(rpr, q('w:sz')).set(q('w:val'), '26'); etree.SubElement(rpr, q('w:szCs')).set(q('w:val'), '26')
    tc.append(e1)
    tc.append(escribir_par(Par.from_text('%s   │   Variante: %s — %s' % (d['banda'], info['variante'], info['frac'])), 'CertusTabla'))
    ctx.body.append(tbl)

def construir(kit):
    d = S[kit]; info = canon.KITS[kit]; ctx = Ctx(kit); s = d['suf']
    banda(ctx, info, d)
    p(ctx, '')
    p(ctx, 'Expediente único de identificación (art. 12 RCG). Se integra antes del acto u operación o, a más tardar, al '
           'establecer la Relación de negocios, y se conserva por al menos diez años (art. 18, fr. IV, de la Ley). Se '
           'llenan solo los apartados que correspondan al tipo de cliente y al acto u operación.')
    ficha(ctx, [('No. de expediente:', '[Número secuencial]'),
                ('Tipo de cliente:', '[ ] Persona física   [ ] Persona moral   [ ] Fideicomiso u otra figura jurídica'),
                ('Fecha de apertura:', '[dd/mm/aaaa]'),
                ('Acto u operación:', d['actos']),
                ('Grado de Riesgo (D6%s):' % s, '[ ] BAJO   [ ] MEDIO   [ ] ALTO'),
                ('Próxima reevaluación (al menos semestral):', '[dd/mm/aaaa]')])
    # A
    sec(ctx, 'A')
    sub(ctx, 'A.1 Persona física (Anexos 3 y 5 RCG)')
    ficha(ctx, [('Nombre completo (sin abreviaturas):', ''), ('CURP:', ''), ('RFC (si cuenta con él):', ''),
                ('Fecha de nacimiento:', ''), ('País de nacimiento y nacionalidad:', ''),
                ('Estado civil:', '[ ] Soltero   [ ] Casado   [ ] Divorciado   [ ] Unión libre'),
                ('Actividad, ocupación o profesión:', ''), ('Domicilio:', ''), ('Teléfono:', ''), ('Correo electrónico:', '')])
    sub(ctx, 'A.2 Persona moral (Anexos 4 y 6 RCG)')
    ficha(ctx, [('Denominación o razón social:', ''), ('RFC:', ''), ('Giro o actividad:', ''), ('Fecha de constitución:', ''),
                ('Instrumento constitutivo (número, fecha y fedatario):', ''), ('Domicilio:', ''),
                ('Representante legal o apoderado que comparece:', ''), ('Instrumento que acredita sus facultades:', ''),
                ('Teléfono:', ''), ('Correo electrónico:', '')])
    sub(ctx, 'A.3 Fideicomiso u otra figura jurídica (Anexo 8 RCG)')
    ficha(ctx, [('Número o referencia del fideicomiso o figura:', ''), ('Fiduciario (denominación y RFC):', ''),
                ('Fideicomitente(s):', ''), ('Fideicomisario(s):', ''), ('¿Fideicomisarios determinados?:', SI_NO),
                ('Delegado fiduciario o apoderado que comparece:', '')])
    p(ctx, 'El Beneficiario Controlador del fideicomiso o figura se identifica en la Sección E.')
    if d['A4']:
        sub(ctx, 'A.4 Persona que solicita materialmente la %s (art. 13 RCG)' % d['A4'])
        ficha(ctx, [('Nombre completo:', ''), ('Fecha de nacimiento:', ''),
                    ('Demás datos del Anexo 3, inciso a), numerales ix) y x), RCG:', ''),
                    ('¿Es una de las partes del acto?:', '[ ] Sí   [ ] No — relación con las partes: ____')])
    # B
    sec(ctx, 'B')
    ficha(ctx, [('Tipo de documento:', '[ ] Credencial para votar   [ ] Pasaporte   [ ] Documento migratorio (extranjero)   [ ] Otra identificación oficial vigente con fotografía y firma: ____ (no se acepta la cédula profesional)'),
                ('No. de documento:', ''), ('Vigencia:', ''),
                ('Comprobante de domicilio (antigüedad no mayor a tres meses, cuando se requiera):', SI_NO),
                ('Cotejado contra original o copia certificada por:', ''), ('Fecha del cotejo:', '')])
    # C
    sec(ctx, 'C')
    ficha(ctx, [('Fuente principal de los recursos:', '[ ] Nómina   [ ] Actividad empresarial   [ ] Inversiones   [ ] Herencia   [ ] Bien propio aportado   [ ] Otro: ____'),
                ('Documentos de soporte:', '[ ] Declaración fiscal   [ ] Estados de cuenta   [ ] Contrato laboral   [ ] Escritura o título antecedente   [ ] Otro: ____'),
                ('¿Recursos propios?:', '[ ] Sí   [ ] No — fuente: ____'),
                ('Crédito de Entidad Financiera:', '[ ] Sí — institución: ____   [ ] No'),
                ('Cuenta(s) de origen (institución y titular):', ''),
                ('Destino de los recursos o del bien:', '')])
    # D
    sec(ctx, 'D')
    ficha(ctx, [('Puntuación EBR (ficha de la sección XI del D6%s):' % s, '_____ / 21 puntos'),
                ('Grado de Riesgo:', '[ ] BAJO (7–11)   [ ] MEDIO (12–16)   [ ] ALTO (17–21)'),
                ('Riesgo alto obligatorio (D6%s, sección VII; art. 23 Bis 4 RCG):' % s, '[ ] No aplica   [ ] No residente vinculado con jurisdicción de Riesgo   [ ] PEP extranjera — razones por las que opera en México: ____'),
                ('Medidas de debida diligencia:', '[ ] Simplificadas (art. 17 RCG)   [ ] Completas (art. 17 Bis RCG)   [ ] Reforzadas (arts. 23 Ter 3 y 23 Ter 4 RCG)'),
                ('Cuestionario de Identificación Reforzada (D5%s, Anexo B):' % s, '[ ] Aplicado   [ ] No aplica'),
                ('Fundamento de la clasificación:', '')])
    # E
    sec(ctx, 'E')
    p(ctx, 'Si el cliente es persona moral, fideicomiso u otra figura jurídica, se recaban los documentos que permitan '
           'identificar a su Beneficiario Controlador con el orden de prelación del art. 23 Quinquies RCG; en fideicomisos se '
           'asciende en la cadena de titularidad hasta la persona física que ejerza el control efectivo (art. 23 Quinquies 1 '
           'RCG). Si es persona física, se recaba su declaración sobre si tiene o no conocimiento de la existencia de un '
           'Beneficiario Controlador (art. 18, fr. III, de la Ley; Anexos 3 y 5 RCG).')
    ficha(ctx, [('Constancia firmada por el cliente sobre el Beneficiario Controlador:', '[ ] Autógrafa   [ ] Firma Electrónica   [ ] Pendiente'),
                ('¿El cliente declara ser el Beneficiario Controlador?:', SI_NO),
                ('Criterio de identificación (art. 23 Quinquies RCG):', '[ ] 25% o más del capital social   [ ] Control por otros medios   [ ] Funcionario administrativo de mayor grado'),
                ('Nombre completo del Beneficiario Controlador:', ''), ('Nacionalidad:', ''), ('CURP o RFC:', ''),
                ('Porcentaje de participación o naturaleza del control:', ''),
                ('Excepción (art. 23 Quinquies 2 RCG):', '[ ] No   [ ] Cotiza en bolsa (clave de pizarra: ____)   [ ] Persona moral de los Anexos 4 Bis, 6 Bis, 7-A o 7 Bis-A'),
                ('Documentación que soporta la identificación:', '')])
    # F
    sec(ctx, 'F')
    ficha(ctx, [('El cliente es PEP:', '[ ] Sí, nacional   [ ] Sí, extranjera   [ ] No'),
                ('El Beneficiario Controlador, socio o apoderado es PEP:', SI_NO),
                ('Cónyuge, concubina o concubinario, pariente hasta el segundo grado o socio con vínculo patrimonial de una PEP:', SI_NO),
                ('Cargo o vínculo y, en su caso, fecha en que dejó el cargo:', ''),
                ('Consulta PEP 2.0 de la UIF (art. 23 Quáter 1 RCG):', '[ ] Realizada — fecha: ____   [ ] No disponible o no realizada — motivo: ____'),
                ('PEP de Riesgo alto: aprobación del directivo (persona moral) o constancia del titular (persona física), art. 23 Ter 5 RCG:', '[ ] Sí — nombre y firma: ____   [ ] No aplica')])
    p(ctx, 'Ver Documento D8%s — Política de PEP\'s.' % s)
    # G
    sec(ctx, 'G')
    p(ctx, 'Se determina con la información que proporciona el cliente al alta; se reevalúa al menos cada seis meses y '
           'alimenta el sistema de alertas (arts. 23 Ter 1 y 23 Ter 2 RCG; D5%s, capítulo 14).' % s)
    ficha(ctx, [('Monto máximo mensual estimado:', '$ ______________ MXN'), ('Número y frecuencia esperada de actos u operaciones:', ''),
                ('Zona o área geográfica:', ''), ('Origen y destino esperado de los recursos:', ''), ('Actividad económica:', ''),
                ('¿El acto u operación es consistente con el Perfil Transaccional?:', '[ ] Sí   [ ] No — generar alerta (D5%s, sección 14.2)' % s),
                ('Fecha de la última reevaluación:', '[dd/mm/aaaa]')])
    # H
    sec(ctx, 'H')
    ficha(ctx, [('Fecha y hora de la consulta:', ''),
                ('Personas consultadas:', '[ ] Cliente   [ ] Beneficiario Controlador   [ ] Representante o apoderado   [ ] Otras partes'),
                ('Fuentes consultadas:', '[ ] Lista de Personas Bloqueadas (UIF)   [ ] Consejo de Seguridad de la ONU   [ ] OFAC   [ ] SAT, arts. 69 y 69-B CFF   [ ] Otra: ____'),
                ('Resultado:', '[ ] Sin coincidencia   [ ] Coincidencia descartada (homónimo, documentado)   [ ] Coincidencia confirmada: no operar y presentar Aviso dentro de 24 horas (art. 27 RCG)'),
                ('Consultó:', '')])
    # I
    sec(ctx, 'I', d['titI'])
    ficha(ctx, d['I'])
    for t, filas in d['subsI']:
        sub(ctx, t)
        ficha(ctx, filas)
    # J
    sec(ctx, 'J')
    tabla(ctx, [('Documento', 'Integrado', 'Fecha de obtención'),
                ('Identificación oficial (copia legible y cotejada)', SI_NO, ''),
                ('Comprobante de domicilio (si se requiere)', '[ ] Sí   [ ] No   [ ] No aplica', ''),
                ('RFC o Constancia de Situación Fiscal; CURP', SI_NO, ''),
                ('Documentos de la persona moral o fideicomiso (Anexos 4, 6 u 8 RCG)', '[ ] Sí   [ ] No   [ ] No aplica', ''),
                (d['doc'], SI_NO, ''),
                ('Soporte del origen y destino de los recursos', SI_NO, ''),
                ('Constancia y documentación del Beneficiario Controlador (Sección E)', SI_NO, ''),
                ('Constancia de consulta de listas (Sección H)', SI_NO, ''),
                ('Ficha de Evaluación EBR firmada (D6%s)' % s, SI_NO, ''),
                ('Cuestionario de Identificación Reforzada (Riesgo alto)', '[ ] Sí   [ ] No   [ ] No aplica', ''),
                ('Acuse del Aviso (si aplica)', '[ ] Sí   [ ] No   [ ] No aplica', '')],
          [5200, 2300, 2000])
    p(ctx, '')
    ficha(ctx, [('Integró el expediente (nombre y cargo):', ''), ('Fecha:', ''),
                ('Validó (%s):' % d['valida'], ''), ('Fecha de validación:', ''),
                ('Verificación anual del expediente (art. 21 RCG):', '[dd/mm/aaaa] — [ ] Completo y actualizado   [ ] Requiere actualización')])
    pie(ctx)
    return ctx

NOMBRES = {'PROF': 'profesionistas/D7-PROF_Expediente_KYC_v2.docx', 'INM': 'inmobiliaria/D7_Expediente_KYC_v5.docx',
           'NOT': 'notarial/D7-NOT_Expediente_KYC_v3.docx', 'COR': 'corredores/D7-COR_Expediente_KYC_v2.docx'}

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
        cp.title = 'D7%s Expediente de Identificación y Conocimiento del Cliente (KYC) — %s' % (canon.KITS[k]['suf'], canon.KITS[k]['variante'])
        cp.author = 'Certus Consultores · Espinosa-Salazar Abogados Corporativos'
        cp.subject = 'Programa Certus PLD — edición homologada noviembre 2026'
        dest = os.path.join(OUT, NOMBRES[k]); os.makedirs(os.path.dirname(dest), exist_ok=True)
        doc.save(dest); print(k, dest)
