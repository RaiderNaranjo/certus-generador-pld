"""Papeles de trabajo y documentos internos de auditoría: referencias al día, y PT-PROF nuevo (desde el PT-COR)."""
import os, sys, glob
sys.path.insert(0, os.path.dirname(__file__))
from restructura import Doc
from fase2_d0a import aplicar

REPO = '/home/claude/certus-generador-pld/plantillas'

COMUN = [(r'\bInforme de Auditoría PLD\b', 'Dictamen de Auditoría PLD/FT'), (r'\bInforme de Auditoría\b', 'Dictamen de Auditoría PLD/FT'),
         (r'\bdel RC\b', 'del Representante Encargado de Cumplimiento'), (r'\bal RC\b', 'al Representante Encargado de Cumplimiento'),
         (r'\bel RC\b', 'el Representante Encargado de Cumplimiento'), (r'\[PM: RC\]', '[PM: Representante Encargado de Cumplimiento]'),
         (r'\bRC\b(?!_)', 'Representante Encargado de Cumplimiento'),
         (r'REPRESENTANTE DE CUMPLIMIENTO', 'REPRESENTANTE ENCARGADO DE CUMPLIMIENTO'),
         (r'mínimo aprobatorio 14/20', 'mínimo aprobatorio 17/24'),
         (r'los 6 factores del (D6[-A-Z]*)', r'los 7 indicadores del \1 (escala 7–21)'),
         (r'la revisión de los 1[23] módulos conforme al (D13[-A-Z]*)', r'la revisión de las obligaciones de la sección VI del \1'),
         (r'(EBR Rojo)', 'Grado de Riesgo alto')]
SECT = {'NOT': [(r'matriz de 14 rubros', 'matriz de 22 rubros')],
        'COR': [(r'matriz de 15 rubros', 'matriz de 22 rubros'), (r'Sección A\.3 \(solicitante material\)', 'Sección A.4 (solicitante material)'),
                (r'secciones A\.3, E y F', 'secciones A.4, E y F')],
        'INM': [(r'^SEC A Identif\.$', 'SEC A-B Identif. y cotejo'), (r'^SEC B Origen$', 'SEC C Origen y destino'),
                (r'^SEC C Fondos$', 'SEC E Benef. Controlador'), (r'^Expedientes con Sección A completa$', 'Expedientes con Secciones A y B completas'),
                (r'^Expedientes con Sección B completa$', 'Expedientes con Sección C completa (origen y destino)'),
                (r'^Expedientes con Sección C completa \(origen de fondos\)$', 'Expedientes con Sección E completa (Beneficiario Controlador)'),
                (r'^AUDITOR EXTERNO$', 'AUDITOR')]}

PROF = [(r'Corredurías Públicas — Fr\. XII, Apartado B', 'Profesionistas independientes y despachos — Fr. XI'),
        (r'^Correduría \(cliente\):$', '[PF: Profesional (cliente):][PM: Despacho (cliente):]'),
        (r'^\[Denominación de la Correduría\]$', '[Nombre del profesional o denominación del despacho]'),
        (r'-COR', '-PROF'),
        (r'a la Correduría', 'al [PF: profesional][PM: despacho]'), (r'de la Correduría', 'del [PF: profesional][PM: despacho]'),
        (r'la Correduría', '[PF: el profesional][PM: el despacho]'),
        (r'^Asunción personal y designación operativa$', '[PF: Asunción personal][PM: Designación del Representante Encargado de Cumplimiento]'),
        (r'^Registro SPPLD del Corredor Público Titular$', 'Registro SPPLD'),
        (r'^Metodología EBR — calibración por inciso$', 'Metodología EBR — calibración por servicio'),
        (r'Determinar los incisos del Apartado B aplicables', 'Determinar los incisos de la fracción XI que se prestan y el momento de sujeción (preparación o representación)'),
        (r'^Revisión del acuerdo de habilitación.*$', 'Revisión de la cartera de servicios y de los contratos de prestación de servicios profesionales'),
        (r'^Copia del acuerdo de habilitación.*$', 'Relación de servicios por inciso y muestra de contratos'),
        (r'^Entrevista con el Corredor Público Titular sobre los actos formalizados en los últimos 12 meses, por inciso.*$',
         'Entrevista con [PF: el profesional][PM: el Representante Encargado de Cumplimiento] sobre los servicios de los últimos 12 meses, por inciso, distinguiendo preparación y operación financiera en representación'),
        (r'volumen aproximado de actos por', 'volumen aproximado de servicios por'),
        (r'matriz de 22 rubros', 'matriz de 21 rubros'), (r'matriz de 15 rubros', 'matriz de 21 rubros'),
        (r'el Corredor Público Titular formalizó su responsabilidad personal', '[PF: el profesional formalizó su responsabilidad personal][PM: el despacho designó a su Representante Encargado de Cumplimiento]'),
        (r'^Revisión y firma de la Constancia de Asunción Personal$', '[PF: Revisión y firma de la Constancia de Asunción Personal][PM: Revisión del Acta de Designación y de la aceptación en el Portal]'),
        (r'firmado por el Corredor Público Titular', 'firmado por [PF: el profesional][PM: el órgano de administración]'),
        (r'el Corredor Público Titular', '[PF: el profesional][PM: el Representante Encargado de Cumplimiento]'),
        (r'Corredor Público Titular', '[PF: Profesional Titular][PM: Representante Encargado de Cumplimiento]'),
        (r'los incisos reales que maneja', 'los servicios reales que presta'),
        (r'de pólizas o actas formalizadas recientemente', 'de servicios prestados recientemente'),
        (r'Verificación de que la Sección A\.4 \(solicitante material\), la Sección E', 'Verificación de que la Sección E'),
        (r'secciones A\.4, E y F', 'secciones E y F'),
        (r'opera de forma consistente con las reglas del Apartado B: .*$', 'opera de forma consistente con la fracción XI: un Aviso por cada operación financiera en nombre y representación del cliente, sin umbral.'),
        (r'contra los asientos reales del libro de registro', 'contra los registros de operaciones en representación'),
        (r'cotejado contra el libro de registro', 'cotejado contra los registros de operaciones en representación'),
        (r'^Verificación específica de que los incisos b\), c\) y d\).*$', 'Verificación de que cada operación financiera realizada en nombre y representación del cliente fue avisada y de que la sola preparación no se avisó'),
        (r'^Folios de aviso SAT correspondientes.*$', 'Folios de aviso SAT de las operaciones en representación del periodo'),
        (r'^Acuerdo de habilitación y publicación en el DOF:$', 'Cédula profesional del titular o acta constitutiva del despacho:'),
        (r'^Identificación oficial del \[PF: Profesional Titular\]\[PM: Representante Encargado de Cumplimiento\]:$', 'Identificación oficial del [PF: profesional][PM: Representante Encargado de Cumplimiento]:'),
        (r'^Muestra de pólizas y actas formalizadas \(periodo de referencia\):$', 'Muestra de contratos y entregables de servicios (periodo de referencia):'),
        (r'con el \[PF: Profesional Titular\]\[PM: Representante Encargado de Cumplimiento\] — entrega', 'con [PF: el profesional][PM: el despacho] — entrega')]

def run():
    for f in sorted(glob.glob(os.path.join(REPO, 'inmobiliaria', '*.docx'))):
        b = os.path.basename(f)
        if not b.startswith(('PT0', 'ACTA0', 'REQ0', 'NOTA_', 'Diagnostico', 'Certus_Crono')):
            continue
        d = Doc(f, 'INM'); n = aplicar(d, COMUN + SECT['INM']); d.guardar(f); print(b, n)
    for kit, f in (('NOT', 'notarial/PT-NOT_Papeles_de_Trabajo_v2.docx'), ('COR', 'corredores/PT-COR_Papeles_de_Trabajo_v1.docx')):
        d = Doc(os.path.join(REPO, f), kit); n = aplicar(d, COMUN + SECT[kit]); d.guardar(os.path.join(REPO, f)); print(f, n)
    d = Doc(os.path.join(REPO, 'corredores/PT-COR_Papeles_de_Trabajo_v1.docx'), 'PROF')
    print('PROF', aplicar(d, PROF))
    d.guardar(os.path.join(REPO, 'profesionistas/PT-PROF_Papeles_de_Trabajo_v1.docx'))

if __name__ == '__main__':
    run()
