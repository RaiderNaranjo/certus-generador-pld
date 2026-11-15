"""D0A — Acta de Regularización Histórica: correcciones en INM, NOT y COR, y versión nueva para PROF (desde la NOT)."""
import os, sys, re, shutil
sys.path.insert(0, os.path.dirname(__file__))
from restructura import Doc
from build_d5 import Par

REPO = '/home/claude/certus-generador-pld/plantillas'
ARCH = {'INM': ('inmobiliaria/D0A_Acta_Regularizacion_Historica_v2.docx', 'inmobiliaria/D0A_Acta_Regularizacion_Historica_v3.docx'),
        'NOT': ('notarial/D0A-NOT_Acta_Regularizacion_Historica_v1.docx', 'notarial/D0A-NOT_Acta_Regularizacion_Historica_v2.docx'),
        'COR': ('corredores/D0A-COR_Acta_Regularizacion_Historica_v1.docx', 'corredores/D0A-COR_Acta_Regularizacion_Historica_v2.docx')}

COMUNES = [
    (r'Incluye, al menos, las Secciones A \(.*$',
     'Incluye las Secciones A a J del expediente: identificación, cotejo, origen y destino de los recursos, Grado de Riesgo, '
     'Beneficiario Controlador, PEP\'s, Perfil Transaccional, consulta de listas, acto u operación y documentos.'),
    (r'lo que constituye una defensa legal en caso de', 'como evidencia de diligencia en caso de'),
    (r'Este beneficio aplica exclusivamente porque: .*$',
     'El beneficio procede solo si: (i) el cumplimiento es espontáneo, sin requerimiento de la autoridad; (ii) se realiza antes '
     'del inicio de las facultades de verificación de la Secretaría; y (iii) se reconoce expresamente la falta dentro del plazo '
     'inicial del procedimiento, conforme al art. 55 de la Ley y al procedimiento que regula su Reglamento. Esta acta documenta '
     'la diligencia realizada, pero no sustituye ese reconocimiento.'),
    (r'Quedan protegidos por secreto profesional\.', 'Son confidenciales.'),
    (r'LFPIORPI \(reforma DOF 16 julio 2025\) y Reglamento \(reforma DOF 27 marzo 2026\)$',
     'LFPIORPI (reforma DOF 16-07-2025), Reglamento (reforma DOF 27-03-2026) y RCG (Acuerdo 115/2026, DOF 07-08-2026)'),
    (r'Aplicación retrospectiva de la matriz de factores de riesgo', 'Aplicación retrospectiva de la Metodología EBR'),
]

PROF = [
    (r'^Notaría$', '[PF: Profesional][PM: Despacho]'),
    (r'^Notario Titular$', '[PF: Profesional Titular][PM: Representante Encargado de Cumplimiento]'),
    (r'^\[Nombre del Notario Titular\]$', '[Nombre]'),
    (r'Las obligaciones de la LFPIORPI aplican a la Notaría desde antes de la reforma publicada el 16 de julio de 2025, que amplió y precisó sus obligaciones \(art\. 17, Fracción XII, Apartado A\)\.',
     'Las obligaciones de la LFPIORPI aplican a [PF: el profesional][PM: el despacho] por la prestación independiente de servicios profesionales de la fracción XI del art. 17, que la reforma publicada el 16 de julio de 2025 no modificó.'),
    (r'sin perjuicio de que el Notario Titular decida', 'sin perjuicio de que [PF: el profesional][PM: el despacho] decida'),
    (r'la Notaría inicia la auditoría', '[PF: el profesional][PM: el despacho] inicia la auditoría'),
    (r'para todos los actos formalizados en el período, conforme a lo establecido en D7-NOT\.', 'para todos los servicios de los incisos a) a e) prestados en el período, conforme a lo establecido en D7-PROF.'),
    (r'D6-NOT', 'D6-PROF'), (r'D7-NOT', 'D7-PROF'), (r'D12-NOT', 'D12-PROF'), (r'PT-NOT', 'PT-PROF'),
    (r'en los actos históricos', 'en los servicios históricos'),
    (r'Presentación extemporánea al portal SPPLD de avisos de los actos que.*$',
     'Presentación extemporánea al portal SPPLD de los Avisos de las operaciones financieras que [PF: el profesional][PM: el despacho] realizó en nombre y representación del cliente en relación con los incisos a) a e) (fracción XI, sin umbral; D1-PROF, sección 2.2). Obtención y archivo de acuses con folios de confirmación.'),
    (r'no hubo actos objeto de Aviso', 'no hubo operaciones objeto de Aviso'),
    (r'^Total de actos formalizados en el período$', 'Total de servicios de los incisos a) a e) prestados en el período'),
    (r'^Total de actos con expediente KYC reconstruido$', 'Total de clientes con expediente KYC reconstruido'),
    (r'^Total de actos objeto de Aviso \(Apartado A, incisos a\) a e\)\)$', 'Total de operaciones en nombre y representación objeto de Aviso'),
    (r'^Actos sin documentación', 'Servicios sin documentación'),
    (r'_____ actos', '_____ servicios'),
    (r'^Actos o comparecientes para los que', 'Servicios o clientes para los que'),
    (r'^NOTARIO TITULAR$', '[PF: PROFESIONAL TITULAR][PM: REPRESENTANTE ENCARGADO DE CUMPLIMIENTO]'),
    (r'^ENCARGADO OPERATIVO \(SI LO HAY\)$', '[PF: COLABORADOR DESIGNADO (SI LO HAY)][PM: ÓRGANO DE ADMINISTRACIÓN]'),
    (r'^\[Nombre del Encargado Operativo\]$', '[Nombre]'),
    (r'^Encargado Operativo$', '[PF: Colaborador designado][PM: Órgano de administración]'),
    (r'^Notario Titular$', '[PF: Profesional Titular][PM: Representante Encargado de Cumplimiento]'),
    (r'todos los actos formalizados en el período', 'todos los servicios prestados en el período'),
    (r'^No\. de instrumento \(escritura o acta\)$', 'No. de expediente del servicio (D7-PROF)'),
    (r'^Nombre del cliente y de los comparecientes$', 'Nombre del cliente y, en su caso, de las demás partes'),
    (r'^Fecha del acto$', 'Fecha del servicio y, en su caso, de la operación financiera'),
    (r'^Inciso del Apartado A \(a–e\) y tipo de acto$', 'Inciso de la fracción XI (a–e) y tipo de servicio'),
    (r'^Valor del acto en MXN \(cuando sea cuantificable\)$', '¿Hubo operación financiera en nombre y representación del cliente? (Sí / No) y monto en MXN'),
    (r'Fase 0 de regularización, se acoge', 'Fase 0 de regularización, busca acogerse'),
]

def aplicar(R, reglas):
    def f(par):
        t0 = par.text
        for a, b in reglas:
            par.replace(a, b, protect=False)
        return par if par.text != t0 else None
    return R.parrafos(f)

def run():
    for kit, (src, dst) in ARCH.items():
        R = Doc(os.path.join(REPO, src), kit)
        n = aplicar(R, COMUNES + [(r'Fase 0 de regularización, se acoge', 'Fase 0 de regularización, busca acogerse')])
        R.guardar(os.path.join(REPO, dst)); os.remove(os.path.join(REPO, src)); print(kit, n, 'párrafos')
    # PROF desde la NOT ya corregida
    R = Doc(os.path.join(REPO, ARCH['NOT'][1]), 'PROF')
    n = aplicar(R, PROF)
    R.guardar(os.path.join(REPO, 'profesionistas/D0A-PROF_Acta_Regularizacion_Historica_v1.docx')); print('PROF', n, 'párrafos')

if __name__ == '__main__':
    run()
