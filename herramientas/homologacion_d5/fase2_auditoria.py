"""Esquema de auditoría acordado con Luis (14-nov-2026):
- Riesgo bajo o medio: dictamen del área interna del cliente (art. 44 RCG), acompañada por Certus (metodología, PT, formato
  D13, revisión de calidad y capacitación anual del auditor interno), sin que Certus firme.
- Riesgo alto, o si el cliente lo elige: auditor externo independiente con certificación UIF que no le haya prestado
  servicios con conflicto de interés en el periodo (art. 45 RCG).
- C3: mantenimiento con revisión anual preparatoria; no es el dictamen.
- La revisión de cierre de la implementación es un servicio contractual: Certus actúa como revisor, no como auditor externo.
"""
import os, sys, glob
sys.path.insert(0, os.path.dirname(__file__))
from restructura import Doc
from fase2_d0a import aplicar

REPO = '/home/claude/certus-generador-pld/plantillas'
RESP = {'INM': '[PF: titular][PM: Representante Encargado de Cumplimiento]',
        'NOT': 'Notario Titular', 'COR': 'Corredor Público Titular',
        'PROF': '[PF: profesional][PM: Representante Encargado de Cumplimiento]'}

def carta(kit):
    r = RESP[kit]
    return [
        (r'AUDITORÍAS ANUALES C3 \(Después — Solo revisión\)', 'REVISIÓN ANUAL C3 (Después — mantenimiento)'),
        (r'^Condensada — 5-7 días de auditoría$', 'Condensada — 5-7 días de revisión'),
        (r'Incluido en suscripción C3 \(auditoría anual de mantenimiento\)', 'Incluido en suscripción C3 (revisión anual preparatoria)'),
        (r'^Opinión de mantenimiento \+ continuidad del programa$', 'Informe de revisión preparatoria; el dictamen anual se emite según el Riesgo (sección V)'),
        (r'^Las auditorías anuales \(modalidad C3\) son mucho más ágiles porque:$', 'Las revisiones anuales (modalidad C3) y el dictamen anual son mucho más ágiles porque:'),
        (r'Las auditorías anuales \(modalidad C3\) son mucho más ágiles porque:', 'Las revisiones anuales (modalidad C3) y el dictamen anual son mucho más ágiles porque:'),
        (r'^4\. El auditor ya conoce su operación$', '4. Certus ya conoce su operación'),
        (r'^Después de la auditoría de cierre, el auditor entiende', 'Después de la revisión de cierre, Certus entiende'),
        (r'auditoría anual C3 toma 5-7 días', 'revisión anual C3 toma 5-7 días'),
        (r'^Contrato anual de mantenimiento que incluye: .*$',
         'Contrato anual de mantenimiento que incluye: revisión anual preparatoria (5-7 días), actualización del programa ante '
         'cambios normativos (DOF), revisión trimestral del control de avisos, capacitación semestral y 10 consultas de soporte '
         'al %s. La revisión preparatoria deja todo listo para el dictamen anual, pero no lo sustituye: el dictamen se emite '
         'conforme a la opción B o C, según su Riesgo. Costo: diseñado para ser accesible después de haber cumplido con la '
         'implementación inicial.' % r),
        (r'^Opción B: Auditoría sin suscripción$', 'Opción B: Dictamen anual con Riesgo bajo o medio — auditoría interna acompañada por Certus'),
        (r'^Si prefiere no renovar, puede solicitar auditorías puntuales.*$',
         'Si su Riesgo es bajo o medio, el dictamen anual —obligatorio a partir del periodo 2028 (art. 18, fr. XI, de la Ley; art. 42 RCG)— '
         'puede emitirlo su propia área de auditoría o control interno, integrada por personal '
         'independiente del %s y con al menos una capacitación anual acreditada (art. 44 RCG). Certus la acompaña con la '
         'metodología, los papeles de trabajo, el formato del dictamen (D13), una revisión de calidad antes de la firma y la '
         'capacitación anual que exige el art. 44; el dictamen lo firma su área interna, no Certus.' % r),
        (r'^Opción C: Auditoría interna \(solo si su Riesgo es bajo o medio\)$', 'Opción C: Dictamen anual con Riesgo alto, o si así lo elige — auditor externo independiente'),
        (r'^La auditoría anual es obligatoria a partir del periodo 2028.*$',
         'Si su Riesgo es alto, o si prefiere un auditor externo, el dictamen lo emite una persona auditora externa independiente '
         'con certificación vigente de la UIF que no le haya prestado servicios que generen conflicto de interés durante el '
         'periodo auditado (art. 45 RCG). Certus le refiere a un auditor aliado que cumpla esos requisitos. Certus solo puede '
         'actuar como auditor externo de clientes a los que no les haya prestado servicios en el año auditado. Mantener el '
         'programa sin dictamen anual no es una opción legal.'),
    ]

NOTA = [
    (r'el auditor externo \(CERTUS', 'el revisor (CERTUS'),
    (r'Proteger al auditor externo\.', 'Proteger al revisor.'),
    (r'La auditoría anual \(incluida en el C3\) verifica', 'La revisión anual preparatoria (incluida en el C3) verifica'),
    (r'^(Por ello, esta revisión de cierre es un servicio contractual.*?)$',
     r'\1 Esquema del Programa: con Riesgo bajo o medio, el dictamen lo firma el área interna independiente del cliente (art. '
     r'44 RCG) y Certus la acompaña con metodología, papeles de trabajo, revisión de calidad y capacitación anual, sin firmar; '
     r'con Riesgo alto, o a elección del cliente, lo emite un auditor externo certificado por la UIF sin conflicto de interés '
     r'en el periodo (art. 45 RCG).'),
]
ACTAS = [
    (r'^Auditor externo$', 'Revisor de cierre'),
    (r'^Tipo de auditoría$', 'Tipo de revisión'),
    (r'^Programada — Cierre del Programa Certus PLD$', 'Revisión de cierre de la implementación (servicio contractual; no es el dictamen del art. 42 RCG)'),
    (r'^AUDITOR$', 'REVISOR'),
    (r'El auditor externo hizo entrega', 'El revisor hizo entrega'),
    (r'seleccionada por el auditor', 'seleccionada por el revisor'),
]

def run():
    for kit, f in (('INM', 'inmobiliaria/CARTA_Bienvenida_Orientacion_Cliente_v2.docx'), ('NOT', 'notarial/CARTA-NOT_Bienvenida_Orientacion_Cliente_v2.docx'),
                   ('COR', 'corredores/CARTA-COR_Bienvenida_Orientacion_Cliente_v2.docx'), ('PROF', 'profesionistas/CARTA-PROF_Bienvenida_Orientacion_Cliente_v1.docx')):
        d = Doc(os.path.join(REPO, f), kit); n = aplicar(d, carta(kit)); d.restilar_variantes(); d.guardar(os.path.join(REPO, f)); print(f, n)
    for f, reg in (('inmobiliaria/NOTA_Metodologica_Auditoria_Cierre_v2.docx', NOTA), ('inmobiliaria/ACTA01_Inicio_Auditoria.docx', ACTAS),
                   ('inmobiliaria/ACTA02_Cierre_Auditoria.docx', ACTAS)):
        d = Doc(os.path.join(REPO, f), 'INM'); n = aplicar(d, reg); d.restilar_variantes(); d.guardar(os.path.join(REPO, f)); print(f, n)

if __name__ == '__main__':
    run()
