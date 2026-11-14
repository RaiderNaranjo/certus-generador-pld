"""CARTA de Bienvenida: correcciones de fondo en INM, NOT y COR, y versión nueva para PROF (desde la NOT)."""
import os, sys
sys.path.insert(0, os.path.dirname(__file__))
from restructura import Doc
from fase2_d0a import aplicar

REPO = '/home/claude/certus-generador-pld/plantillas'
ARCH = {'INM': ('inmobiliaria/CARTA_Bienvenida_Orientacion_Cliente.docx', 'inmobiliaria/CARTA_Bienvenida_Orientacion_Cliente_v2.docx'),
        'NOT': ('notarial/CARTA-NOT_Bienvenida_Orientacion_Cliente_v1.docx', 'notarial/CARTA-NOT_Bienvenida_Orientacion_Cliente_v2.docx'),
        'COR': ('corredores/CARTA-COR_Bienvenida_Orientacion_Cliente_v1.docx', 'corredores/CARTA-COR_Bienvenida_Orientacion_Cliente_v2.docx')}

COMUNES = [
    (r'^Usted queda protegido legalmente\..*?(El Acta de Regularización Histórica \((D0A[^)]*)\)).*$',
     r'Usted queda en la mejor posición posible ante una revisión: si el cumplimiento es espontáneo, anterior al inicio de '
     r'las facultades de verificación de la Secretaría, y la falta se reconoce expresamente dentro del plazo inicial del '
     r'procedimiento, la autoridad se abstiene de sancionar por única ocasión o reduce las multas hasta en un 50% (art. 55 '
     r'de la Ley). El Acta de Regularización Histórica (\2) documenta la diligencia realizada.'),
    (r'Obtención del beneficio de actuación espontánea', 'Preparación para el beneficio de actuación espontánea'),
    (r'Opinión de cumplimiento \+ defensa legal contra revisiones pasadas', 'Opinión de cumplimiento y evidencia de la regularización histórica'),
    (r'^Si intenta mantener el programa internamente sin auditoría externa.*$',
     'La auditoría anual es obligatoria a partir del periodo 2028 (art. 18, fr. XI, de la Ley; art. 42 RCG). Si su Riesgo es '
     'bajo o medio, puede practicarla un área de auditoría o control interno independiente (art. 44 RCG); si es alto, o si así '
     'lo elige, una persona auditora externa independiente con certificación vigente de la UIF (art. 45 RCG). Mantener el '
     'programa sin ninguna de las dos no es una opción legal.'),
    (r'^Opción C: Interno \(NO RECOMENDADO\)$', 'Opción C: Auditoría interna (solo si su Riesgo es bajo o medio)'),
    (r'\[PF: titular\]\[PM: RC\]', '[PF: titular][PM: Representante Encargado de Cumplimiento]'),
]
# Independencia del auditor externo (art. 45, inciso d), RCG): se agrega al describir la suscripción C3.
C3 = (r'^(Contrato anual de mantenimiento que incluye: .*?)(Costo:.*)$',
      r'\1Si el dictamen del art. 42 RCG debe emitirlo una persona auditora externa, esta no puede haberle prestado servicios '
      r'que generen conflicto de interés durante el periodo auditado (art. 45, inciso d), RCG); en ese caso, la revisión anual '
      r'de la suscripción C3 funciona como revisión preparatoria y el dictamen lo emite un auditor externo independiente. \2')

PROF = [
    (r'Para notarías — Fracción XII, Apartado A', 'Para profesionistas independientes y despachos — Fracción XI'),
    (r'^Notaría$', '[PF: Profesional][PM: Despacho]'),
    (r'^Notario Titular$', '[PF: Profesional Titular][PM: Representante Encargado de Cumplimiento]'),
    (r'^\[Nombre del Notario Titular\]$', '[Nombre]'),
    (r'todos los actos formalizados en su notaría', 'todos los servicios de los incisos a) a e) que [PF: usted prestó][PM: su despacho prestó]'),
    (r'todos los actos formalizados en ese período histórico', 'todos los servicios prestados en ese período histórico'),
    (r'para operaciones que no fueron reportadas originalmente', 'por las operaciones financieras realizadas en nombre y representación del cliente que no fueron reportadas'),
    (r'D0A-NOT', 'D0A-PROF'),
    (r'su notaría opera', '[PF: su práctica][PM: su despacho] opera'),
    (r'el Notario dedica tiempo', '[PF: el profesional][PM: el despacho] dedica tiempo'),
    (r'el Notario solo documenta', '[PF: el profesional][PM: el despacho] solo documenta'),
    (r'Cuando se autoriza un instrumento,', 'Cuando se presta el servicio,'),
    (r'cómo funciona su notaría', 'cómo funciona [PF: su práctica][PM: su despacho]'),
    (r'^Para mantener esto, el Notario Titular debe:$', 'Para mantener esto, [PF: el profesional][PM: el Representante Encargado de Cumplimiento] debe:'),
    (r'ANTES de autorizar cada instrumento', 'ANTES de preparar o ejecutar cada servicio'),
    (r'Presentar los avisos al SAT dentro del plazo — incisos b\), c\) y e\) siempre, sin importar el monto',
     'Presentar los avisos al SAT por cada operación financiera realizada en nombre y representación del cliente'),
    (r'Capacitar al personal de atención al público', 'Capacitar a socios, colaboradores y personal de apoyo'),
    (r'soporte al Notario', 'soporte'),
]

def run():
    for kit, (src, dst) in ARCH.items():
        R = Doc(os.path.join(REPO, src), kit)
        n = aplicar(R, COMUNES + [C3])
        R.guardar(os.path.join(REPO, dst)); os.remove(os.path.join(REPO, src)); print(kit, n)
    R = Doc(os.path.join(REPO, ARCH['NOT'][1]), 'PROF')
    print('PROF', aplicar(R, PROF))
    R.guardar(os.path.join(REPO, 'profesionistas/CARTA-PROF_Bienvenida_Orientacion_Cliente_v1.docx'))

if __name__ == '__main__':
    run()
