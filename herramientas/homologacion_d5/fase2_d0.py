"""D0 — Carta de Recomendaciones Urgentes: siete secciones iguales en los cuatro kits."""
import os, sys, re
sys.path.insert(0, os.path.dirname(__file__))
from restructura import Doc
from build_d5 import p, tabla, aplicar_globales

REPO = '/home/claude/certus-generador-pld/plantillas'
ARCH = {'INM': ('inmobiliaria/D0_Carta_Recomendaciones_Urgentes_v4.docx', 'inmobiliaria/D0_Carta_Recomendaciones_Urgentes_v5.docx'),
        'NOT': ('notarial/D0-NOT_Carta_Recomendaciones_Urgentes_v2.docx', 'notarial/D0-NOT_Carta_Recomendaciones_Urgentes_v3.docx'),
        'COR': ('corredores/D0-COR_Carta_Recomendaciones_Urgentes_v1.docx', 'corredores/D0-COR_Carta_Recomendaciones_Urgentes_v2.docx'),
        'PROF': ('profesionistas/D0-PROF_Carta_Recomendaciones_Urgentes_v1.docx', 'profesionistas/D0-PROF_Carta_Recomendaciones_Urgentes_v2.docx')}

QUIEN = {'INM': 'su empresa', 'NOT': 'la Notaría', 'COR': 'la Correduría', 'PROF': '[PF: usted][PM: el despacho]'}
DIRIGIDO = {'INM': 'empresas inmobiliarias', 'NOT': 'notarías públicas', 'COR': 'corredurías públicas',
            'PROF': 'profesionistas independientes y despachos'}
RECIBE = {'INM': ('Empresa:', '[EMPRESA]'), 'NOT': ('Notaría:', '[Denominación de la Notaría]'),
          'COR': ('Correduría:', '[Denominación de la Correduría]'), 'PROF': ('[PF: Profesional:][PM: Despacho:]', '')}

def comunes(t):
    t = t.replace('Omisión de registro ante el SAT (SPPLD)', 'Omisión de alta y registro en el padrón de Actividades Vulnerables')
    t = t.replace('(portal SPPLD)', '(Portal en Internet del SAT)')
    t = t.replace('diagnostico preliminar', 'diagnóstico preliminar')
    t = t.replace('notifica a la autoridad estatal', 'informa a la autoridad estatal')
    t = t.replace('a) — Reglamento, art. 29', 'a); art. 17, fr. XII-B, de la Ley')
    t = t.replace('Sin auditoría interna del programa PLD', 'Sin auditoría anual del programa (primer periodo: 2028)')
    t = t.replace('Avisos por sospecha (24 horas) y Avisos mensuales no presentados', 'Avisos mensuales no presentados y falta de procedimiento para los Avisos de 24 horas')
    return t

def run():
    for kit, (src, dst) in ARCH.items():
        R = Doc(os.path.join(REPO, src), kit)
        R.texto(comunes)
        tit = [R.titulo(i) for i in range(1, len(R.caps()) + 1)]
        if kit == 'INM':
            R.renombrar(4, 'Actuación espontánea — ventana de oportunidad')
            R.insertar_despues(3, 'Un riesgo particular del sector: el uso de efectivo', lambda ctx: (
                p(ctx, 'En el sector inmobiliario la exposición más alta no está en las multas por obligaciones formales, sino en '
                       'el efectivo: está prohibido liquidar o aceptar en efectivo la transmisión o constitución de derechos reales '
                       'sobre inmuebles de 8,025 UMA o más ($941,413 en 2026) y la renta mensual de 3,210 UMA o más ($376,565) '
                       '(art. 32, fr. I y VII, de la Ley). Participar en esos actos se sanciona con multa de 10,000 a 65,000 UMA, o '
                       'del 10% al 100% del valor del acto, la que resulte mayor (arts. 53, fr. VII, y 54, fr. III). Sobre un inmueble '
                       'de varios millones de pesos, la multa puede superar el valor de la propia comisión o utilidad de la operación.'),))
        else:
            R.renombrar(4, 'Un riesgo particular del sector')
            R.agregar_a(5, lambda ctx: p(ctx, 'Las multas de los arts. 53 y 54 de la Ley van de 200 a 2,000 UMA por incumplir las '
                                              'obligaciones del art. 18 y de 10,000 a 65,000 UMA (o del 10% al 100% del valor del '
                                              'acto) por omitir Avisos u operar con efectivo prohibido.'))
            R.insertar_despues(5, 'Recomendación', lambda ctx: p(ctx,
                'Se recomienda iniciar de inmediato la implementación del Programa Certus PLD, desarrollado por Certus Consultores '
                'y Espinosa-Salazar Abogados Corporativos para %s, que cubre la totalidad de las obligaciones de la LFPIORPI, su '
                'Reglamento y las Reglas de Carácter General reformadas por el Acuerdo 115/2026.' % DIRIGIDO[kit]))
            R.insertar_despues(6, 'Acuse de recibo', lambda ctx: tabla(ctx, [
                ('**Nombre:**', ''), ('**Cargo:**', ''), ('**%s**' % RECIBE[kit][0] if not RECIBE[kit][0].startswith('[') else RECIBE[kit][0], RECIBE[kit][1]),
                ('**Fecha de recepción:**', '[dd/mm/aaaa]'), ('**Firma:**', '')], [3600, 5900], encabezado=False))
        R.guardar(os.path.join(REPO, dst))
        os.remove(os.path.join(REPO, src))
        print(kit, [R.titulo(i) for i in range(1, len(R.caps()) + 1)])

if __name__ == '__main__':
    run()
