"""D1 — Dictamen de Sujeción Normativa: seis secciones en los cuatro kits (se conservan 2.2 de PROF y 2.3 de COR)."""
import os, sys
sys.path.insert(0, os.path.dirname(__file__))
from restructura import Doc
from build_d5 import p
from fase2_d0a import aplicar

REPO = '/home/claude/certus-generador-pld/plantillas'
ARCH = {'INM': ('inmobiliaria/D1_Dictamen_Sujecion_Normativa_v4.docx', 'inmobiliaria/D1_Dictamen_Sujecion_Normativa_v5.docx'),
        'NOT': ('notarial/D1-NOT_Dictamen_Sujecion_Normativa_v1.docx', 'notarial/D1-NOT_Dictamen_Sujecion_Normativa_v2.docx'),
        'COR': ('corredores/D1-COR_Dictamen_Sujecion_Normativa_v1.docx', 'corredores/D1-COR_Dictamen_Sujecion_Normativa_v2.docx'),
        'PROF': ('profesionistas/D1-PROF_Dictamen_Sujecion_Normativa_v1.docx', 'profesionistas/D1-PROF_Dictamen_Sujecion_Normativa_v2.docx')}

def lista(ctx, intro, items):
    p(ctx, intro)
    for t in items:
        p(ctx, t, bullet=True)

ALCANCE = {
 'INM': ('No toda operación con inmuebles actualiza la sujeción. No se activa, o no genera obligaciones, en estos casos:', [
    'Fracción V: requiere que la actividad se realice de manera habitual o profesional; una venta aislada de quien no se dedica a la construcción, el desarrollo o la intermediación no la actualiza (la fracción V Bis, en cambio, no exige habitualidad).',
    'Fracción XV: la constitución de derechos de uso o goce con renta mensual de 1,605 UMA o menos ($188,283 en 2026) no es Actividad Vulnerable.',
    'Los actos u operaciones por montos inferiores a los señalados no dan lugar a obligación alguna, salvo que la suma acumulada con el mismo cliente en seis meses supere el umbral de Aviso (art. 17, penúltimo párrafo, de la Ley).']),
 'NOT': ('No todo acto ante el Notario actualiza la sujeción. Quedan fuera, por disposición expresa del art. 17, fr. XII, Apartado A, de la Ley:', [
    'Inciso a): las garantías que se constituyan a favor de instituciones del sistema financiero u organismos públicos de vivienda.',
    'Inciso d): los fideicomisos que se constituyan para garantizar algún crédito a favor de instituciones del sistema financiero u organismos públicos de vivienda.',
    'Inciso e): los mutuos o créditos en los que el acreedor forme parte del sistema financiero o sea un organismo público de vivienda.',
    'Los actos distintos de los incisos a) a e) (por ejemplo, testamentos o poderes revocables), que no son Actividad Vulnerable.',
    'Cuando el Notario presta servicios de fe pública al Poder Judicial o a autoridades administrativas en funciones de impartición de justicia, el expediente se integra solo con los datos del Anexo 7 Bis RCG (art. 13 Bis RCG).']),
 'COR': ('No todo acto ante el Corredor actualiza la sujeción. Quedan fuera, por disposición expresa del art. 17, fr. XII, Apartado B, de la Ley:', [
    'Inciso a): los avalúos sobre bienes con valor inferior a 8,025 UMA ($941,413 en 2026).',
    'Inciso c): los fideicomisos que se constituyan para garantizar algún crédito a favor de instituciones del sistema financiero.',
    'Inciso d): los mutuos o créditos mercantiles en los que el acreedor forme parte del sistema financiero.',
    'Los actos distintos de los incisos a) a d), que no son Actividad Vulnerable.',
    'Cuando el Corredor presta servicios de fe pública al Poder Judicial o a autoridades administrativas en funciones de impartición de justicia, el expediente se integra solo con los datos del Anexo 7 Bis RCG (art. 13 Bis RCG).']),
}

def run():
    # INM
    R = Doc(os.path.join(REPO, ARCH['INM'][0]), 'INM')
    R.renombrar(2, 'Sujeción — Fracciones V, V Bis y XV')
    R.degradar(3, '2.3 Fracción V Bis: recepción de recursos para un desarrollo inmobiliario')
    R.degradar(3, '2.4 Fracción XV: arrendamiento')
    aplicar(R, [(r'^Identificación del arrendatario$', 'Identificación del arrendatario (renta mensual superior a)')])
    R.insertar_despues(2, 'Alcance: cuándo no se activa la sujeción', lambda ctx: lista(ctx, *ALCANCE['INM']))
    R.insertar_despues(3, 'Identificación del Cliente y del Beneficiario Controlador', lambda ctx: (
        p(ctx, 'La empresa identifica y conoce de manera directa a cada cliente, verifica su identidad con documentos oficiales y '
               'recaba copia de ellos (art. 18, fr. I, de la Ley; art. 12 RCG); en las relaciones de negocios, recaba además la '
               'información sobre su actividad u ocupación (art. 18, fr. II). Si el cliente es persona moral o fideicomiso, '
               'identifica a su Beneficiario Controlador conforme al orden de prelación del art. 23 Quinquies RCG; si es persona '
               'física, recaba su declaración sobre si conoce la existencia de uno (art. 18, fr. III, de la Ley). El detalle está '
               'en el Expediente KYC (D7) y en el Manual del Beneficiario Controlador (D14).'),))
    R.insertar_despues(4, 'Régimen Sancionatorio', lambda ctx: (
        p(ctx, 'Incumplir las obligaciones del art. 18 o presentar Avisos con hasta 30 días de retraso se sanciona con multa de '
               '200 a 2,000 UMA; omitir Avisos o participar en operaciones con efectivo prohibidas por el art. 32 (inmuebles de '
               '8,025 UMA o más y rentas mensuales de 3,210 UMA o más), con multa de 10,000 a 65,000 UMA o del 10% al 100% del '
               'valor del acto, la que resulte mayor (arts. 53 y 54 de la Ley). La autoridad se abstiene de sancionar por única '
               'ocasión si el cumplimiento es espontáneo y se reconoce la falta en los términos del art. 55.'),))
    R.guardar(os.path.join(REPO, ARCH['INM'][1])); os.remove(os.path.join(REPO, ARCH['INM'][0]))
    # NOT
    R = Doc(os.path.join(REPO, ARCH['NOT'][0]), 'NOT')
    aplicar(R, [(r'no lo sanciona directamente: notifica a la autoridad', 'no lo sanciona directamente: informa a la autoridad')])
    R.renombrar(3, 'Identificación del Cliente, del Beneficiario Controlador y del Solicitante')
    R.agregar_a(3, lambda ctx: p(ctx, 'Además, al momento del acto el Notario identifica a la persona que solicita materialmente la '
        'protocolización o formalización del instrumento, con los datos del Anexo 3, inciso a), numerales i), ii), ix) y x), RCG '
        '(art. 13 RCG; D7-NOT, Sección A.4).'))
    R.insertar_despues(2, 'Alcance: cuándo no se activa la sujeción', lambda ctx: lista(ctx, *ALCANCE['NOT']))
    R.guardar(os.path.join(REPO, ARCH['NOT'][1])); os.remove(os.path.join(REPO, ARCH['NOT'][0]))
    # COR
    R = Doc(os.path.join(REPO, ARCH['COR'][0]), 'COR')
    aplicar(R, [(r'notifica a la autoridad', 'informa a la autoridad')])
    R.insertar_despues(2, 'Alcance: cuándo no se activa la sujeción', lambda ctx: lista(ctx, *ALCANCE['COR']))
    R.guardar(os.path.join(REPO, ARCH['COR'][1])); os.remove(os.path.join(REPO, ARCH['COR'][0]))
    # PROF: estructura de referencia; solo el título de la sección V
    R = Doc(os.path.join(REPO, ARCH['PROF'][0]), 'PROF')
    R.renombrar(5, 'Régimen Sancionatorio según el Tipo de Sujeto Obligado')
    R.guardar(os.path.join(REPO, ARCH['PROF'][1])); os.remove(os.path.join(REPO, ARCH['PROF'][0]))
    for k, (_, dst) in ARCH.items():
        R = Doc(os.path.join(REPO, dst), k)
        print(k, [R.titulo(i) for i in range(1, len(R.caps()) + 1)])

if __name__ == '__main__':
    run()
