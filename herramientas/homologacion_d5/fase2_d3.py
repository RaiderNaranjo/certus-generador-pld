"""D3 — Expediente de Registro: sección V común y abreviaturas desarrolladas."""
import os, sys, glob
sys.path.insert(0, os.path.dirname(__file__))
from restructura import Doc
from build_d5 import p

REPO = '/home/claude/certus-generador-pld/plantillas'
ARCH = {'INM': ('inmobiliaria/D3_Expediente_Registro_SPPLD_v4.docx', 'inmobiliaria/D3_Expediente_Registro_SPPLD_v5.docx'),
        'NOT': ('notarial/D3-NOT_Expediente_Registro_SPPLD_v2.docx', 'notarial/D3-NOT_Expediente_Registro_SPPLD_v3.docx'),
        'COR': ('corredores/D3-COR_Expediente_Registro_SPPLD_v1.docx', 'corredores/D3-COR_Expediente_Registro_SPPLD_v2.docx'),
        'PROF': ('profesionistas/D3-PROF_Expediente_Registro_SPPLD_v1.docx', 'profesionistas/D3-PROF_Expediente_Registro_SPPLD_v2.docx')}

SECT = {
 'INM': dict(cambio='por ejemplo, de domicilio, de actividades o de fracciones', baja='',
    avisos='Presentar en el Portal, a más tardar el día 17 del mes siguiente, los Avisos de las fracciones V, V Bis y XV que alcancen el umbral, o el informe en ceros (art. 23 de la Ley; art. 25 RCG; D12).',
    extra=['[PM: Mantener vigente la designación del Representante Encargado de Cumplimiento y registrar en el Portal cualquier sustitución (art. 20 de la Ley; art. 10 RCG; D2).]']),
 'NOT': dict(cambio='incluidos cambio de adscripción, de domicilio o de incisos', baja=', por ejemplo por licencia o cesación de la patente',
    avisos='Presentar los Avisos a más tardar el día 17 del mes siguiente en el Portal (o, para el inciso a), por el medio fiscal del art. 24, último párrafo, de la Ley, si así se definió), o el informe en ceros (art. 25 RCG; D12-NOT).',
    extra=['Si el Colegio de Notarios de la entidad opera como Entidad Colegiada (arts. 26 y 27 de la Ley), evaluar si conviene presentar los Avisos por ese conducto. [VERIFICAR disponibilidad en la entidad correspondiente.]']),
 'COR': dict(cambio='incluidos cambio de plaza, de domicilio o de incisos', baja=', por ejemplo por suspensión o cancelación de la habilitación',
    avisos='Presentar en el Portal, a más tardar el día 17 del mes siguiente, los Avisos de todos los actos de los incisos b), c) y d) y de los avalúos con fe pública de 8,025 UMA o más, o el informe en ceros (art. 25 RCG; D12-COR).',
    extra=['Si el Colegio de Corredores Públicos de la plaza opera como Entidad Colegiada (arts. 26 y 27 de la Ley), evaluar si conviene presentar los Avisos por ese conducto. [VERIFICAR disponibilidad en la plaza correspondiente.]']),
 'PROF': dict(cambio='por ejemplo, de domicilio o de los incisos de la fracción XI en que se presten servicios', baja='',
    avisos='Presentar en el Portal, a más tardar el día 17 del mes siguiente, un Aviso por cada operación financiera realizada en nombre y representación del cliente en relación con los incisos a) a e), o el informe en ceros (art. 25 RCG; D12-PROF).',
    extra=['[PM: Mantener vigente la designación del Representante Encargado de Cumplimiento y registrar en el Portal cualquier sustitución (art. 20 de la Ley; art. 10 RCG; D2-PROF-PM).]']),
}

ABREV = [('[PF: Titular][PM: Repr. Legal]', '[PF: Titular][PM: Representante legal]'),
         ('[PF: Titular][PM: Repr. Cumpl.]', '[PF: Titular][PM: Representante Encargado de Cumplimiento]'),
         ('[PF: Profesional Titular][PM: Repr. Cumpl.]', '[PF: Profesional Titular][PM: Representante Encargado de Cumplimiento]'),
         ('[PF: Profesional Titular][PM: Repr. Legal]', '[PF: Profesional Titular][PM: Representante legal]'),
         ('[PF: Profesional][PM: Repr. Cumpl.]', '[PF: Profesional Titular][PM: Representante Encargado de Cumplimiento]'),
         ('Enc. Operativo', 'Encargado Operativo'), ('Corredor Titular', 'Corredor Público Titular'),
         ('Corredor Público Público', 'Corredor Público')]

def cuerpo(kit):
    d = SECT[kit]; s = {'INM': '', 'NOT': '-NOT', 'COR': '-COR', 'PROF': '-PROF'}[kit]
    def f(ctx):
        for t in ['Actualizar los datos del registro dentro de los seis días hábiles siguientes a cualquier cambio, %s (art. 7 RCG), y dar de baja la actividad cuando se deje de realizar%s (art. 8 RCG).' % (d['cambio'], d['baja']),
                  'Consultar el Portal al menos una vez cada día hábil. Se cuenta con tres días hábiles para abrir los documentos pendientes y la notificación electrónica se tiene por realizada al cuarto día hábil siguiente a su envío (arts. 5 Bis y 6 RCG).',
                  'Atender los requerimientos de información del SAT sobre el alta y registro dentro de los diez días hábiles siguientes a su notificación, prorrogables hasta por cinco a solicitud del interesado (art. 9 RCG).',
                  d['avisos'],
                  'Presentar el Aviso de 24 horas por sospecha o por hechos o indicios (art. 18, fr. VI, de la Ley; D9%s); su envío por el Portal se habilita seis meses después de la Resolución de formatos (Transitorio Quinto del Acuerdo 115/2026).' % s,
                  'Contar con el Manual de Políticas Internas (D5%s) a los noventa días naturales del alta y registro (art. 37 RCG); la metodología EBR y el Manual que la incluya son exigibles a partir del 1 de marzo de 2027 y los mecanismos automatizados a más tardar el 1 de junio de 2027 (Transitorios Segundo, Tercero y Noveno).' % s,
                  'Conservar la constancia del registro y los acuses durante diez años (D10%s).' % s] + d['extra']:
            if t.startswith('[PM:') and t.endswith(']'):
                p(ctx, t[5:-1], pm=True)
            else:
                p(ctx, t, bullet=True)
    return f

def run():
    for kit, (src, dst) in ARCH.items():
        R = Doc(os.path.join(REPO, src), kit)
        n = R.texto(lambda t: [t := t.replace(a, b) for a, b in ABREV][-1])
        R.reemplazar_cuerpo(5, cuerpo(kit))
        R.guardar(os.path.join(REPO, dst)); os.remove(os.path.join(REPO, src)); print(kit, n)

if __name__ == '__main__':
    run()
