"""D4 — Gap Analysis PLD/FT, homologado: 20 rubros comunes (misma numeración) más los rubros propios del sector.

Uso: python3 build_d4.py <dir_contenedor> <dir_salida>
"""
import os, sys
sys.path.insert(0, os.path.dirname(__file__))
import build_d5 as B
from build_d5 import Ctx, p, tabla, portada, pie, escribir_par, Par
from docxlib import instalar_estilos, q, espaciar_tras_tablas
from docx import Document
import canon

SECCIONES = [('I', 'Instrucciones de Uso', 'Instrucciones'), ('II', 'Matriz de Brecha', 'Matriz'),
             ('III', 'Resumen Ejecutivo', 'Resumen'), ('IV', 'Plan de Acción Correctiva', 'Plan de acción')]

def sec(ctx, rom, extra=''):
    t = [s[1] for s in SECCIONES if s[0] == rom][0] + extra
    ctx.body.append(escribir_par(Par.from_text('Sección %s: %s' % (rom, t)), 'CertusCapitulo'))

M18, M65 = '$234,620', '$7,625,150 o 10%–100% del valor'
M33 = '$1,173,100'

S = {
 'INM': dict(suf='', R='[PF: Titular][PM: Representante Encargado de Cumplimiento]', RL='[PF: Titular][PM: Representante legal]',
    org='[PF: Titular][PM: Órgano de administración]',
    d2=('[PF: Constancia de Asunción Personal del titular][PM: Representante Encargado de Cumplimiento designado (persona física)]', 'art. 20 de la Ley; art. 10 RCG',
        '[PF: Firmar la Constancia de Asunción Personal (D2-PF)][PM: Celebrar la sesión y firmar el Acta de Designación (D2)]', 'D2'),
    avisos=('Avisos mensuales de las fracciones V, V Bis y XV e informes en ceros', 'arts. 17 y 23 de la Ley; arts. 24 a 25 RCG'),
    acum='Implementar el seguimiento y la acumulación de actos u operaciones en periodos de hasta seis meses',
    efectivo='Difundir al personal de ventas y cobranza que no se acepta efectivo cuando el valor del inmueble es de 8,025 UMA o más, ni en rentas de 3,210 UMA mensuales o más (art. 32, fr. I y VII)',
    extra=[('Aviso consolidado por desarrollo y copropietarios en arrendamiento', 'arts. 24 Bis 1 y 24 Bis 6 RCG', M65,
            'Definir el criterio de Aviso por desarrollo (V Bis) y por copropietario (XV)', 'D12', 'Semana 3-4')]),
 'NOT': dict(suf='-NOT', R='Notario Titular', RL='Notario Titular', org='Notario Titular',
    d2=('Constancia de Asunción Personal del Notario Titular', 'art. 20, cuarto párrafo, de la Ley',
        'Firmar la Constancia de Asunción Personal y, en su caso, designar Encargado Operativo', 'D2-NOT'),
    avisos=('Avisos mensuales (incisos b), c) y e) siempre; a) desde 8,000 UMA; d) desde 4,000 UMA) e informes en ceros', 'arts. 17, 23 y 24 de la Ley; arts. 24 a 25 RCG'),
    acum='Implementar la acumulación de los actos de los incisos a) y d) en periodos de hasta seis meses (art. 19 RCG)',
    efectivo='Verificar que no se liquiden en efectivo los inmuebles de 8,025 UMA o más ni las acciones o partes sociales de 3,210 UMA o más (art. 32, fr. I y VI)',
    extra=[('Identificación de la persona que solicita materialmente la protocolización', 'art. 13 RCG', M18,
            'Incorporar al expediente la identificación del solicitante material', 'D7-NOT, Sección A.4', 'Semana 2'),
           ('Identificación de la forma de pago en el instrumento', 'art. 33 de la Ley', M33,
            'Asentar en cada instrumento la forma de pago cuando el valor sea de 8,025 UMA o más', 'D7-NOT, Sección I', 'Semana 2')]),
 'COR': dict(suf='-COR', R='Corredor Público Titular', RL='Corredor Público Titular', org='Corredor Público Titular',
    d2=('Constancia de Asunción Personal del Corredor Público Titular', 'art. 20, cuarto párrafo, de la Ley',
        'Firmar la Constancia de Asunción Personal y, en su caso, designar Encargado Operativo', 'D2-COR'),
    avisos=('Avisos mensuales (incisos b), c) y d) siempre; a) avalúos desde 8,025 UMA) e informes en ceros', 'arts. 17 y 23 de la Ley; arts. 24 a 25 RCG'),
    acum='Documentar que los incisos del Apartado B no se acumulan (art. 19 RCG) y vigilar el fraccionamiento artificial (Aviso de 24 horas)',
    efectivo='Verificar que no se liquiden en efectivo las acciones o partes sociales de 3,210 UMA o más (art. 32, fr. VI)',
    extra=[('Identificación de la persona que solicita materialmente la formalización', 'art. 13 RCG', M18,
            'Incorporar al expediente la identificación del solicitante material', 'D7-COR, Sección A.4', 'Semana 2'),
           ('Identificación de la forma de pago en la póliza o acta', 'art. 33 de la Ley', M33,
            'Asentar la forma de pago cuando el valor sea de 8,025 UMA o más', 'D7-COR, Sección I', 'Semana 2')]),
 'PROF': dict(suf='-PROF', R='[PF: Profesional Titular][PM: Representante Encargado de Cumplimiento]', RL='[PF: Profesional Titular][PM: Representante legal]',
    org='[PF: Profesional Titular][PM: Órgano de administración del despacho]',
    d2=('[PF: Constancia de Asunción Personal del profesional][PM: Representante Encargado de Cumplimiento designado (persona física)]', 'art. 20 de la Ley; art. 10 RCG',
        '[PF: Firmar la Constancia de Asunción Personal (D2-PROF-PF)][PM: Celebrar la sesión y firmar el Acta de Designación (D2-PROF-PM)]', 'D2-PROF'),
    avisos=('Avisos por cada operación financiera en nombre y representación del cliente e informes en ceros', 'art. 17, fr. XI, y 23 de la Ley; arts. 24 a 25 RCG'),
    acum='Documentar que la fracción XI genera un Aviso por cada operación financiera en representación, sin umbral ni acumulación',
    efectivo='Verificar que en los actos que se preparen o ejecuten no se pacte el pago en efectivo prohibido por el art. 32 (inmuebles de 8,025 UMA o más; acciones de 3,210 UMA o más)',
    extra=[('Distinción entre preparación y operación financiera en representación; secreto profesional', 'art. 17, fr. XI, y 22 de la Ley', M65,
            'Clasificar cada servicio por momento de sujeción y documentar el criterio de secreto profesional', 'D1-PROF, secciones 2.2 y IV', 'Semana 1-2')]),
}

def rubros(d, D):
    s = d['suf']
    return [  # (rubro, ref, multa, acción, doc, plazo, responsable)
        ('Alta y registro en el padrón de Actividades Vulnerables', 'art. 18, fr. IV Bis, de la Ley; art. 4 RCG', M18, 'Iniciar el trámite de alta en el Portal en Internet', D('3'), 'Semana 1', d['RL']),
        (d['d2'][0], d['d2'][1], M18, d['d2'][2], d['d2'][3], 'Semana 1', d['RL']),
        ('(*) Manual de Políticas Internas (14 fracciones del art. 37 Bis RCG)', 'art. 18, fr. VIII, de la Ley; Cap. X RCG', M18, 'Adoptar e implementar el Manual de Políticas Internas', D('5'), 'Semana 2-3', d['R']),
        ('Expedientes de identificación de clientes', 'art. 18, fr. I y II, de la Ley; art. 12 RCG', M18, 'Integrar los expedientes de los clientes actuales', D('7'), 'Semana 2-4', d['R']),
        ('(*) Metodología EBR documentada e implementada', 'art. 18, fr. VII, de la Ley; Cap. II Quáter RCG', M18, 'Documentar y aplicar la Metodología EBR', D('6'), 'Semana 2', d['R']),
        ('(*) Clasificación del Grado de Riesgo del cliente', 'Cap. III Bis RCG', M18, 'Clasificar la cartera de clientes en bajo, medio y alto', '%s / %s' % (D('6'), D('7')), 'Antes del 1 mar 2027', d['R']),
        ('(*) Identificación del Beneficiario Controlador', 'art. 18, fr. III, de la Ley; Cap. III Quinquies RCG', M18, 'Identificar y documentar al Beneficiario Controlador', D('7') + (' / D14' if s == '' else ''), 'Antes del 1 mar 2027', d['R']),
        ("Política de PEP's", 'art. 18, fr. VIII, de la Ley; Cap. III Quáter RCG', M18, "Adoptar la Política de PEP's", D('8'), 'Semana 2', d['R']),
        ('(*) Perfil Transaccional y sistema de alertas', 'Cap. III Ter RCG', M18, 'Determinar el perfil inicial de cada cliente e implementar las alertas', '%s, cap. 14 / %s' % (D('5'), D('7')), 'Antes del 1 mar 2027', d['R']),
        ('Avisos de 24 horas por sospecha o indicios', 'art. 18, fr. VI, de la Ley; arts. 26 Bis a 26 Bis 2 RCG', M65, 'Ajustar el proceso interno al plazo de 24 horas', D('9'), 'Semana 3-4', d['R']),
        (d['avisos'][0], d['avisos'][1], M65, 'Implementar el control de Avisos mensuales', D('12'), 'Semana 3-4', d['R']),
        ('(*) Seguimiento y acumulación de actos u operaciones', 'art. 17, penúltimo párrafo, de la Ley; art. 19 RCG', M18, d['acum'], '%s, cap. 15' % D('5'), 'Antes del 1 mar 2027', d['R']),
        ('Programa anual de capacitación', 'art. 18, fr. IX, de la Ley; Cap. XII RCG', M18, 'Ejecutar el primer ciclo de capacitación', D('11'), 'Mes 2', d['R']),
        ('(*) Selección de personal (declaración firmada)', 'art. 39 Bis 2 RCG', M18, 'Recabar la declaración en las nuevas contrataciones', D('11D'), 'A partir del 1 mar 2027', d['R']),
        ('Resguardo documental (plazos diferenciados)', 'art. 18, fr. IV, de la Ley; arts. 39 Bis 1, 41 y 51 RCG', M18, 'Adoptar la Política de Resguardo Documental', D('10'), 'Semana 3', d['R']),
        ('(*) Mecanismos automatizados', 'art. 18, fr. X, de la Ley; art. 41 RCG', M18, 'Habilitar mecanismos automatizados con las seis funciones del art. 41 RCG', '%s, cap. 19' % D('5'), 'Antes del 1 jun 2027', d['R']),
        ('(*) Auditoría anual (primer periodo: 2028)', 'art. 18, fr. XI, de la Ley; Cap. XIV RCG', M18, 'Designar área o persona auditora independiente según el Riesgo (arts. 44 y 45 RCG)', D('13'), 'Dictamen a más tardar el último día hábil de marzo de 2029', d['org']),
        ('(*) Identificación de personas en listas de autoridades', 'arts. 27 y 38 RCG', M65, 'Incorporar la consulta de listas al proceso de identificación', '%s, cap. 16; %s, Sección H' % (D('5'), D('7')), 'Antes del 1 mar 2027', d['R']),
        ('(*) Confidencialidad', 'art. 50 de la Ley; art. 31 RCG', M18, 'Formalizar las medidas de confidencialidad', '%s, cap. 17' % D('5'), 'Semana 2-3', d['R']),
        ('Restricción del uso de efectivo', 'art. 32 de la Ley', M65, d['efectivo'], '%s, cap. 9' % D('5'), 'Semana 2', d['R']),
    ] + [(r, ref, m, a, doc, pl, d['R']) for r, ref, m, a, doc, pl in d['extra']]

def construir(kit):
    d = S[kit]; ctx = Ctx(kit)
    D = lambda n: 'D%s%s' % (n, d['suf'])
    filas = rubros(d, D); n = len(filas)
    portada(ctx, codigo='D4', titulo='Gap Analysis PLD/FT',
            indice='Índice: I. Instrucciones │ II. Matriz de brecha │ III. Resumen ejecutivo │ IV. Plan de acción')
    sec(ctx, 'I')
    p(ctx, 'Esta matriz evalúa %d rubros de cumplimiento PLD/FT: los 20 rubros comunes a todos los kits del Programa y %d '
           'propio%s del sector. Para cada rubro se marca el estatus (C cumple, P parcial, N no cumple), el semáforo (V verde, A '
           'amarillo, R rojo) y la multa máxima estimada, y en la sección IV se asigna la acción correctiva. Los rubros '
           'marcados con (*) son nuevos o cambiaron con el Acuerdo 115/2026 (DOF 07/08/2026).'
           % (n, n - 20, 's' if n - 20 != 1 else ''))
    sec(ctx, 'II', ' — %d Rubros' % n)
    tabla(ctx, [('No.', 'Rubro de cumplimiento', 'Ref. legal', 'Estatus', 'Semáforo', 'Multa máx.')] +
          [(str(i), r[0], r[1], '[ ]C [ ]P [ ]N', '[ ]V [ ]A [ ]R', r[2]) for i, r in enumerate(filas, 1)],
          [500, 3300, 2100, 1100, 1100, 1400])
    p(ctx, 'Multas máximas del art. 54 de la Ley con UMA 2026 ($117.31): 2,000 UMA = $234,620 (incumplir el art. 18 o Avisos '
           'con hasta 30 días de retraso); 10,000 UMA = $1,173,100 (forma de pago en instrumentos, art. 33; arts. 53, '
           'fr. V, y 54, fr. II); 65,000 UMA = $7,625,150 o del 10% al 100% del valor del acto, la que resulte mayor (omisión '
           'de Avisos y efectivo prohibido).')
    sec(ctx, 'III')
    tabla(ctx, [('**Total de rubros evaluados:**', str(n)), ('**Rubros en cumplimiento (verde):**', '[ ]'),
                ('**Rubros en cumplimiento parcial (amarillo):**', '[ ]'), ('**Rubros en incumplimiento (rojo):**', '[ ]'),
                ('**De ellos, rubros (*) del Acuerdo 115/2026:**', '[ ]'),
                ('**Multa máxima acumulada estimada:**', '$[ ] MXN'),
                ('**Beneficio por actuación espontánea (art. 55):**', 'Abstención por única ocasión o reducción de hasta 50%: $[ ] MXN')],
          [4200, 5300], encabezado=False)
    sec(ctx, 'IV')
    tabla(ctx, [('Rubro', 'Acción correctiva', 'Documento', 'Plazo sugerido', 'Responsable')] +
          [(str(i), r[3], r[4], r[5], r[6]) for i, r in enumerate(filas, 1)],
          [750, 3650, 1700, 1500, 1900])
    pie(ctx)
    return ctx

NOMBRES = {'PROF': 'profesionistas/D4-PROF_Gap_Analysis_PLD_v2.docx', 'INM': 'inmobiliaria/D4_Gap_Analysis_PLD_v5.docx',
           'NOT': 'notarial/D4-NOT_Gap_Analysis_PLD_v3.docx', 'COR': 'corredores/D4-COR_Gap_Analysis_PLD_v2.docx'}

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
        cp.title = 'D4%s Gap Analysis PLD/FT — %s' % (canon.KITS[k]['suf'], canon.KITS[k]['variante'])
        cp.author = 'Certus Consultores · Espinosa-Salazar Abogados Corporativos'
        cp.subject = 'Programa Certus PLD — edición homologada noviembre 2026'
        dest = os.path.join(OUT, NOMBRES[k]); os.makedirs(os.path.dirname(dest), exist_ok=True)
        from docxlib import numero_de_pagina; numero_de_pagina(doc)
        doc.save(dest); print(k, dest)
