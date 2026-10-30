// D15 — Protocolo de Atención de Visitas de Verificación — generador por variante de Kit
// Uso: node d15v.js <inmobiliaria|notarial|corredores|prof-pf|prof-pm> <salida.docx>
const fs = require('fs');
const L = require('./lib');
const { d, P, H1, H2, B, N, CB, T, KV, BOX, SP, BR, W } = L;
const { Document, Packer, Paragraph, TextRun, Footer, AlignmentType, PageNumber } = d;

// ───────────── Configuración por variante ─────────────
const V = {
  inmobiliaria: {
    code: 'D15', sfx: '', pf: false,
    scope: 'Programa Certus PLD — Inmobiliarias (Fracc. V, V Bis, XV)',
    sujeto: '[Empresa]',
    titular: 'la persona Representante Encargada del Cumplimiento',
    responsable: 'Persona Representante Encargada del Cumplimiento: [Nombre del Representante de Cumplimiento]',
    obligados: 'el órgano de administración, la persona Representante Encargada del Cumplimiento, sus suplentes y todo el personal que pueda tener el primer contacto con la autoridad (recepción, ventas, administración)',
    primerContacto: 'recepción, ventas, administración',
    firmante: '[NOMBRE DEL REPRESENTANTE LEGAL]', firmanteCalidad: 'en representación de [Empresa], personalidad que acredito con [instrumento]',
    d2: ['Designación y aceptación de la persona Representante Encargada del Cumplimiento', 'D2', 'Art. 20'],
    bc: 'D7 / D14', selPers: 'D11D',
  },
  notarial: {
    code: 'D15-NOT', sfx: '-NOT', pf: true,
    scope: 'Programa Certus PLD — Notarías (Fracc. XII, Apartado A)',
    sujeto: 'la Notaría Pública a cargo de [Nombre del Notario Titular]',
    titular: 'la persona notaria titular',
    responsable: 'Notario(a) Titular: [Nombre del Notario Titular]',
    obligados: 'la persona notaria titular, la persona encargada operativa (si se designó), el personal de la notaría y cualquier persona que pueda tener el primer contacto con la autoridad (recepción, mesa de control, archivo)',
    primerContacto: 'recepción, mesa de control, archivo',
    firmante: '[Nombre del Notario Titular]', firmanteCalidad: 'por mi propio derecho, en mi carácter de titular de la Notaría Pública número [número] de [plaza]',
    d2: ['Constancia de asunción personal del cumplimiento por la persona notaria titular', 'D2-NOT', 'Art. 20, último párr.'],
    bc: 'D7-NOT', selPers: 'D11-NOT',
  },
  corredores: {
    code: 'D15-COR', sfx: '-COR', pf: true,
    scope: 'Programa Certus PLD — Correduría Pública (Fracc. XII, Apartado B)',
    sujeto: 'la Correduría Pública a cargo de [Nombre del Corredor Público Titular]',
    titular: 'la persona corredora pública titular',
    responsable: 'Corredor(a) Público(a) Titular: [Nombre del Corredor Público Titular]',
    obligados: 'la persona corredora pública titular, la persona encargada operativa (si se designó), el personal de la correduría y cualquier persona que pueda tener el primer contacto con la autoridad (recepción, archivo, valuación)',
    primerContacto: 'recepción, archivo, valuación',
    firmante: '[Nombre del Corredor Público Titular]', firmanteCalidad: 'por mi propio derecho, en mi carácter de corredor público con habilitación en la plaza de [Entidad federativa]',
    d2: ['Constancia de asunción personal del cumplimiento por la persona corredora pública titular', 'D2-COR', 'Art. 20, último párr.'],
    bc: 'D7-COR', selPers: 'D11-COR',
  },
  'prof-pf': {
    code: 'D15-PROF-PF', sfx: '-PROF', pf: true,
    scope: 'Programa Certus PLD — Profesionistas, persona física (Fracc. XI)',
    sujeto: '[Nombre del profesional]',
    titular: 'el profesional titular',
    responsable: 'Profesional Titular: [Nombre del profesional]',
    obligados: 'el profesional titular, la persona de apoyo operativo (si se designó) y cualquier colaborador que pueda tener el primer contacto con la autoridad (recepción, asistentes)',
    primerContacto: 'recepción, asistentes',
    firmante: '[Nombre del profesional]', firmanteCalidad: 'por mi propio derecho, en mi carácter de profesional independiente',
    d2: ['Constancia de asunción personal del cumplimiento por el profesional', 'D2-PROF-PF', 'Art. 20, último párr.'],
    bc: 'D7-PROF', selPers: 'D11-PROF',
  },
  'prof-pm': {
    code: 'D15-PROF-PM', sfx: '-PROF', pf: false,
    scope: 'Programa Certus PLD — Profesionistas, despacho persona moral (Fracc. XI)',
    sujeto: '[Empresa]',
    titular: 'la persona Representante Encargada del Cumplimiento',
    responsable: 'Persona Representante Encargada del Cumplimiento: [Nombre del Representante de Cumplimiento]',
    obligados: 'el órgano de administración del despacho, la persona Representante Encargada del Cumplimiento, sus suplentes y todo el personal que pueda tener el primer contacto con la autoridad (recepción, asistentes, socios y asociados)',
    primerContacto: 'recepción, asistentes',
    firmante: '[NOMBRE DEL REPRESENTANTE LEGAL]', firmanteCalidad: 'en representación de [Empresa], personalidad que acredito con [instrumento]',
    d2: ['Designación y aceptación de la persona Representante Encargada del Cumplimiento', 'D2-PROF-PM', 'Art. 20'],
    bc: 'D7-PROF', selPers: 'D11-PROF',
  },
};

// Particularidades por variante (sección XIII)
const PART = {
  inmobiliaria: [
    '**Estatus de la designación.** Mientras no haya persona Representante Encargada del Cumplimiento o su designación no esté aceptada, las obligaciones recaen en los integrantes del órgano de administración o en el administrador único (art. 20, 2º párr.). Verificar el estatus en el Portal como parte de la preparación.',
    '**Desarrolladores (Fracc. V Bis).** La actividad vulnerable es la recepción de recursos destinados a un Desarrollo Inmobiliario para su venta o renta. La Carpeta de Visita debe separar, por desarrollo, los recursos recibidos, su origen y su aplicación.',
    '**Arrendamiento (Fracc. XV).** El monto relevante es el valor mensual de la renta (Reglamento art. 31). Si la renta se pacta en otra periodicidad, conservar el cálculo del valor mensual que se utilizó.',
    '**Uso de efectivo.** Tener a la mano la evidencia de la forma de pago de cada operación, pues la verificación revisa también la prohibición del art. 32 (inmuebles: 8,025 UMA; uso o goce: 3,210 UMA mensuales). Su incumplimiento se sanciona en el rango más alto del art. 54, fracc. III.',
    '**Acumulación.** Conservar el control de acumulación por cliente en periodos de hasta seis meses (art. 17, penúltimo párr.; Reglamento art. 7), ya que es uno de los puntos que más se revisan en avisos omitidos.',
  ],
  notarial: [
    '**Cumplimiento personal.** La persona notaria cumple personal y directamente (art. 20, último párr.); no existe Representante registrable. La persona encargada operativa, si existe, apoya pero no sustituye al titular. Si el titular no está, la visita se atiende con quien se encuentre en la notaría y el titular debe ser localizado de inmediato.',
    '**Secreto profesional.** La entrega de información y documentación conforme a la Ley no transgrede el secreto profesional ni las obligaciones de confidencialidad (art. 22). No es causa para negar la documentación directamente relacionada con los actos del Apartado A; sí lo es para no entregar lo que no se relaciona con ellos (art. 34).',
    '**Doble vía sancionadora.** Cuando el SAT determine notorias deficiencias (reincidencia en el art. 53, fracc. I a V, o las conductas de las fracc. VI y VII), lo informa a la autoridad que supervisa la fe pública para que inicie su propio procedimiento (art. 58). Involucrar a la defensa jurídica desde el día de la visita.',
    '**Avisos del inciso a).** Los avisos por transmisión o constitución de derechos reales sobre inmuebles se presentan por los medios fiscales (art. 24, último párr.; Reglamento art. 17). Integrar a la Carpeta de Visita los acuses de esa vía junto con los del Portal.',
    '**Fecha del acto y forma de pago.** La fecha relevante para el aviso es la de otorgamiento del instrumento (Reglamento art. 5, 2º párr.). En operaciones de 8,025 UMA o más, el instrumento debe identificar la forma de pago (art. 33): tener ubicados esos instrumentos y sus comprobantes.',
    '**Umbrales reformados.** Para los incisos a), c) y d) aplican los nuevos umbrales o escenarios a partir de los actos realizados el 17 de julio de 2025 (Reglamento, Transitorio Sexto, fracc. IV). Conservar el análisis por periodo.',
  ],
  corredores: [
    '**Cumplimiento personal.** La persona corredora pública cumple personal y directamente (art. 20, último párr.). Si no está presente, la visita se atiende con quien se encuentre en la correduría y el titular debe ser localizado de inmediato.',
    '**Riesgo sobre la habilitación.** La reincidencia en las infracciones del art. 53, fracc. I a IV, es causa de cancelación definitiva de la habilitación (art. 57); además, las notorias deficiencias se informan a la autoridad que supervisa la fe pública (art. 58). Toda multa debe valorarse con la defensa antes de dejarla firme.',
    '**Avalúos.** El inciso a) del Apartado B aplica a los avalúos en los que se utiliza la fe pública (Reglamento art. 29). Separar en la Carpeta de Visita los avalúos con fe pública de los demás.',
    '**Secreto profesional.** La entrega de información conforme a la Ley no transgrede el secreto profesional (art. 22); se limita a lo directamente relacionado con los actos del Apartado B (art. 34).',
    '**Fecha del acto y forma de pago.** La fecha relevante es la del otorgamiento o protocolización (Reglamento art. 5, 2º párr.). En operaciones de 8,025 UMA o más, identificar la forma de pago en el instrumento (art. 33).',
  ],
  'prof-pf': [
    '**Cumplimiento personal.** El profesional cumple personal y directamente (art. 20, último párr.). Si no está presente, la visita se atiende con quien se encuentre en la oficina; debe ser localizado de inmediato.',
    '**Preparación y representación.** La identificación aplica cuando se prepara para un cliente o se ejecuta en su nombre alguna de las operaciones de los incisos a) a e); el aviso, sólo cuando se lleva a cabo en nombre y representación del cliente una operación financiera relacionada (art. 17, fracc. XI; Reglamento art. 27). La Carpeta de Visita debe distinguir ambos supuestos por asunto.',
    '**Secreto profesional y garantía de defensa.** La obligación de aviso se cumple con respeto al secreto profesional y a la garantía de defensa (art. 17, fracc. XI, último párr.). Identificar previamente qué documentación está amparada por ese régimen y, si se solicita, hacerlo constar en el acta y consultarlo con la defensa antes de entregarla.',
    '**Expedientes de asuntos.** Entregar sólo los expedientes de asuntos que encuadren en los incisos de la fracción XI y del periodo de cinco años; los demás asuntos del despacho no son materia de la verificación (art. 34).',
  ],
  'prof-pm': [
    '**Estatus de la designación.** Mientras no haya persona Representante Encargada del Cumplimiento o su designación no esté aceptada, las obligaciones recaen en el órgano de administración del despacho (art. 20, 2º párr.).',
    '**Preparación y representación.** La identificación aplica cuando se prepara para un cliente o se ejecuta en su nombre alguna de las operaciones de los incisos a) a e); el aviso, sólo cuando se lleva a cabo en nombre y representación del cliente una operación financiera relacionada (art. 17, fracc. XI; Reglamento art. 27). La Carpeta de Visita debe distinguir ambos supuestos por asunto y por socio o asociado responsable.',
    '**Secreto profesional y garantía de defensa.** La obligación de aviso se cumple con respeto al secreto profesional y a la garantía de defensa (art. 17, fracc. XI, último párr.). Identificar previamente qué documentación está amparada por ese régimen y, si se solicita, hacerlo constar en el acta y consultarlo con la defensa antes de entregarla.',
    '**Expedientes de asuntos.** Entregar sólo los expedientes de asuntos que encuadren en los incisos de la fracción XI y del periodo de cinco años; los demás asuntos del despacho no son materia de la verificación (art. 34).',
  ],
};

function build(key) {
  const v = V[key];
  const s = v.sfx; // sufijo de documentos del Kit
  const C = [];
  C.push(L.kitHeader(v.code, 'Protocolo de Atención de Visitas de Verificación y Requerimientos de Autoridad', v.scope));
  C.push(SP(200));
  C.push(KV([
    [v.pf ? 'Sujeto obligado:' : 'Empresa:', v.pf ? v.sujeto : '[Empresa]'],
    ['Fecha de adopción:', '[Fecha]'],
    ['Elaboró:', 'CERTUS CONSULTORES │ Certus PLD'],
    ['Versión:', '1.0 — emisión inicial (octubre 2026)'],
    ['Marco normativo considerado:', 'LFPIORPI (última reforma DOF 16-07-2025); Reglamento (última reforma DOF 27-03-2026); Reglas de Carácter General reformadas por el Acuerdo 115/2026 (DOF 07-08-2026, en vigor 30-11-2026); Ley Federal de Procedimiento Administrativo (supletoria).'],
    ['Integración al Manual:', v.pf
      ? `Este protocolo forma parte del Manual de Políticas Internas (D5${s}) por referencia, como mecanismo de control interno (RCG art. 37, 2º párrafo, y art. 37 Bis, fracc. XII).`
      : `Este protocolo forma parte del Manual de Políticas Internas (D5${s}) por referencia, como mecanismo de control interno y función de la persona Representante Encargada del Cumplimiento (RCG art. 37, 2º párrafo, y art. 37 Bis, fracc. X y XII).`],
  ]));

  // I
  C.push(H1('I. Objeto y alcance'));
  C.push(P(`Este protocolo establece cómo debe actuar ${v.sujeto} cuando la Secretaría de Hacienda y Crédito Público, por conducto del SAT, ejerza sus facultades de comprobación en materia de prevención de lavado de dinero, ya sea mediante una **visita de verificación** en el domicilio o mediante un **requerimiento de información** por escrito o a través del Portal de Prevención de Lavado de Dinero.`));
  C.push(P('Su propósito es doble: **cumplir sin obstaculizar** (la falta de facilidades es en sí misma una infracción) y **preservar los derechos de defensa** desde el primer minuto, porque lo que ocurre durante la diligencia y en los cinco días siguientes determina en gran medida el resultado de cualquier procedimiento posterior.'));
  C.push(P(`Es de observancia obligatoria para ${v.obligados}.`));

  // II
  C.push(H1('II. Fundamento legal'));
  const fund = [
    ['LFPIORPI art. 18, fracc. V', 'Obligación de brindar las facilidades necesarias para que se lleven a cabo las visitas de verificación.'],
    ['LFPIORPI art. 22', 'Entregar información y documentación conforme a la Ley no implica transgresión a obligaciones de confidencialidad o secreto legal, profesional, fiscal, bancario o fiduciario.'],
    ['LFPIORPI art. 25', 'La autoridad puede requerir, por escrito o durante la visita, el soporte documental de las operaciones con Clientes o Usuarias y de los Avisos e Informes.'],
    ['LFPIORPI art. 34', 'Comprobación de oficio y en cualquier tiempo, por visita o por requerimiento. El visitado o requerido entrega **exclusivamente** la información directamente relacionada con las Actividades Vulnerables que realiza.'],
    ['LFPIORPI art. 35', 'Las visitas, requerimientos y sanciones se rigen por la Ley, su Reglamento y las Reglas de Carácter General; la LFPA aplica de manera supletoria. No aplica el Código Fiscal de la Federación como régimen procesal.'],
    ['LFPIORPI art. 36', 'La verificación sólo puede abarcar actos u operaciones realizados en los **cinco años** inmediatos anteriores al inicio de la visita.'],
    ['LFPIORPI art. 37', 'La autoridad puede solicitar el auxilio de la fuerza pública.'],
  ];
  if (!v.pf) fund.push(['LFPIORPI art. 41 Bis', 'Las diligencias que requieran comparecencia de la persona Representante Encargada del Cumplimiento pueden desahogarse por representantes legales o apoderados, para proteger su identidad.']);
  fund.push(
    ['LFPIORPI arts. 52–54', 'Las multas tienen carácter de créditos fiscales. No atender requerimientos o incumplir cualquier obligación del art. 18: 200 a 2,000 UMA por infracción.'],
    ['LFPIORPI art. 55; Reglamento art. 55 Bis', 'Abstención de sanción (única ocasión) si se regularizó de forma espontánea antes de la visita y se reconoce la falta dentro del plazo inicial (5 días, art. 68 LFPA). Si ya se usó el beneficio: reducción de hasta 50%.'],
  );
  if (key === 'notarial') fund.push(['LFPIORPI art. 58', 'Las notorias deficiencias de la persona notaria se informan a la autoridad que supervisa la fe pública, para que instaure su propio procedimiento sancionador.']);
  if (key === 'corredores') fund.push(['LFPIORPI arts. 57 y 58', 'La reincidencia en el art. 53, fracc. I a IV, es causa de cancelación definitiva de la habilitación; las notorias deficiencias se informan a la autoridad que supervisa la fe pública.']);
  fund.push(
    ['LFPIORPI art. 61', 'Impugnación: recurso de revisión (LFPA) o juicio contencioso administrativo ante el TFJA.'],
    ['LFPIORPI art. 62, fracc. II y III', 'Delito (2 a 8 años) alterar información o incorporar información ilegible al desahogar requerimientos; admite comisión culposa.'],
    ['Reglamento art. 4, fracc. III', 'La visita o requerimiento se practica en el domicilio manifestado en el alta como Actividad Vulnerable; si no se localiza, en el domicilio fiscal (RFC).'],
    ['Reglamento arts. 8, 9 y 10 Bis', 'Requerimientos: 10 días hábiles, prórroga de 5. Oficio de observaciones con 5 días para desvirtuar; los hechos no desvirtuados se tienen por consentidos. Información de terceros: 10 días para manifestarse.'],
    ['RCG arts. 5 Bis y 6', 'Notificaciones electrónicas por el Portal: 3 días hábiles para abrir; se tiene por notificado al 4º día hábil. Consulta obligatoria al menos una vez cada día hábil.'],
    ['RCG arts. 36 Bis, 36 Bis 1 y 36 Bis 2', 'Diligencias en días y horas hábiles y cómputo de plazos; forma de presentar escritos ante el SAT; citatorio por instructivo y acta de hechos ante oposición o domicilio cerrado.'],
    ['LFPA arts. 62–68', 'Orden escrita con firma autógrafa (objeto, alcance, fundamento); obligación de dar acceso e informes; credencial vigente con fotografía; acta circunstanciada ante dos testigos; observaciones y pruebas en el acto o dentro de 5 días.'],
    ['LFPA art. 72', 'Antes de sancionar: notificación de inicio del procedimiento y 15 días para manifestar y ofrecer pruebas.'],
  );
  C.push(T(['Disposición', 'Contenido relevante para la atención'], fund, [2700, W - 2700]));

  // III
  C.push(H1('III. Principios de actuación'));
  [
    '**Cooperar no es renunciar.** Se da acceso y se atiende a la autoridad con cortesía; las objeciones se hacen constar por escrito, nunca impidiendo la diligencia.',
    '**Sólo lo directamente relacionado.** Se entrega lo que pide la orden y que corresponde a la Actividad Vulnerable y al periodo de cinco años (arts. 34 y 36). Nada se ofrece por iniciativa propia.',
    '**Nada se altera, nada se destruye.** Los documentos se entregan tal como existen. Una corrección hecha durante la visita puede configurar el delito del art. 62.',
    '**Todo queda por escrito.** Cada documento entregado se registra en el inventario (Anexo B) y cada desacuerdo en el acta o en el escrito de observaciones.',
    `**Un solo interlocutor.** Sólo ${v.titular} (o quien lo supla en la atención) dialoga con la autoridad sobre el fondo. El resto del personal se limita a identificarse y facilitar.`,
    '**Escalamiento inmediato.** El asesor de cumplimiento y la defensa jurídica se enteran en la primera hora, no al final del día.',
  ].forEach(t => C.push(N(t, 'num')));

  // IV
  C.push(H1('IV. Roles y responsables'));
  const roles = [
    ['Primer contacto', '[Nombre y puesto del primer contacto]', 'Recibe a la autoridad, aplica la Ficha de Recepción (Anexo A) y avisa de inmediato al responsable de atención.'],
    ['Responsable de atención', v.responsable, 'Revisa la orden, conduce la diligencia, propone testigos, entrega documentación, revisa y firma el acta.'],
  ];
  if (v.pf) {
    roles.push(
      ['Encargado(a) operativo(a)', '[Nombre del encargado operativo]', `Apoya a ${v.titular}; atiende la diligencia mientras se localiza al titular, sin sustituirlo en sus obligaciones.`],
      ['Apoderado(a)', '[Nombre del apoderado]', `Puede comparecer en representación de ${v.titular} con poder suficiente.`],
    );
  } else {
    roles.push(
      ['Suplente 1', '[Nombre y puesto del suplente 1]', 'Asume la atención si el responsable no está presente.'],
      ['Suplente 2', '[Nombre y puesto del suplente 2]', 'Asume la atención si ninguno de los anteriores está presente.'],
      ['Representante legal / apoderado', '[NOMBRE DEL REPRESENTANTE LEGAL]', 'Comparece en lugar de la persona Representante Encargada del Cumplimiento cuando se requiera proteger su identidad (art. 41 Bis).'],
    );
  }
  roles.push(
    ['Asesor de cumplimiento', 'CERTUS CONSULTORES │ 999 163 1363 │ contacto@certusconsultores.com.mx', 'Acompaña (presencial o remoto), revisa la orden y el alcance, prepara la Carpeta de Visita y el escrito de observaciones.'],
    ['Defensa jurídica', 'ESPINOSA-SALAZAR ABOGADOS CORPORATIVOS │ 999 575 3129 │ contacto@espinosasalazarabogados.com', 'Valora vicios de la orden y del acta, define estrategia de reconocimiento (art. 55) y conduce el procedimiento sancionador e impugnaciones.'],
    ['Testigos de confianza', '1. [Nombre del testigo 1]   2. [Nombre del testigo 2]', 'Personas previamente identificadas que pueden fungir como testigos del acta (LFPA art. 66).'],
  );
  C.push(T(['Rol', 'Persona designada', 'Funciones en la visita'], roles, [2100, 2900, W - 5000]));
  C.push(SP(80));
  C.push(BOX(v.pf ? 'Ausencia del titular' : 'Ausencia del responsable', [
    `La visita **no se suspende** porque no esté ${v.titular}: se entiende con quien se encuentre en el establecimiento (LFPA arts. 64 y 67, fracc. V). Se puede pedir con cortesía un margen breve para localizarle, **sin condicionar el acceso** a su llegada; negarse a recibir la visita por esa razón puede interpretarse como falta de facilidades (art. 18, fracc. V).`,
  ]));

  // V
  C.push(H1('V. Fase 1 — Preparación permanente (antes de cualquier visita)'));
  C.push(P('La calidad de la respuesta se decide antes de que llegue la autoridad. Esta lista se revisa al menos cada trimestre y se documenta en la Bitácora (sección XII).'));
  [
    '**Domicilio actualizado.** El domicilio manifestado en el alta como Actividad Vulnerable coincide con el lugar donde se opera y se resguarda la información (Reglamento art. 4, fracc. III). Cualquier cambio se actualiza dentro de los 6 días hábiles siguientes (RCG art. 7).',
    '**Portal consultado cada día hábil.** Una persona designada revisa el buzón del Portal de Prevención de Lavado de Dinero al menos una vez cada día hábil y deja evidencia (RCG art. 6). Responsable: [Nombre del responsable del Portal].',
    '**Carpeta de Visita integrada** conforme a la matriz de la sección XI, por periodo mensual, para los últimos cinco años.',
    '**Ficha de Recepción (Anexo A)** impresa y visible para el personal de primer contacto.',
    `**Roles de la sección IV** asignados, con teléfonos vigentes${v.pf ? '' : ' y suplentes capacitados'}.`,
    `**Simulacro anual** de visita, con acta interna de hallazgos, como parte del programa de capacitación (D11${s}).`,
    '**Regularización previa.** Toda omisión detectada (aviso no presentado, expediente incompleto) se corrige de inmediato y se documenta su fecha: sólo lo regularizado **antes** del inicio de la visita puede acceder a la abstención del art. 55 (ver sección VIII).',
  ].forEach(t => C.push(CB(t)));

  // VI
  C.push(H1('VI. Fase 2 — Llegada de la autoridad (primeros 15 minutos)'));
  [
    'Recibir a las personas verificadoras con cortesía y en un espacio adecuado. **No negar el acceso.**',
    'Solicitar que exhiban **credencial vigente con fotografía** expedida por la autoridad competente y la **orden escrita** de visita (LFPA arts. 63 y 65). Anotar nombres y números de credencial.',
    'Pedir la **copia de la orden** que debe dejarse al visitado (LFPA art. 65). Fotografiar o escanear la orden y la credencial si la autoridad lo permite.',
    'Avisar de inmediato al responsable de atención, al asesor de cumplimiento y a la defensa jurídica, y enviarles la imagen de la orden.',
    'El responsable revisa la orden con la lista de verificación siguiente **antes** de entregar cualquier documento.',
  ].forEach(t => C.push(N(t, 'num2')));
  C.push(H2('Lista de verificación de la orden'));
  C.push(T(['Elemento', 'Qué revisar', 'Fundamento'], [
    ['Forma', 'Orden escrita con firma autógrafa de la autoridad emisora.', 'LFPA art. 63'],
    ['Autoridad', 'Unidad administrativa del SAT que la emite y su fundamento de competencia.', 'LFPIORPI art. 35; Reglamento art. 4, fracc. III'],
    ['Destinatario', v.pf ? 'Nombre correcto del titular y, en su caso, número y plaza.' : 'Nombre o denominación correctos.', 'LFPA arts. 63 y 67, fracc. I'],
    ['Lugar', 'Domicilio del alta como Actividad Vulnerable o, en su defecto, domicilio fiscal.', 'Reglamento art. 4, fracc. III'],
    ['Objeto y alcance', 'Obligaciones a verificar y documentos solicitados. Delimita lo que puede revisarse.', 'LFPA art. 63'],
    ['Periodo', 'No más de cinco años anteriores a la fecha de inicio de la visita.', 'LFPIORPI art. 36'],
    ['Fundamento', 'Disposiciones legales que sustentan la visita.', 'LFPA art. 63'],
    ['Día y hora', 'Diligencia en día y hora hábiles; si es fuera de ellos, verificar que la orden lo habilite.', 'RCG art. 36 Bis; LFPA art. 62'],
    ['Personal', 'Credencial vigente con fotografía de cada verificador.', 'LFPA art. 65'],
  ], [1600, 5260, W - 6860]));
  C.push(SP(80));
  C.push(BOX('Si algo falta o no coincide', [
    'No se impide la diligencia. Se hace constar en el acta, en el apartado de declaración del visitado (LFPA art. 67, fracc. VIII), y se informa a la defensa jurídica. Los vicios de la orden se hacen valer después, por escrito y en los medios de defensa, no en la puerta.',
  ], 'navy'));

  // VII
  C.push(H1('VII. Fase 3 — Desarrollo de la diligencia'));
  [
    `**Testigos.** Proponer a dos testigos de confianza (sección IV). Si no se proponen, los designa quien practica la visita (LFPA art. 66). Conviene que los testigos no sean ${v.titular}.`,
    '**Entrega de documentación.** Entregar únicamente lo solicitado en la orden y relacionado con la Actividad Vulnerable (art. 34, 2º párrafo). Preferir copias; si se solicitan originales para cotejo, que se devuelvan en el acto y se asiente.',
    '**Inventario.** Registrar cada documento entregado en el Formato de Inventario (Anexo B) y solicitar que se relacione en el acta.',
    '**Información no disponible en el momento.** Manifestar que se entregará, pedir que se asiente y fijar el compromiso por escrito. Nunca improvisar ni reconstruir documentos en el momento.',
    '**Declaraciones.** Responder con hechos y de forma breve. No emitir opiniones, no reconocer incumplimientos durante la visita y no especular sobre operaciones de clientes. Cualquier reconocimiento se hace por escrito, con asesoría, dentro del plazo (sección VIII).',
    '**Acompañamiento.** La persona responsable acompaña en todo momento a los verificadores dentro de las instalaciones.',
    '**Trato.** Mantener un tono cordial. Ante cualquier tensión, pausar y llamar a la defensa; la autoridad puede solicitar el auxilio de la fuerza pública (art. 37).',
  ].forEach(t => C.push(B(t)));

  // VIII
  C.push(H1('VIII. Fase 4 — El acta de verificación y los días siguientes'));
  C.push(H2('Revisión del acta antes de firmar'));
  C.push(P('El acta debe contener, como mínimo, los elementos del art. 67 de la LFPA. Leerla completa antes de firmar:', { keepNext: true }));
  C.push(T(['Elemento (LFPA art. 67)', 'Revisado'], [
    ['I. Nombre, denominación o razón social del visitado', '☐'],
    ['II. Hora, día, mes y año de inicio y conclusión', '☐'],
    ['III. Domicilio completo donde se practicó la visita', '☐'],
    ['IV. Número y fecha del oficio de comisión', '☐'],
    ['V. Nombre y cargo de la persona con quien se entendió la diligencia', '☐'],
    ['VI. Nombre y domicilio de los testigos', '☐'],
    ['VII. Datos relativos a la actuación (incluir el inventario de lo entregado)', '☐'],
    ['VIII. Declaración del visitado, si quiere hacerla', '☐'],
    ['IX. Nombre y firma de quienes intervinieron', '☐'],
  ], [W - 1400, 1400]));
  C.push(SP(80));
  [
    '**Declaración en el acta.** Usar el espacio para dejar constancia de: documentos que se entregarán después, irregularidades de la orden, y la reserva del derecho a formular observaciones y ofrecer pruebas por escrito. Fórmula sugerida: «Me reservo el derecho de formular observaciones y ofrecer pruebas por escrito dentro del plazo de cinco días previsto en el artículo 68 de la Ley Federal de Procedimiento Administrativo».',
    '**Firma.** Se recomienda firmar, de ser posible con la leyenda «Firmo de recibido, sin que ello implique conformidad con su contenido». La negativa a firmar no invalida el acta (LFPA art. 67, fracc. IX) y sólo pierde la oportunidad de dejar constancia propia.',
    '**Copia.** Obtener copia íntegra del acta y de sus anexos.',
  ].forEach(t => C.push(B(t)));

  C.push(H2('Calendario posterior a la visita'));
  C.push(T(['Momento', 'Acción', 'Responsable', 'Fundamento'], [
    ['Día 0 (mismo día)', 'Reunión de evaluación; enviar acta y orden al asesor y a la defensa; abrir expediente de la visita.', 'Responsable de atención', '—'],
    ['Dentro de 5 días', 'Presentar escrito de observaciones y pruebas sobre los hechos del acta (Anexo D).', 'Asesor / defensa', 'LFPA art. 68'],
    ['Dentro de los mismos 5 días', 'Decidir si procede el **reconocimiento expreso** del art. 55: sólo aplica a faltas regularizadas espontáneamente antes del inicio de la visita. Se presenta escrito libre con los requisitos del Reglamento.', 'Defensa', 'LFPIORPI art. 55, 1er párr.; Reglamento art. 55 Bis'],
    ['Al notificarse el inicio del procedimiento sancionador', '15 días para exponer lo que convenga y aportar pruebas. Si ya se usó el beneficio anterior, es el plazo para reconocer y acceder a la reducción de hasta 50%.', 'Defensa', 'LFPA art. 72; LFPIORPI art. 55, 2º párr.'],
    ['Al notificarse la multa', 'Recurso de revisión ante la propia Secretaría, o juicio contencioso administrativo ante el TFJA.', 'Defensa', 'LFPIORPI art. 61'],
  ], [1900, 3700, 1500, W - 7100]));
  C.push(SP(80));
  C.push(BOX('Presentación de escritos', [
    'Los escritos dirigidos al SAT se presentan en sus oficinas autorizadas; cuando el domicilio está fuera de la entidad sede de esas oficinas, pueden enviarse por Correos de México, correo certificado con acuse de recibo (RCG art. 36 Bis 1). Los plazos se cuentan en días hábiles conforme al art. 36 Bis de las RCG. **Confirmar cada cómputo con la defensa jurídica.**',
  ], 'navy'));

  // IX
  C.push(H1('IX. Requerimientos de información sin visita'));
  C.push(P('Desde la reforma de 2025, la autoridad puede verificar sin presentarse en el domicilio, mediante requerimientos por escrito o notificados en el Portal. El procedimiento reglamentario es **más rápido y con menos oportunidades de defensa** que la visita, por lo que los plazos deben vigilarse con especial cuidado.'));
  C.push(T(['Etapa', 'Plazo', 'Consecuencia / nota', 'Fundamento'], [
    ['Notificación electrónica', '3 días hábiles para abrir', 'Si no se abre, se tiene por notificado al 4º día hábil y el plazo corre igual.', 'RCG art. 6'],
    ['Atención del requerimiento', '10 días hábiles desde el día siguiente a su recepción', 'Si no se atiende o se atiende incompleto, el SAT puede imponer la sanción en un plazo no mayor a 10 días hábiles **sin seguir el procedimiento sancionador de la LFPA**.', 'Reglamento art. 8'],
    ['Prórroga', 'Hasta 5 días hábiles', 'Sólo si se solicita **dentro** de los 10 días originales (Anexo E).', 'Reglamento art. 8, 2º párr.'],
    ['Oficio de observaciones', '5 días hábiles para desvirtuar', 'Los hechos no desvirtuados se tienen por consentidos; la resolución se emite en un plazo no mayor a 20 días hábiles.', 'Reglamento art. 9'],
    ['Información de otras autoridades', '10 días hábiles para manifestarse', 'Cuando el SAT dé a conocer información, expedientes o documentos de terceros.', 'Reglamento art. 10 Bis'],
    ['Requerimiento sobre alta y registro', '10 días hábiles (+5 a solicitud)', 'Para corroborar la información del padrón.', 'RCG art. 9'],
  ], [1900, 2100, 3600, W - 7600]));
  C.push(SP(80));
  C.push(BOX('Punto de atención para la defensa', [
    'Los arts. 8 y 9 del Reglamento permiten sancionar sin el procedimiento de la LFPA, mientras que el art. 35 de la Ley prevé la aplicación supletoria de esa ley. La compatibilidad entre ambas disposiciones es un aspecto que la defensa jurídica debe valorar en cada caso. **Para efectos operativos, este protocolo asume que los plazos reglamentarios se aplicarán en sus términos.**',
  ], 'red'));
  C.push(SP(80));
  C.push(H2('Pasos ante un requerimiento'));
  [
    'Registrar en la Bitácora la fecha de notificación (o la fecha en que se tiene por notificado) y calcular el vencimiento.',
    'Enviar copia al asesor de cumplimiento y a la defensa el mismo día.',
    'Delimitar lo solicitado contra el criterio de «directamente relacionado» (art. 34) y el periodo de cinco años (art. 36).',
    'Integrar la respuesta con documentos legibles y sin modificaciones (art. 62). Si el volumen lo exige, solicitar la prórroga antes del día 10.',
    'Presentar la respuesta y conservar el acuse junto con una copia exacta de lo entregado.',
  ].forEach(t => C.push(N(t, 'num3')));

  // X
  C.push(H1('X. Conductas prohibidas'));
  C.push(BOX(null, [
    '✕  Negar o condicionar el acceso a la autoridad, o pedir que «regrese otro día».',
    '✕  Destruir, ocultar, modificar o «completar» documentos antes o durante la visita.',
    '✕  Entregar copias ilegibles o incompletas a sabiendas.',
    '✕  Entregar información no solicitada o ajena a la Actividad Vulnerable.',
    '✕  Reconocer verbalmente incumplimientos durante la diligencia.',
    '✕  Firmar documentos en blanco o actas que no se han leído.',
    '✕  Dejar vencer el plazo de 5 días sin presentar observaciones ni decidir sobre el art. 55.',
    '✕  Dejar de revisar el Portal: las notificaciones surten efectos aunque no se abran.',
  ], 'red'));

  // XI
  C.push(H1('XI. Matriz de la Carpeta de Visita'));
  C.push(P('La Carpeta de Visita reúne, por periodo mensual y para los últimos cinco años, la evidencia que la autoridad suele solicitar. La columna «Exigible desde» refleja los plazos de entrada en vigor de las obligaciones incorporadas por la reforma de 2025 y el Acuerdo 115/2026.'));
  const mat = [
    ['Acuse de alta y registro y sus actualizaciones', `D3${s}`, 'Art. 18, fracc. IV Bis; RCG arts. 4–7', 'Vigente'],
    [v.d2[0], v.d2[1], v.d2[2], 'Vigente'],
    ['Manual de Políticas Internas', `D5${s}`, 'Art. 18, fracc. VIII; RCG arts. 37 y 37 Bis', 'Vigente; con metodología EBR: 01-03-2027'],
    ['Evaluación con enfoque basado en riesgos', `D6${s}`, 'Art. 18, fracc. VII; RCG Cap. II Quáter', '01-03-2027'],
    ['Expedientes de identificación de Clientes o Usuarias', `D7${s}`, 'Art. 18, fracc. I y II', 'Vigente'],
    ['Identificación del Beneficiario Controlador', v.bc, 'Art. 18, fracc. III', 'Vigente'],
    ['Identificación y seguimiento de Personas Políticamente Expuestas', `D8${s}`, 'Art. 18, fracc. VIII', 'Vigente'],
    ['Avisos e Informes con acuses y documentación soporte', `D12${s}`, 'Arts. 18, fracc. VI, y 25; Reglamento art. 20', 'Vigente'],
  ];
  if (key === 'notarial') mat.push(['Acuses de avisos del inciso a) presentados por medios fiscales', 'D12-NOT', 'Art. 24, último párr.; Reglamento art. 17', 'Vigente']);
  if (key === 'notarial' || key === 'corredores') mat.push(['Instrumentos con identificación de la forma de pago (8,025 UMA o más)', 'Protocolo / archivo', 'Art. 33; Reglamento art. 45', 'Vigente']);
  if (key === 'inmobiliaria') mat.push(['Evidencia de forma de pago por operación (restricción de efectivo)', 'Archivo operativo / D7', 'Art. 32', 'Vigente']);
  mat.push(
    ['Avisos de 24 horas y análisis de operaciones inusuales', `D9${s}`, 'Art. 18, fracc. VI, 2º párr.; Reglamento art. 7 Bis', 'Vigente (sujeto a formatos)'],
    ['Control de acumulación de operaciones en seis meses', `D12${s}`, 'Art. 17, penúltimo párr.; Reglamento art. 7', 'Vigente'],
    [key.startsWith('prof') ? 'Expedientes de asuntos: preparación vs. representación' : 'Contratos, comprobantes de pago y correspondencia de cada operación', `Archivo operativo / D10${s}`, 'Art. 18, fracc. IV', 'Vigente'],
    ['Política y bitácora de resguardo documental (10 años)', `D10${s}`, 'Art. 18, fracc. IV; Reglamento art. 20', 'Vigente'],
    ['Programa anual de capacitación y constancias', `D11${s} / D11B${s} / D11C${s}`, 'Art. 18, fracc. IX; RCG art. 39 Bis', 'Periodo 2027'],
    ['Expediente de selección de personal', v.selPers, 'Art. 18, fracc. IX; RCG art. 39 Bis 2', 'Contrataciones desde 01-03-2027'],
    ['Mecanismos automatizados de monitoreo', '—', 'Art. 18, fracc. X; RCG Cap. XIII', '01-06-2027'],
    ['Dictamen de auditoría y regularización de observaciones', `D13${s}`, 'Art. 18, fracc. XI; Reglamento art. 12 Bis', 'Periodo 2028'],
    ['Evidencia de consulta diaria del Portal', 'Bitácora (sección XII)', 'RCG art. 6', '30-11-2026'],
  );
  C.push(T(['Evidencia', 'Documento del Programa', 'Fundamento', 'Exigible desde'], mat, [3300, 1700, 2560, W - 7560]));

  // XII
  C.push(H1('XII. Bitácora de visitas, requerimientos y consultas al Portal'));
  C.push(T(['Fecha', 'Tipo (visita / requerimiento / notificación)', 'Autoridad y número de oficio', 'Vencimiento', 'Acciones y acuse', 'Responsable'],
    Array.from({ length: 5 }, () => ['', '', '', '', '', '']), [1200, 1900, 1900, 1300, 1860, 1200]));

  // XIII
  C.push(H1('XIII. Particularidades de este tipo de sujeto obligado'));
  PART[key].forEach(t => C.push(B(t)));

  // Aprobación
  C.push(SP(200));
  C.push(KV([['Aprobado por:', ''], ['Cargo:', ''], ['Fecha:', ''], ['Firma:', '']]));

  // ANEXO A
  C.push(BR());
  C.push(H1('Anexo A. Ficha de Recepción — para el personal de primer contacto'));
  C.push(P('Imprimir y tener a la vista en recepción.', { italics: true }));
  const llamada = v.pf
    ? 'Llame de inmediato al titular ([Teléfono del titular]) y, si no responde, al encargado operativo ([Teléfono del encargado operativo]).'
    : 'Llame de inmediato, en este orden: responsable ([Teléfono del responsable]) → suplente 1 ([Teléfono del suplente 1]) → suplente 2 ([Teléfono del suplente 2]).';
  C.push(BOX('Si llega personal del SAT o de Hacienda para una visita de verificación:', [
    N('Salude con cortesía y ofrezca un lugar para esperar. **No niegue el acceso.**', 'num4'),
    N('Pida amablemente su **credencial con fotografía** y la **orden de visita por escrito**.', 'num4'),
    N(llamada, 'num4'),
    N('Avise también a CERTUS CONSULTORES (999 163 1363) y a ESPINOSA-SALAZAR ABOGADOS CORPORATIVOS (999 575 3129).', 'num4'),
    N('Si nadie del equipo responsable está disponible, **la visita continúa con usted**: es normal y no es una falta suya.', 'num4'),
    N('No entregue documentos ni responda preguntas sobre clientes u operaciones: indique que la persona responsable las atenderá.', 'num4'),
    N('No firme nada sin haberlo leído. Si debe firmar, escriba: «Firmo de recibido, sin que implique conformidad».', 'num4'),
    N('Pida copia de todo lo que le entreguen o firme.', 'num4'),
  ]));

  // ANEXO B
  C.push(H1('Anexo B. Formato de inventario de documentación entregada'));
  C.push(KV([['Oficio / orden núm.:', ''], ['Fecha de la diligencia:', ''], ['Persona verificadora:', ''], ['Entregó:', '']]));
  C.push(SP(60));
  C.push(T(['#', 'Documento', 'Periodo', 'Original / copia', 'Fojas', 'Devuelto (sí/no)'],
    [1, 2, 3, 4, 5].map(n => [String(n), '', '', '', '', '']), [500, 3760, 1400, 1400, 900, 1400]));
  C.push(SP(100));
  C.push(KV([['Firma de quien recibe por la autoridad:', ''], ['Firma de quien entrega:', '']], [4200, W - 4200]));

  // ANEXO C (escrito de observaciones)
  C.push(BR());
  C.push(H1('Anexo C. Modelo de escrito de observaciones al acta de verificación'));
  C.push(P('[Ciudad], a [fecha del escrito].', { align: AlignmentType.RIGHT }));
  C.push(P('[Unidad administrativa del SAT que emitió la orden]\nPresente.', { after: 60, align: AlignmentType.LEFT }));
  C.push(P(`${v.firmante}, ${v.firmanteCalidad}, señalando como domicilio para oír y recibir notificaciones [domicilio para notificaciones] y autorizando para tales efectos a [nombres de autorizados], con fundamento en los artículos 35 de la Ley Federal para la Prevención e Identificación de Operaciones con Recursos de Procedencia Ilícita y 68 de la Ley Federal de Procedimiento Administrativo, dentro del plazo de cinco días siguientes al levantamiento del acta de verificación de fecha [fecha del acta], derivada de la orden [número de orden], comparezco a formular las siguientes:`));
  C.push(P('**OBSERVACIONES**', { align: AlignmentType.CENTER }));
  C.push(P('PRIMERA. [Hecho asentado en el acta y precisión o aclaración que corresponde, con referencia al documento que lo acredita.]'));
  C.push(P('SEGUNDA. [Documentación que quedó pendiente de entregar y que se exhibe en este acto.]'));
  C.push(P('TERCERA. [En su caso, irregularidades de la orden o del acta que se hacen valer.]'));
  C.push(P('**PRUEBAS**', { align: AlignmentType.CENTER }));
  C.push(P('1. Documental: [descripción], que se relaciona con la observación [número de observación].'));
  C.push(P('Por lo expuesto, solicito se tengan por formuladas las presentes observaciones y por ofrecidas las pruebas señaladas, y se valoren al emitir la resolución que corresponda.'));
  C.push(P('Atentamente,\n[Firma]', { align: AlignmentType.CENTER, before: 200 }));
  C.push(BOX(null, ['Si se opta por el reconocimiento expreso del art. 55, éste se presenta en escrito libre por separado, con los requisitos del art. 55 Bis del Reglamento (totalidad de las faltas, operación y periodo, manifestación bajo protesta de que fueron corregidas y documentación que lo acredite). **Redactarlo siempre con la defensa jurídica.**'], 'navy'));

  // ANEXO D (prórroga)
  C.push(H1('Anexo D. Modelo de solicitud de prórroga de requerimiento'));
  C.push(P('[Ciudad], a [fecha del escrito].', { align: AlignmentType.RIGHT }));
  C.push(P('[Unidad administrativa del SAT]\nPresente.', { after: 60, align: AlignmentType.LEFT }));
  C.push(P(`${v.firmante}, ${v.firmanteCalidad}, con relación al requerimiento contenido en el oficio [número de oficio], notificado el [fecha de notificación], con fundamento en el artículo 8, segundo párrafo, del Reglamento de la Ley Federal para la Prevención e Identificación de Operaciones con Recursos de Procedencia Ilícita, y estando dentro del plazo original de diez días hábiles, solicito se prorrogue hasta por cinco días hábiles el plazo para la entrega de la información requerida, en atención a [volumen de la documentación / necesidad de recabar información en archivo / otra causa concreta].`));
  C.push(P('Atentamente,\n[Firma]', { align: AlignmentType.CENTER, before: 200 }));

  const footer = new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [
    new TextRun({ text: 'CERTUS CONSULTORES  |  Certus PLD  |  Mérida, Yucatán  |  999 163 1363     ·     Pág. ', size: 16, color: '7A869A' }),
    new TextRun({ children: [PageNumber.CURRENT], size: 16, color: '7A869A' }),
  ] })] });
  return new Document({ styles: L.styles, numbering: L.numbering, sections: [{ properties: { page: L.page }, footers: { default: footer }, children: C }] });
}

const [key, out] = process.argv.slice(2);
if (!V[key]) { console.error('variante inválida'); process.exit(1); }
Packer.toBuffer(build(key)).then(b => { fs.writeFileSync(out, b); console.log('ok', key); });
