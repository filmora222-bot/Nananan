const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType,
  LevelFormat, PageBreak
} = require('docx');
const fs = require('fs');

// ── helpers ──────────────────────────────────────────────────────────────────
const bdr  = c => ({ style: BorderStyle.SINGLE, size: 2, color: c });
const borders    = { top: bdr("AAAAAA"), bottom: bdr("AAAAAA"), left: bdr("AAAAAA"), right: bdr("AAAAAA") };
const noBorder   = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders  = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };
const hdBorders  = { top: bdr("1F3864"), bottom: bdr("1F3864"), left: bdr("1F3864"), right: bdr("1F3864") };

const sp = (b, a) => ({ spacing: { before: b, after: a } });

function pb() { return new Paragraph({ children: [new PageBreak()] }); }

function el(lines = 1) {
  return Array.from({ length: lines }, () =>
    new Paragraph({ children: [new TextRun({ text: "", size: 22, font: "Arial" })], spacing: { before: 40, after: 40 } })
  );
}

function sectionBanner(num, title) {
  return new Paragraph({
    children: [new TextRun({ text: `  ${num}  ${title}`, bold: true, size: 24, font: "Arial", color: "FFFFFF" })],
    shading: { fill: "1F3864", type: ShadingType.CLEAR },
    spacing: { before: 0, after: 0 }
  });
}

function subBanner(text) {
  return new Paragraph({
    children: [new TextRun({ text: `  ${text}`, bold: true, size: 22, font: "Arial", color: "FFFFFF" })],
    shading: { fill: "2E5797", type: ShadingType.CLEAR },
    spacing: { before: 200, after: 0 }
  });
}

function titulo(text) {
  return new Paragraph({
    children: [new TextRun({ text, bold: true, size: 28, font: "Arial", color: "1F3864" })],
    alignment: AlignmentType.CENTER,
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: "1F3864", space: 4 } },
    spacing: { before: 300, after: 200 }
  });
}

function h2(text) {
  return new Paragraph({
    children: [new TextRun({ text, bold: true, size: 24, font: "Arial", color: "1F3864" })],
    spacing: { before: 260, after: 120 }
  });
}

function h3(text) {
  return new Paragraph({
    children: [new TextRun({ text, bold: true, size: 22, font: "Arial", color: "2E5797" })],
    spacing: { before: 180, after: 80 }
  });
}

function body(text, opts = {}) {
  return new Paragraph({
    children: [new TextRun({ text, size: 22, font: "Arial", ...opts })],
    alignment: opts.center ? AlignmentType.CENTER : AlignmentType.JUSTIFIED,
    spacing: { before: 80, after: 80 }
  });
}

function bullet(text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    children: [new TextRun({ text, size: 22, font: "Arial" })],
    spacing: { before: 60, after: 60 }
  });
}

function num(text) {
  return new Paragraph({
    numbering: { reference: "numbers", level: 0 },
    children: [new TextRun({ text, size: 22, font: "Arial" })],
    spacing: { before: 60, after: 60 }
  });
}

function fichaRow(campo, valor, fill = "EAF0F8") {
  return new TableRow({
    children: [
      new TableCell({
        borders,
        width: { size: 3200, type: WidthType.DXA },
        shading: { fill, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 140, right: 120 },
        children: [new Paragraph({ children: [new TextRun({ text: campo, bold: true, size: 20, font: "Arial" })] })]
      }),
      new TableCell({
        borders,
        width: { size: 6160, type: WidthType.DXA },
        margins: { top: 80, bottom: 80, left: 140, right: 120 },
        children: [new Paragraph({ children: [new TextRun({ text: valor, size: 20, font: "Arial" })] })]
      }),
    ]
  });
}

function fichaTable(rows) {
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [3200, 6160],
    rows
  });
}

function sigLine(left, right) {
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [4320, 720, 4320],
    rows: [
      new TableRow({ children: [
        new TableCell({ borders: noBorders, width: { size: 4320, type: WidthType.DXA },
          children: [new Paragraph({ border: { top: { style: BorderStyle.SINGLE, size: 4, color: "333333" } },
            children: [new TextRun({ text: "", size: 22 })] })] }),
        new TableCell({ borders: noBorders, width: { size: 720, type: WidthType.DXA }, children: [new Paragraph({ children: [new TextRun("")] })] }),
        new TableCell({ borders: noBorders, width: { size: 4320, type: WidthType.DXA },
          children: [new Paragraph({ border: { top: { style: BorderStyle.SINGLE, size: 4, color: "333333" } },
            children: [new TextRun({ text: "", size: 22 })] })] }),
      ]})
    ]
  });
}

function sigLabel(left, right) {
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [4320, 720, 4320],
    rows: [
      new TableRow({ children: [
        new TableCell({ borders: noBorders, width: { size: 4320, type: WidthType.DXA },
          children: [new Paragraph({ alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: left, size: 20, font: "Arial", italics: true })] })] }),
        new TableCell({ borders: noBorders, width: { size: 720, type: WidthType.DXA }, children: [new Paragraph({ children: [new TextRun("")] })] }),
        new TableCell({ borders: noBorders, width: { size: 4320, type: WidthType.DXA },
          children: [new Paragraph({ alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: right, size: 20, font: "Arial", italics: true })] })] }),
      ]})
    ]
  });
}

function citeBox(text) {
  return new Paragraph({
    children: [new TextRun({ text, size: 19, font: "Arial", italics: true, color: "555555" })],
    border: {
      top: { style: BorderStyle.SINGLE, size: 4, color: "2E5797", space: 3 },
      bottom: { style: BorderStyle.SINGLE, size: 4, color: "2E5797", space: 3 },
      left: { style: BorderStyle.THICK, size: 12, color: "2E5797", space: 6 },
    },
    indent: { left: 360 },
    spacing: { before: 120, after: 120 }
  });
}

function pregunta(text) {
  return new Paragraph({
    children: [new TextRun({ text, size: 22, font: "Arial", bold: true, color: "2E5797" })],
    spacing: { before: 140, after: 60 }
  });
}

function respuesta(text) {
  return new Paragraph({
    children: [new TextRun({ text, size: 22, font: "Arial" })],
    alignment: AlignmentType.JUSTIFIED,
    indent: { left: 360 },
    spacing: { before: 40, after: 100 }
  });
}

function notaFoto() {
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [9360],
    rows: [
      new TableRow({ children: [
        new TableCell({
          borders: hdBorders,
          width: { size: 9360, type: WidthType.DXA },
          shading: { fill: "F0F4FA", type: ShadingType.CLEAR },
          margins: { top: 120, bottom: 120, left: 200, right: 200 },
          children: [
            new Paragraph({ alignment: AlignmentType.CENTER,
              children: [new TextRun({ text: "[ESPACIO RESERVADO PARA FOTOGRAFÍA]", bold: true, size: 22, font: "Arial", color: "2E5797" })] }),
            new Paragraph({ alignment: AlignmentType.CENTER,
              children: [new TextRun({ text: "Profesional de frente — Paciente de espaldas (rostro no visible)", size: 20, font: "Arial", italics: true, color: "666666" })] }),
            new Paragraph({ alignment: AlignmentType.CENTER,
              children: [new TextRun({ text: "Vestimenta formal / bata. Evidencia de interacción clínica.", size: 20, font: "Arial", italics: true, color: "666666" })] }),
          ]
        })
      ]})
    ]
  });
}

// ── DOCUMENTO ────────────────────────────────────────────────────────────────
const doc = new Document({
  styles: { default: { document: { run: { font: "Arial", size: 22 } } } },
  numbering: {
    config: [
      { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "numbers", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    ]
  },
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1280, right: 1280, bottom: 1280, left: 1280 } } },
    children: [

      // ══════════════════════════════════════════════════════════════════════
      //  ENCABEZADO INSTITUCIONAL
      // ══════════════════════════════════════════════════════════════════════
      new Paragraph({
        children: [new TextRun({ text: "BENEMÉRITA UNIVERSIDAD AUTÓNOMA DE PUEBLA", bold: true, size: 24, font: "Arial", color: "1F3864" })],
        alignment: AlignmentType.CENTER, spacing: { before: 0, after: 40 }
      }),
      new Paragraph({
        children: [new TextRun({ text: "FACULTAD DE MEDICINA  ·  LICENCIATURA EN MEDICINA", bold: true, size: 22, font: "Arial", color: "1F3864" })],
        alignment: AlignmentType.CENTER, spacing: { before: 0, after: 40 }
      }),
      new Paragraph({
        children: [new TextRun({ text: "TERCER EXAMEN PRÁCTICO DE TANATOLOGÍA  ·  CASO CLÍNICO 3", size: 21, font: "Arial", color: "444444" })],
        alignment: AlignmentType.CENTER, spacing: { before: 0, after: 40 }
      }),
      new Paragraph({
        children: [new TextRun({ text: "Profesor y autor: DC. Rosario Robles Galindo  ·  14 de abril de 2026", size: 20, font: "Arial", italics: true, color: "555555" })],
        alignment: AlignmentType.CENTER, spacing: { before: 0, after: 40 }
      }),
      new Paragraph({
        children: [new TextRun({ text: "González Félix Jorge Omar  ·  Matrícula: 202327266", size: 21, font: "Arial", bold: true, color: "2E5797" })],
        alignment: AlignmentType.CENTER, spacing: { before: 40, after: 40 }
      }),
      new Paragraph({
        border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "1F3864", space: 2 } },
        children: [new TextRun({ text: "", size: 4 })], spacing: { before: 80, after: 200 }
      }),

      sectionBanner("▌", "EXPEDIENTE CLÍNICO ETAPA 1 — ENTREVISTA CLÍNICA"),
      ...el(1),

      // ══════════════════════════════════════════════════════════════════════
      //  PUNTO 1: CONSENTIMIENTO INFORMADO
      // ══════════════════════════════════════════════════════════════════════
      subBanner("PUNTO 1 · CONSENTIMIENTO INFORMADO"),
      titulo("CONSENTIMIENTO INFORMADO PARA ATENCIÓN TANATOLÓGICA"),

      body("El presente documento tiene como propósito garantizar que la paciente y/o su representante legal cuenten con información suficiente, clara y comprensible acerca de la naturaleza, los objetivos y el alcance del proceso de acompañamiento tanatológico, a fin de que su participación sea libre, voluntaria e informada.", { italics: true }),
      ...el(1),

      h2("I. INFORMACIÓN SOBRE EL PROCESO DE ACOMPAÑAMIENTO"),
      body("La tanatología es una disciplina que ofrece acompañamiento emocional, espiritual y psicológico a personas que atraviesan procesos de duelo, enfermedad terminal, pérdidas significativas o situaciones de fin de vida. El objetivo primordial es brindar apoyo integral que favorezca la calidad de vida, la dignidad y el bienestar de la paciente y de sus redes de apoyo."),
      ...el(1),
      body("El proceso de acompañamiento incluye, entre otras actividades:"),
      bullet("Entrevistas individuales de exploración y seguimiento."),
      bullet("Aplicación de técnicas de contención emocional y manejo del duelo."),
      bullet("Orientación a la familia y red de apoyo cercana."),
      bullet("Derivación a otros especialistas cuando sea necesario (psicología clínica, trabajo social, cuidados paliativos)."),
      ...el(1),

      h2("II. DERECHOS DE LA PACIENTE"),
      body("La paciente tiene derecho a:"),
      num("Recibir información clara y veraz sobre su proceso de atención."),
      num("Aceptar o rechazar cualquier intervención propuesta."),
      num("Retirar su consentimiento en cualquier momento, sin consecuencia alguna para su atención."),
      num("Mantener la confidencialidad de su información personal y clínica, conforme a la normativa vigente de protección de datos."),
      num("Ser tratada con respeto, dignidad y sin discriminación de ningún tipo."),
      ...el(1),

      h2("III. CONFIDENCIALIDAD"),
      body("Toda la información compartida durante las sesiones de acompañamiento tanatológico será tratada con estricta confidencialidad. La información solo podrá ser revelada cuando exista riesgo inminente para la vida de la paciente o de terceros, o cuando sea requerida por orden judicial, notificando a la paciente en la medida en que sea posible."),
      ...el(1),

      h2("IV. DECLARACIÓN DE CONSENTIMIENTO"),
      body("Yo, la paciente o representante legal, declaro que:"),
      bullet("He leído y comprendido la información contenida en este documento."),
      bullet("He tenido la oportunidad de realizar preguntas, las cuales han sido respondidas de manera satisfactoria."),
      bullet("Doy mi consentimiento libre y voluntario para participar en el proceso de acompañamiento tanatológico."),
      ...el(1),

      new Paragraph({
        children: [new TextRun({ text: "Lugar y fecha: Tlaxcala, Tlax., a   16   de   Abril   de 2026.", size: 22, font: "Arial" })],
        spacing: { before: 120, after: 200 }
      }),
      ...el(1),
      sigLine("Paciente", "Especialista"),
      new Paragraph({ spacing: { before: 40, after: 0 }, children: [new TextRun("")] }),
      sigLabel("Firma del Paciente / Representante Legal", "Firma del Especialista de Primer Contacto"),
      ...el(1),
      new Table({
        width: { size: 9360, type: WidthType.DXA }, columnWidths: [4320, 720, 4320],
        rows: [new TableRow({ children: [
          new TableCell({ borders: noBorders, width: { size: 4320, type: WidthType.DXA }, children: [
            new Paragraph({ alignment: AlignmentType.CENTER,
              children: [new TextRun({ text: "Nombre:   Omitido", size: 20, font: "Arial" })] })] }),
          new TableCell({ borders: noBorders, width: { size: 720, type: WidthType.DXA }, children: [new Paragraph({ children: [new TextRun("")] })] }),
          new TableCell({ borders: noBorders, width: { size: 4320, type: WidthType.DXA }, children: [
            new Paragraph({ alignment: AlignmentType.CENTER,
              children: [new TextRun({ text: "Nombre:  González Félix Jorge Omar", size: 20, font: "Arial" })] })] }),
        ]})]
      }),
      new Paragraph({ spacing: { before: 20, after: 0 }, children: [new TextRun("")] }),
      new Table({
        width: { size: 9360, type: WidthType.DXA }, columnWidths: [4320, 720, 4320],
        rows: [new TableRow({ children: [
          new TableCell({ borders: noBorders, width: { size: 4320, type: WidthType.DXA }, children: [
            new Paragraph({ alignment: AlignmentType.CENTER,
              children: [new TextRun({ text: "Testigo: _______________________", size: 20, font: "Arial" })] })] }),
          new TableCell({ borders: noBorders, width: { size: 720, type: WidthType.DXA }, children: [new Paragraph({ children: [new TextRun("")] })] }),
          new TableCell({ borders: noBorders, width: { size: 4320, type: WidthType.DXA }, children: [
            new Paragraph({ alignment: AlignmentType.CENTER,
              children: [new TextRun({ text: "Cédula Profesional:  202327266", size: 20, font: "Arial" })] })] }),
        ]})]
      }),
      ...el(1),
      citeBox("Rosario Robles Galindo. Materiales Didácticos de Tanatología. Caso 3. Expediente y Entrevista Clínica (Consentimiento Informado y Fichas de Identificación). Periodo 2026. Benemérita Universidad Autónoma de Puebla, Facultad de Medicina, Licenciatura en Medicina."),

      // ══════════════════════════════════════════════════════════════════════
      //  PUNTO 2: FICHA DE IDENTIFICACIÓN
      // ══════════════════════════════════════════════════════════════════════
      pb(),
      subBanner("PUNTO 2 · FICHA DE IDENTIFICACIÓN (FDI)"),
      titulo("FICHA DE IDENTIFICACIÓN DEL PACIENTE — CASO CLÍNICO 3"),

      body("Los datos personales del presente expediente se encuentran protegidos conforme a la normativa vigente. El campo correspondiente al nombre de la paciente se omite intencionalmente para salvaguardar su identidad.", { italics: true }),
      ...el(1),

      h2("A. DATOS GENERALES"),
      fichaTable([
        fichaRow("Folio de Expediente", "TAN-CAS03-2026"),
        fichaRow("Nombre", "Omitido / Confidencial"),
        fichaRow("Fecha de la consulta", "16 de abril de 2026"),
        fichaRow("Sexo", "Femenino"),
        fichaRow("Edad", "43 años"),
        fichaRow("Domicilio", "Calle Reforma #120, Col. El Sabinal"),
        fichaRow("Lugar de origen", "Pachuca, Hidalgo"),
        fichaRow("Vive con su familia", "Sí"),
        fichaRow("Estado civil", "Viuda"),
        fichaRow("Nivel socioeconómico", "Medio"),
        fichaRow("Escolaridad", "Maestría"),
        fichaRow("Religión / Ideología cultural", "Católica"),
        fichaRow("Ocupación", "Docente"),
      ]),
      ...el(1),

      h2("B. DATOS DEL DUELO"),
      fichaTable([
        fichaRow("Origen del duelo", "Pérdida de familiares"),
        fichaRow("Duración del duelo", "Aproximadamente 3 años"),
        fichaRow("Tipo de duelo", "Duelo circunstancial / biológico por pérdidas múltiples"),
        fichaRow("Estado actual del duelo", "Proceso activo con manifestaciones emocionales persistentes"),
      ]),
      ...el(1),

      h2("C. DATOS DEL ESPECIALISTA RESPONSABLE"),
      fichaTable([
        fichaRow("Nombre del especialista", "González Félix Jorge Omar"),
        fichaRow("Matrícula / Cédula profesional", "202327266"),
        fichaRow("Área de formación", "Medicina — Tanatología clínica"),
        fichaRow("Institución de adscripción", "Facultad de Medicina, BUAP"),
        fichaRow("Fecha de apertura del expediente", "16 de abril de 2026"),
      ]),
      ...el(1),
      citeBox("Rosario Robles Galindo. Materiales Didácticos de Tanatología. Caso 3. Expediente y Entrevista Clínica (Consentimiento Informado y Fichas de Identificación). Periodo 2026. Benemérita Universidad Autónoma de Puebla, Facultad de Medicina, Licenciatura en Medicina."),

      // ══════════════════════════════════════════════════════════════════════
      //  PUNTO 2.1: EVIDENCIA FOTOGRÁFICA + DESCRIPCIÓN DEL ENTORNO
      // ══════════════════════════════════════════════════════════════════════
      pb(),
      subBanner("PUNTO 2.1 · EVIDENCIA DEL CASO CLÍNICO 3 — DESCRIPCIÓN DEL ENTORNO DE ATENCIÓN"),
      titulo("DESCRIPCIÓN DEL ESPACIO FÍSICO DE ATENCIÓN PRESENCIAL"),

      ...el(1),
      notaFoto(),
      ...el(1),

      h2("I. UBICACIÓN DEL ESPACIO DE ATENCIÓN"),
      fichaTable([
        fichaRow("Nombre del espacio", "Consultorio de Atención Clínica Tanatológica — Sede Práctica"),
        fichaRow("Tipo de espacio", "Consultorio privado dentro de instalaciones universitarias / clínica de prácticas"),
        fichaRow("Dirección", "Av. San Claudio s/n, Ciudad Universitaria, Puebla, Pue. C.P. 72570"),
        fichaRow("Referencia", "Edificio de Ciencias de la Salud, planta baja, ala norte. A un costado del Departamento de Medicina Familiar."),
        fichaRow("Fecha y hora de la sesión", "16 de abril de 2026, 10:00 hrs."),
      ]),
      ...el(1),

      h2("II. DESCRIPCIÓN FÍSICA DETALLADA DEL CONSULTORIO"),
      body("El consultorio destinado a la atención tanatológica presencial es un espacio de aproximadamente 16 m², ubicado en la planta baja del edificio, alejado de zonas de tránsito intenso para garantizar tranquilidad y privacidad. El acceso se realiza a través de un pasillo interior con iluminación tenue, lo que permite a la paciente una transición gradual desde el entorno exterior hacia el espacio terapéutico."),
      ...el(1),
      body("Las paredes están pintadas en tonos neutros cálidos: beige arena y blanco roto, lo que genera una atmósfera de contención y calma. El piso es de madera laminada color miel, que contribuye a la sensación de calidez y confort. No existen objetos decorativos recargados; los elementos visuales se reducen a dos cuadros con motivos naturalistas abstractos (bosque y agua en movimiento), seleccionados para favorecer la evocación de calma y continuidad de la vida."),
      ...el(1),

      h3("Mobiliario y distribución espacial"),
      body("El mobiliario principal consta de dos sillones tapizados en tela gris marengo, de respaldo medio y brazos suaves, posicionados en un ángulo de 90° entre sí —no frente a frente—, conforme a las recomendaciones de psicología ambiental para entornos de acompañamiento en duelo. Este ángulo facilita la comunicación sin generar confrontación visual directa, reduciendo la percepción de evaluación por parte de la paciente."),
      ...el(1),
      body("Entre ambos sillones se ubica una pequeña mesa auxiliar de madera donde descansa una caja de pañuelos desechables, un vaso con agua para la paciente y una planta pequeña de interior (suculenta), que aporta un elemento de vida y naturaleza al espacio. Frente al sillón del especialista hay un escritorio lateral compacto con una silla ergonómica, utilizado únicamente para el registro de notas clínicas, de manera que no interfiera visualmente con la interacción durante la sesión."),

      h3("Iluminación"),
      body("La iluminación es predominantemente cálida. Cuenta con un ventanal lateral con vista a un patio interior ajardinado que permite el paso de luz natural indirecta, la cual puede ser regulada mediante persianas de láminas en color crema. Se complementa con una lámpara de pie de luz amarilla tenue ubicada en la esquina posterior del consultorio, evitando focos de luz directa que pudieran generar incomodidad o tensión visual. Durante la sesión del 16 de abril, la iluminación natural fue suficiente y se mantuvo sin necesidad de luz artificial."),

      h3("Condiciones acústicas y térmicas"),
      body("Las paredes del consultorio cuentan con panel de absorción acústica básico. En el pasillo exterior se dispone de una fuente de ruido blanco (ventilador de baja potencia) que difumina los sonidos del corredor, protegiendo la confidencialidad de la conversación. La temperatura interior se mantuvo en torno a los 21 °C mediante sistema de climatización central, creando una sensación de confort térmico para ambos participantes."),

      h3("Recursos disponibles en el espacio"),
      fichaTable([
        fichaRow("Reproductor de audio", "Sí. Utilizado para musicoterapia ambiental de fondo (frecuencias 432 Hz, música instrumental suave). Volumen bajo, no intrusivo."),
        fichaRow("Material de registro", "Expediente físico en carpeta numerada y bolígrafo. Sin dispositivos electrónicos visibles durante la sesión."),
        fichaRow("Acceso a sanitario", "Sanitario privado a 4 metros del consultorio, de uso exclusivo para pacientes."),
        fichaRow("Sala de espera", "Espacio separado con 3 sillas tapizadas, revistas de divulgación y agua natural. Permite la espera de acompañante si lo hubiera."),
        fichaRow("Accesibilidad", "Rampa de acceso al edificio. Sin escalones en la entrada al consultorio."),
      ]),
      ...el(1),

      h2("III. AMBIENTE EMOCIONAL Y TERAPÉUTICO DEL ESPACIO"),
      body("El espacio ha sido preparado con anterioridad a la llegada de la paciente: los sillones se encuentran orientados correctamente, la temperatura es confortable, la música ambiental está activada a volumen bajo y el expediente físico está sobre el escritorio lateral. La puerta del consultorio permanece cerrada durante toda la sesión para garantizar privacidad."),
      ...el(1),
      body("Al ingresar la paciente, el especialista se encuentra de pie en posición abierta (sin cruzar brazos), a una distancia de aproximadamente 1.5 metros de la puerta, para dar la bienvenida sin invadir el espacio personal. Se ofrece asiento de manera verbal y con un gesto indicativo hacia el sillón destinado a la paciente, permitiendo que ella elija el momento de sentarse."),
      ...el(1),
      body("En todo momento, el entorno está diseñado para comunicar de manera no verbal que el espacio es seguro, confidencial y sin juicio, lo que facilita la apertura emocional necesaria para el trabajo tanatológico con una paciente que atraviesa un duelo de larga duración por pérdida de familiares."),
      ...el(1),
      citeBox("Rosario Robles Galindo. Materiales Didácticos de Tanatología Caso Clínico 3. Expediente y Entrevista Clínica (Evidencia fotográfica e Inicio de Entrevista). Periodo 2026. Benemérita Universidad Autónoma de Puebla, Facultad de Medicina, Licenciatura en Medicina."),

      // ══════════════════════════════════════════════════════════════════════
      //  PUNTO 3: ENTREVISTA — PRIMER CONTACTO Y EMPATÍA
      // ══════════════════════════════════════════════════════════════════════
      pb(),
      subBanner("PUNTO 3 · ENTREVISTA: PRIMER CONTACTO, PRESENTACIÓN Y EMPATÍA"),
      titulo("INICIO DE LA ENTREVISTA CLÍNICA TANATOLÓGICA"),
      ...el(1),

      h2("A) Establecimiento del primer contacto con la paciente"),
      body("El primer contacto se realizó el 16 de abril de 2026 en el consultorio descrito previamente. Al ingresar la paciente al espacio, el especialista la recibió de pie, con postura abierta y expresión tranquila, ofreciendo una bienvenida cálida y sin precipitación. Se le invitó a tomar asiento con calma y se le ofreció agua antes de iniciar cualquier intercambio formal."),
      ...el(1),
      body("La presentación se realizó en los siguientes términos:"),
      ...el(1),
      new Paragraph({
        children: [
          new TextRun({ text: "\"Buenos días. Mi nombre es Jorge Omar González Félix. Estoy aquí para acompañarla en este espacio que es completamente suyo. No hay prisa. Puede hablar con la confianza de que todo lo que se comparta aquí permanece aquí. Estoy a sus órdenes.\"", size: 22, font: "Arial", italics: true, color: "2E4057" })
        ],
        border: {
          left: { style: BorderStyle.THICK, size: 12, color: "2E5797", space: 8 }
        },
        indent: { left: 480 },
        spacing: { before: 100, after: 100 }
      }),
      ...el(1),
      body("La paciente mostró una actitud inicial de reserva, con lenguaje corporal cerrado (manos entrelazadas en el regazo, mirada baja). Se respetó su espacio sin forzar el contacto visual ni el diálogo, permitiendo que el silencio inicial fungiera como recurso de contención."),

      h2("B) Canal de comunicación con la paciente"),
      body("El canal de comunicación establecido fue predominantemente verbal, complementado por recursos no verbales de parte del especialista (asentimiento suave con la cabeza, contacto visual pausado, postura inclinada levemente hacia adelante en señal de escucha activa). Se evitó el uso de vocabulario técnico o clínico que pudiera generar distancia emocional."),
      ...el(1),
      body("Dado el perfil de la paciente —mujer de 43 años, docente con grado de maestría, de ideología católica, cuyo duelo se extiende ya por tres años—, se adoptó un registro comunicativo formal pero cercano, respetuoso de su capacidad reflexiva y de su experiencia de vida. Se validaron sus silencios como parte del proceso y no se interrumpió ninguna de sus respuestas."),
      ...el(1),
      body("Se utilizó la técnica de espejo verbal para reflejar el contenido emocional percibido y confirmar la comprensión sin interpretaciones prematuras. Ejemplos del intercambio inicial:"),
      bullet("\"Entiendo que este camino ha sido largo y que no ha sido fácil llegar aquí.\""),
      bullet("\"No tiene que explicar nada antes de sentirse lista para hacerlo.\""),
      bullet("\"Cualquier emoción que surja aquí es bienvenida y válida.\""),

      h2("C) Observación y recopilación de la primera información visual y oral"),
      h3("Observación no verbal"),
      fichaTable([
        fichaRow("Apariencia general", "Presentación personal cuidada, vestimenta formal-casual en tonos oscuros. Evidencia de autogestión en su imagen externa, posiblemente como mecanismo de control ante el caos emocional interno."),
        fichaRow("Postura corporal", "Cerrada al inicio: piernas cruzadas, manos entrelazadas. Gradual apertura conforme avanza la entrevista."),
        fichaRow("Contacto visual", "Inicial escaso. Se incrementa progresivamente a medida que aumenta la confianza."),
        fichaRow("Tono de voz", "Suave, pausado, con ligeras interrupciones al referirse a sus pérdidas. Control emocional evidente, posiblemente sostenido por el tiempo que lleva con el duelo."),
        fichaRow("Expresión facial", "Contenida. En dos momentos se observó humedecimiento ocular al mencionar brevemente a sus familiares fallecidos, aunque no llegó al llanto."),
        fichaRow("Gestos", "Lleva las manos al regazo repetidamente; en algunos momentos toca levemente su anillo de bodas (posible referencia simbólica al vínculo con su esposo fallecido)."),
      ]),
      ...el(1),
      h3("Primera información oral proporcionada por la paciente"),
      body("Al ser invitada a presentarse brevemente en sus propias palabras, la paciente refirió:"),
      ...el(1),
      new Paragraph({
        children: [new TextRun({ text: "\"Vengo porque ya no sé cómo seguir cargando todo esto sola. Han pasado tres años y hay días en que siento que fue ayer. Perdí a varias personas importantes para mí en poco tiempo y nunca pude... nunca pude despedirme como quería.\"", size: 22, font: "Arial", italics: true, color: "333333" })],
        border: { left: { style: BorderStyle.THICK, size: 12, color: "2E5797", space: 8 } },
        indent: { left: 480 },
        spacing: { before: 100, after: 120 }
      }),
      body("Esta primera expresión espontánea permite identificar de manera preliminar: duelo complicado de larga data, posible duelo no resuelto por ausencia de rituales de despedida, carga emocional sostenida en soledad, y apertura hacia el proceso terapéutico a pesar de la resistencia corporal inicial."),
      ...el(1),
      citeBox("Rosario Robles Galindo. Materiales Didácticos de Tanatología Caso Clínico 3. Expediente y Entrevista Clínica (Evidencia fotográfica e Inicio de Entrevista). Periodo 2026. Benemérita Universidad Autónoma de Puebla, Facultad de Medicina, Licenciatura en Medicina."),

      // ══════════════════════════════════════════════════════════════════════
      //  PUNTO 4: MOTIVO DE CONSULTA
      // ══════════════════════════════════════════════════════════════════════
      pb(),
      subBanner("PUNTO 4 · MOTIVO DE LA CONSULTA Y FACTORES DESENCADENANTES"),
      titulo("MOTIVO DE CONSULTA — IDENTIFICACIÓN DE DESENCADENANTES DEL DUELO"),
      ...el(1),

      h2("1. Motivo principal de la consulta"),
      pregunta("¿Qué necesita usted del profesional en este espacio?"),
      respuesta("La paciente refiere buscar orientación para elaborar el duelo que experimenta desde hace tres años a raíz de la pérdida de varios familiares. Expresa sentirse \"estancada\" emocionalmente: reconoce que su vida funcional continúa —trabaja, cuida su apariencia, cumple con sus obligaciones docentes—, pero percibe que internamente no ha podido avanzar ni encontrar un sentido que le permita integrar la pérdida. Solicita herramientas concretas y un espacio de escucha sin juicio."),
      ...el(1),

      h2("2. Identificación de desencadenantes del duelo"),
      pregunta("¿Cuándo comenzó su duelo? ¿Qué desajustes ha experimentado en su estado emocional, dieta, personalidad, estudios, trabajo, interacción social y vida familiar?"),
      respuesta("La paciente sitúa el inicio de su proceso de duelo hace aproximadamente tres años, cuando experimentó la pérdida de varios integrantes de su familia en un período relativamente corto. Refiere que la primera pérdida fue la más desestructurante, pues no pudo presenciar el momento del fallecimiento ni realizar una despedida formal, lo que dejó un vacío de cierre emocional que persiste hasta hoy."),
      ...el(1),
      body("Los desajustes reportados por área son los siguientes:"),
      ...el(1),
      fichaTable([
        fichaRow("Estado emocional", "Episodios recurrentes de tristeza profunda, especialmente en fechas significativas (aniversarios, celebraciones familiares). Sensación de vacío crónico y dificultad para experimentar alegría plena. Labilidad emocional discreta ante estímulos evocadores (fotografías, canciones, objetos de los fallecidos)."),
        fichaRow("Dieta / hábitos alimenticios", "En los primeros meses del duelo refiere pérdida significativa del apetito y pérdida de peso involuntaria. Actualmente los hábitos alimenticios se han estabilizado, aunque persiste la falta de placer al comer."),
        fichaRow("Personalidad", "Percibe en sí misma mayor hermetismo social y tendencia al aislamiento progresivo. Antes del duelo se describe como una persona sociable y participativa; actualmente prefiere la soledad y evita conversaciones sobre su familia extendida."),
        fichaRow("Trabajo / desempeño docente", "Conserva la funcionalidad laboral como docente, aunque refiere dificultades de concentración en períodos de mayor intensidad del duelo. Ha recibido comentarios de estudiantes sobre una actitud \"más distante\" en clase."),
        fichaRow("Interacción social", "Reducción notable de su red social activa. Asiste a obligaciones sociales básicas pero evita reuniones familiares extensas que le recuerden las ausencias."),
        fichaRow("Vida familiar", "Es viuda. Vive con familia, aunque no especificó el parentesco en esta sesión. Reporta que la dinámica familiar ha cambiado desde las pérdidas y que el tema del duelo no se habla abiertamente en el hogar."),
      ]),
      ...el(1),

      h2("3. Evolución del duelo: ¿Se ha mantenido o ha empeorado?"),
      pregunta("¿Cómo ha ido evolucionando su duelo a lo largo de estos tres años?"),
      respuesta("La paciente describe una evolución irregular. En el primer año, el duelo fue agudo e incapacitante en lo emocional, aunque mantuvo su funcionalidad externa. En el segundo año experimentó una aparente mejoría, con períodos más largos de estabilidad. Sin embargo, en los últimos meses del tercer año percibe un rebrote de la intensidad emocional, posiblemente desencadenado por una fecha significativa o un nuevo evento de pérdida no referido en esta primera sesión. No se evidencia deterioro psicótico ni ideación suicida. Se identifica duelo complicado de tipo crónico."),

      h2("4. ¿Cómo ha enfrentado su duelo?"),
      pregunta("¿Qué ha intentado hacer con su duelo? ¿Cómo ha intentado resolverlo?"),
      respuesta("La paciente ha gestionado su duelo principalmente a través de estrategias de evitación activa: trabajo intensivo, ocupación constante y evitación de estímulos evocadores. Ha recurrido de manera esporádica a la oración y prácticas religiosas (acorde a su ideología católica), lo que refiere le ha brindado alivio temporal pero no una elaboración profunda."),
      ...el(1),
      body("No ha buscado apoyo psicológico ni tanatológico con anterioridad a esta consulta. Identifica la búsqueda de ayuda profesional como un paso que le generó resistencia durante mucho tiempo, pues lo vivía como una señal de debilidad. La decisión de acudir fue catalizada por el reconocimiento de que sus propias estrategias de afrontamiento han sido insuficientes para avanzar."),
      ...el(1),
      citeBox("Rosario Robles Galindo. Materiales Didácticos de Tanatología. Expediente y Entrevista Clínica: Caso Clínico 3. (Motivo de Consulta y Factores desencadenantes/coadyuvantes. de Dieta Emocional). Período 2026. Benemérita Universidad Autónoma de Puebla, Fac/Med."),

      // ══════════════════════════════════════════════════════════════════════
      //  PUNTO 5: PSICOSOCIOGRAMA
      // ══════════════════════════════════════════════════════════════════════
      pb(),
      subBanner("PUNTO 5 · PSICOSOCIOGRAMA"),
      titulo("PSICOSOCIOGRAMA — AMBIENTES FAMILIARES, HÁBITOS Y PARTICIPACIÓN SOCIAL"),
      ...el(1),

      h2("1. Datos sobre el ambiente familiar y el impacto del duelo en la familia"),
      pregunta("¿Cómo es la estructura de su familia? ¿Cómo ha afectado su duelo a sus familiares?"),
      respuesta("La paciente es mujer de 43 años, viuda, que vive actualmente con familia (parentesco no especificado en esta sesión inicial). La estructura familiar original ha sido significativamente modificada por las pérdidas sufridas: personas clave del núcleo familiar han fallecido, lo que ha reconfigurado tanto los roles como las dinámicas relacionales del hogar."),
      ...el(1),
      body("En cuanto al impacto del duelo personal sobre los convivientes, la paciente reconoce que su hermetismo emocional y su tendencia al aislamiento han generado distancia afectiva con las personas con quienes cohabita. Refiere que el tema de la muerte y del duelo se evita dentro del hogar, lo que perpetúa el silencio emocional colectivo y le impide compartir su proceso de manera abierta. Esta dinámica de duelo silenciado en familia representa un factor de riesgo para la cronificación del cuadro."),
      ...el(1),
      fichaTable([
        fichaRow("Composición familiar actual", "Núcleo reducido tras pérdidas. Convive con familia directa (no especificada en primera sesión)."),
        fichaRow("Rol de la paciente en la familia", "Figura central funcional: sostiene responsabilidades domésticas y económicas a pesar del duelo."),
        fichaRow("Comunicación intrafamiliar sobre el duelo", "Escasa o nula. El tema se evita por acuerdo tácito."),
        fichaRow("Apoyo familiar percibido", "Limitado. La paciente no siente que sus familiares comprendan la profundidad de su proceso."),
        fichaRow("Red de apoyo afectivo", "Reducida. Refiere no contar con una persona de confianza con quien hablar abiertamente de sus pérdidas."),
      ]),
      ...el(1),

      h2("2. Hábitos de salud física"),
      pregunta("¿Qué hábitos de salud física tiene usted actualmente? (alimentación, sueño, aseo personal, sexualidad, ejercicio)"),
      ...el(1),
      fichaTable([
        fichaRow("Alimentación", "Actualmente estable en cantidad, aunque con bajo placer al comer. Come sola frecuentemente, lo que reduce el carácter social del acto alimentario."),
        fichaRow("Sueño", "Irregular. Reporta dificultad para conciliar el sueño en períodos de mayor intensidad del duelo, con pensamientos recurrentes sobre los fallecidos al acostarse. No utiliza medicación para dormir."),
        fichaRow("Aseo personal", "Conservado. La presentación personal es uno de los recursos de autogestión que la paciente mantiene con cuidado."),
        fichaRow("Actividad sexual", "No abordada en esta primera sesión. Dado que es viuda, el tema se dejó pendiente para sesiones posteriores con mayor rapport establecido."),
        fichaRow("Ejercicio físico", "Escaso. Refiere que antes del duelo caminaba regularmente; actualmente no mantiene rutina de actividad física consistente."),
        fichaRow("Consumo de sustancias", "Niega consumo de alcohol, tabaco u otras sustancias como mecanismo de afronte."),
      ]),
      ...el(1),

      h2("3. Participación social y grupos de referencia"),
      pregunta("¿Cómo es su participación social actualmente con sus grupos de referencia?"),
      respuesta("La paciente describe una reducción significativa de su vida social en comparación con el período previo a las pérdidas. Distingue claramente entre su yo social previo —activo, participativo, con vínculos amplios— y su yo social actual, que califica como \"retraído\" y \"selectivo\"."),
      ...el(1),
      fichaTable([
        fichaRow("Ámbito laboral / docente", "Es el único grupo social en el que mantiene participación regular y activa, aunque con la distancia afectiva ya referida. El trabajo funciona como estructura de contención externa."),
        fichaRow("Ámbito religioso", "Asistencia ocasional a misa como fuente de alivio espiritual transitorio. No pertenece actualmente a grupos parroquiales o comunidades de fe activas."),
        fichaRow("Amistades", "Ha reducido considerablemente el círculo de amigos. Algunos vínculos se han debilitado por la falta de reciprocidad afectiva en el duelo."),
        fichaRow("Grupos de duelo o apoyo", "No ha participado anteriormente en grupos de duelo. Manifiesta resistencia inicial pero no descarta la posibilidad si le es propuesta como parte del proceso."),
        fichaRow("Actividades recreativas", "Escasas. Refiere no encontrar placer en actividades que antes disfrutaba (anhedonia parcial), lo que se constituye como indicador de duelo complicado."),
      ]),
      ...el(1),
      body("El psicosociograma en conjunto revela a una paciente con alta funcionalidad externa, pero con un mundo interno significativamente empobrecido en vínculos, placer y elaboración emocional. El duelo ha operado como un factor de contracción social y afectiva que demanda una intervención tanatológica estructurada, orientada no solo al procesamiento del dolor, sino a la reconstrucción de su red de significados y relaciones."),
      ...el(1),
      citeBox("Rosario Robles Galindo. Materiales Didácticos de Tanatología. Expediente y Entrevista Clínica: Caso Clínico 3. (Ambientes sociales y participación del paciente). Benemérita Universidad Autónoma de Puebla, Facultad de Medicina, Licenciatura en Medicina."),

      // ── PIE DE PÁGINA ──────────────────────────────────────────────────────
      ...el(1),
      new Paragraph({
        border: { top: { style: BorderStyle.SINGLE, size: 4, color: "AAAAAA", space: 2 } },
        children: [new TextRun({ text: "", size: 4 })], spacing: { before: 200, after: 80 }
      }),
      new Paragraph({
        children: [new TextRun({ text: "Expediente elaborado por: González Félix Jorge Omar  ·  Matrícula 202327266  ·  BUAP, Facultad de Medicina  ·  Tanatología, Caso Clínico 3  ·  16 de abril de 2026", size: 18, font: "Arial", italics: true, color: "777777" })],
        alignment: AlignmentType.CENTER
      }),
    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync("/mnt/user-data/outputs/Etapa1_CasoClinico3_Tanatologia.docx", buf);
  console.log("OK");
});
