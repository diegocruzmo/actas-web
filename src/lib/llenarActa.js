// Generación de actas CMC-F-05 mediante la edición del XML del documento DOCX en el navegador.
import JSZip from "jszip";

export const PLANTILLA_POR_DEFECTO = "Plantilla_Acta_Reunion_CMC-F-05.docx";
const PENDIENTE = "[PENDIENTE DE CONFIRMAR]";
const NO_APLICA = "NO APLICA";
const SIN_DECISIONES = "NO SE IDENTIFICARON DECISIONES EXPRESAS EN LA TRANSCRIPCIÓN.";

const SECCIONES = [
  "DATOS_GENERALES",
  "PARTICIPANTES",
  "AGENDA",
  "ACTIVIDADES_PENDIENTES",
  "DESARROLLO",
  "COMPROMISOS",
  "SEGUIMIENTO_DIRECTIVO",
  "DECISIONES",
];

const CLAVES_DATOS = [
  "NUMERO_ACTA", "FECHA", "SITIO", "HORA_INICIO",
  "HORA_CIERRE", "CONVOCA", "TEMA", "OBJETIVO",
];

// Correspondencia entre los campos de la plantilla institucional y las claves del TXT.
const MARCADORES_SIMPLES = {
  "[NÚMERO O CÓDIGO DEL ACTA]": "NUMERO_ACTA",
  "[DÍA de MES de AÑO]": "FECHA",
  "[PRESENCIAL / VIRTUAL / HÍBRIDA - PLATAFORMA O LUGAR]": "SITIO",
  "[HORA DE INICIO]": "HORA_INICIO",
  "[HORA DE CIERRE]": "HORA_CIERRE",
  "[DEPENDENCIA / NOMBRE Y CARGO DE QUIEN CONVOCA]": "CONVOCA",
  "[TEMA PRINCIPAL DE LA REUNIÓN]": "TEMA",
  "[OBJETIVO GENERAL, REDACTADO EN INFINITIVO Y BASADO EN LA TRANSCRIPCIÓN]": "OBJETIVO",
};

// Configuración: encabezados, número de columnas e índice de coincidencia desde cero.
const TABLA_PARTICIPANTES = [["asistente", "cargo"], 2, 0];
const TABLA_PENDIENTES = [["actividad", "responsable", "fecha de ejecucion"], 3, 0];
// Las tablas de pendientes y compromisos comparten encabezados; la segunda coincidencia corresponde a compromisos.
const TABLA_COMPROMISOS = [["actividad", "responsable", "fecha de ejecucion"], 3, 1];
const TABLA_SEGUIMIENTO = [["tema", "riesgo u observacion", "responsable", "accion de seguimiento"], 4, 0];

const RE_AGENDA =
  /\[(PUNTO DE AGENDA \d+|OTROS PUNTOS IDENTIFICADOS EN LA TRANSCRIPCIÓN|CIERRE Y PRÓXIMOS PASOS)\]/;
const RE_DECISIONES = /\[(DECISIÓN ADOPTADA \d+|OTRAS DECISIONES EXPRESAMENTE ACORDADAS)\]/;
const MARCADOR_DESARROLLO = "[REDACTAR UN RELATO";
const MODELO_VINETA = "Apertura e instalación";

const TEXTOS_INSTRUCTIVOS = [
  "Estructura sugerida",
  "Apertura e instalación.",
  "Presentación de los temas y lineamientos.",
  "Intervenciones relevantes por participante, dependencia o territorio.",
  "Solicitudes, alertas, riesgos y aclaraciones.",
  "Respuestas u orientaciones impartidas.",
  "Cierre y próximos pasos.",
  "No atribuir afirmaciones si el hablante",
  "Marcar toda duda como",
  "Si no hubo decisiones expresas, registrar",
  "Nota de uso: los textos entre corchetes",
];

// El valor de confirmación pendiente es válido y se excluye de los marcadores sin resolver.
const SRC_MARCADOR_RESTANTE = String.raw`\[(?!PENDIENTE DE CONFIRMAR\])[^\[\]]{3,}\]`;
const RE_MARCADOR_RESTANTE = new RegExp(SRC_MARCADOR_RESTANTE);
const RE_NUMERO_LITERAL = /^\s*(\d+)\s*[.)]\s*/;
const RE_VINETA_LITERAL = /^\s*([•·▪●\-–])\s*/;

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const XML_NS = "http://www.w3.org/XML/1998/namespace";
const DC = "http://purl.org/dc/elements/1.1/";
const MIME_DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

const nfc = (t) => (t || "").normalize("NFC");

const sinTildes = (t) =>
  (t || "").normalize("NFD").replace(/\p{Mn}/gu, "").toLowerCase().trim();

const splitlines = (t) => {
  const partes = t.split(/\r\n|[\n\r\v\f\x1c-\x1e\x85\u2028\u2029]/);
  if (partes.length && partes[partes.length - 1] === "") partes.pop();
  return partes;
};

// Decodificación UTF-8 con alternativa Windows-1252 para secuencias de bytes incompatibles.
export function leerTxt(buffer) {
  const bytes = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);
  try {

    return nfc(new TextDecoder("utf-8", { fatal: true }).decode(bytes));
  } catch {

    return nfc(new TextDecoder("windows-1252").decode(bytes));
  }
}

function parsearTxt(contenido, aviso) {
  if (!contenido.trim()) throw new Error("El archivo de entrada está vacío.");

  // Eliminación de los delimitadores de bloques Markdown que pueden envolver el TXT del agente.
  contenido = contenido.replace(/^```.*$/gm, "");

  const secciones = {};
  let actual = null;
  for (const linea of splitlines(contenido)) {
    const m = linea.match(/^\s*###\s*([A-ZÁÉÍÓÚÑ_ ]+?)\s*###\s*$/);
    if (m) {
      const nombre = sinTildes(m[1]).toUpperCase().replaceAll(" ", "_");
      if (nombre === "FIN") break;
      actual = nombre;
      if (Object.hasOwn(secciones, actual)) {
        aviso(`La sección ${actual} está repetida; se conservaron sus datos.`);
      } else secciones[actual] = [];
      continue;
    }
    if (actual !== null) secciones[actual].push(linea.trimEnd());
  }

  if (!SECCIONES.some((s) => Object.hasOwn(secciones, s))) {
    throw new Error("El documento no contiene secciones de acta reconocibles. Se requiere el archivo estructurado del agente (###DATOS_GENERALES###, ###PARTICIPANTES###, etc.).");
  }
  if (!SECCIONES.some((s) => secciones[s]?.some((linea) => linea.trim()))) {
    throw new Error("Las secciones del documento no contienen datos para generar el acta.");
  }
  for (const s of SECCIONES) {
    if (!(s in secciones)) {
      aviso(`Falta la sección ###${s}### en el TXT.`);
      secciones[s] = [];
    }
  }

  return {
    datos: parsearDatos(secciones.DATOS_GENERALES, aviso),
    participantes: parsearTabla(secciones.PARTICIPANTES, 2, "PARTICIPANTES", aviso),
    agenda: parsearLista(secciones.AGENDA),
    pendientes: parsearTabla(secciones.ACTIVIDADES_PENDIENTES, 3, "ACTIVIDADES_PENDIENTES", aviso),
    desarrollo: parsearDesarrollo(secciones.DESARROLLO, aviso),
    compromisos: parsearTabla(secciones.COMPROMISOS, 3, "COMPROMISOS", aviso),
    seguimiento: parsearTabla(secciones.SEGUIMIENTO_DIRECTIVO, 4, "SEGUIMIENTO_DIRECTIVO", aviso),
    decisiones: parsearLista(secciones.DECISIONES),
  };
}

function parsearDatos(lineas, aviso) {
  const datos = {};
  for (const linea of lineas) {
    const m = linea.match(/^\s*([A-Z_]+)\s*:\s*(.*)$/);
    if (m) datos[m[1].trim()] = m[2].trim();
  }
  for (const clave of CLAVES_DATOS) {
    if (!datos[clave]) {
      aviso(`DATOS_GENERALES: '${clave}' vacío o ausente; se usará ${PENDIENTE}.`);
      datos[clave] = PENDIENTE;
    }
  }
  return datos;
}

function parsearTabla(lineas, nCols, nombre, aviso) {
  const encabezados = {
    PARTICIPANTES: ["asistente", "cargo"],
    ACTIVIDADES_PENDIENTES: ["actividad", "responsable", "fecha de ejecucion"],
    COMPROMISOS: ["actividad", "responsable", "fecha de ejecucion"],
    SEGUIMIENTO_DIRECTIVO: ["tema", "riesgo u observacion", "responsable", "accion de seguimiento"],
  };
  let filas = [];
  for (const linea of lineas) {
    if (!linea.trim()) continue;

    // Eliminación de los delimitadores externos de tablas Markdown antes de separar las columnas.
    const texto = linea.trim().replace(/^\|/, "").replace(/\|$/, "");
    let celdas = texto.split("|").map((c) => c.trim());
    if (celdas.every((c) => /^:?-{3,}:?$/.test(c))) continue;
    if (celdas.length === nCols && celdas.every((c, i) => sinTildes(c) === encabezados[nombre]?.[i])) continue;
    if (celdas.length !== nCols) {
      aviso(`${nombre}: la fila '${linea.slice(0, 60)}...' tiene ${celdas.length} columnas ` +
            `(se esperaban ${nCols}). Se ajustó automáticamente.`);
      if (celdas.length > nCols) {
        celdas = [...celdas.slice(0, nCols - 1), celdas.slice(nCols - 1).join(" / ")];
      } else {
        celdas = [...celdas, ...Array(nCols - celdas.length).fill(PENDIENTE)];
      }
    }
    filas.push(celdas.map((c) => c || PENDIENTE));
  }
  if (!filas.length) {
    aviso(`${nombre}: sin filas; se registró ${NO_APLICA}.`);
    filas = [Array(nCols).fill(NO_APLICA)];
  }
  return filas;
}

function parsearLista(lineas) {
  const items = [];
  for (const linea of lineas) {
    let t = linea.trim();
    if (!t) continue;
    t = t.replace(RE_NUMERO_LITERAL, "");
    t = t.replace(RE_VINETA_LITERAL, "");
    items.push(t);
  }
  return items;
}

// Representación del desarrollo en bloques [tipo, texto]: párrafo (p) o viñeta (v).
function parsearDesarrollo(lineas, aviso) {
  let bloques = [];
  const buffer = [];
  const vaciar = () => {
    if (buffer.length) {
      bloques.push(["p", buffer.join(" ")]);
      buffer.length = 0;
    }
  };
  for (const linea of lineas) {
    const t = linea.trim();
    if (!t) vaciar();
    else if (/^[-•]\s+/.test(t)) {
      vaciar();
      bloques.push(["v", t.replace(/^[-•]\s+/, "")]);
    } else buffer.push(t);
  }
  vaciar();
  if (!bloques.length) {
    aviso(`DESARROLLO vacío; se registró ${PENDIENTE}.`);
    bloques = [["p", PENDIENTE]];
  }
  return bloques;
}

const esW = (n, local) =>
  !!n && n.nodeType === 1 && n.namespaceURI === W && n.localName === local;

const hijos = (el, local) => {
  const r = [];
  for (let n = el.firstChild; n; n = n.nextSibling) if (esW(n, local)) r.push(n);
  return r;
};

const hijo = (el, local) => hijos(el, local)[0] || null;

const anteriorElemento = (el) => {
  let n = el.previousSibling;
  while (n && n.nodeType !== 1) n = n.previousSibling;
  return n;
};

const insertarDespues = (ref, nuevo) => {
  ref.parentNode.insertBefore(nuevo, ref.nextSibling);
  return nuevo;
};

const quitar = (el) => el.parentNode && el.parentNode.removeChild(el);

function textoRun(r) {
  let t = "";
  for (let n = r.firstChild; n; n = n.nextSibling) {
    if (esW(n, "t")) t += n.textContent;
    else if (esW(n, "tab")) t += "\t";
    else if (esW(n, "br") || esW(n, "cr")) t += "\n";
    else if (esW(n, "noBreakHyphen")) t += "-";
  }
  return t;
}

// Sustitución del contenido del fragmento de Word con conservación de sus propiedades de formato.
function ponerTextoRun(r, texto) {
  const doc = r.ownerDocument;
  for (let n = r.firstChild; n; ) {
    const sig = n.nextSibling;
    if (!esW(n, "rPr")) r.removeChild(n);
    n = sig;
  }
  let buf = "";
  const volcar = () => {
    if (!buf) return;
    const t = doc.createElementNS(W, "w:t");
    if (/^\s|\s$/.test(buf)) t.setAttributeNS(XML_NS, "xml:space", "preserve");
    t.appendChild(doc.createTextNode(buf));
    r.appendChild(t);
    buf = "";
  };
  for (const ch of texto) {
    if (ch === "\t" || ch === "\n") {
      volcar();
      r.appendChild(doc.createElementNS(W, ch === "\t" ? "w:tab" : "w:br"));
    } else buf += ch;
  }
  volcar();
}

const runs = (p) => hijos(p, "r");

function textoParrafo(p) {
  let t = "";
  for (let n = p.firstChild; n; n = n.nextSibling) {
    if (esW(n, "r")) t += textoRun(n);
    else if (esW(n, "hyperlink")) for (const r of hijos(n, "r")) t += textoRun(r);
  }
  return t;
}

const textoCelda = (tc) => hijos(tc, "p").map(textoParrafo).join("\n");

const celdasUnicas = (tr) => hijos(tr, "tc");

function* iterarTablas(contenedor) {
  for (const tbl of hijos(contenedor, "tbl")) {
    yield tbl;
    for (const tr of hijos(tbl, "tr"))
      for (const tc of celdasUnicas(tr)) yield* iterarTablas(tc);
  }
}

function* iterarParrafos(contenedor) {
  for (const p of hijos(contenedor, "p")) yield p;
  for (const tbl of hijos(contenedor, "tbl"))
    for (const tr of hijos(tbl, "tr"))
      for (const tc of celdasUnicas(tr)) yield* iterarParrafos(tc);
}

// Un marcador puede estar dividido en varios fragmentos de texto por el formato de Word.
function reemplazarEnParrafo(p, buscar, nuevo) {
  let reemplazos = 0;
  let desde = 0;
  while (true) {
    const rs = runs(p);
    const textos = rs.map((r) => nfc(textoRun(r)));
    const completo = textos.join("");
    const pos = completo.indexOf(buscar, desde);
    if (pos < 0 || !rs.length) return reemplazos;
    const fin = pos + buscar.length;
    let acumulado = 0;
    let primero = true;
    rs.forEach((run, i) => {
      const txt = textos[i];
      const iniRun = acumulado;
      const finRun = acumulado + txt.length;
      acumulado = finRun;
      if (finRun <= pos || iniRun >= fin) return;
      const a = Math.max(pos, iniRun) - iniRun;
      const b = Math.min(fin, finRun) - iniRun;
      if (primero) {
        ponerTextoRun(run, txt.slice(0, a) + nuevo + txt.slice(b));
        primero = false;
      } else {
        ponerTextoRun(run, txt.slice(0, a) + txt.slice(b));
      }
    });
    // La búsqueda continúa después del valor insertado para impedir su procesamiento como un nuevo marcador.
    desde = pos + nuevo.length; 
    reemplazos++;
  }
}

function ponerTextoParrafo(p, texto) {
  const rs = runs(p);
  if (!rs.length) {
    const r = p.ownerDocument.createElementNS(W, "w:r");
    p.appendChild(r);
    ponerTextoRun(r, texto);
    return;
  }
  ponerTextoRun(rs[0], texto);
  for (const r of rs.slice(1)) quitar(r);
}

function ponerTextoCelda(tc, texto) {
  let ps = hijos(tc, "p");
  if (!ps.length) {
    tc.appendChild(tc.ownerDocument.createElementNS(W, "w:p"));
    ps = hijos(tc, "p");
  }
  ponerTextoParrafo(ps[0], texto);
  for (const p of ps.slice(1)) eliminarParrafo(p);
}

function eliminarParrafo(p) {
  const padre = p.parentNode;
  if (!padre) return;

  // Word requiere conservar al menos un párrafo dentro de cada celda.
  if (esW(padre, "tc") && hijos(padre, "p").length <= 1) {
    ponerTextoParrafo(p, "");
    return;
  }
  padre.removeChild(p);
}

const clonarParrafoDespues = (ref, modelo) => insertarDespues(ref, modelo.cloneNode(true));

// numId = 0 desactiva también la numeración heredada del estilo del párrafo.
function quitarNumeracion(p) {
  const doc = p.ownerDocument;
  let ppr = hijo(p, "pPr");
  if (!ppr) {
    ppr = doc.createElementNS(W, "w:pPr");
    p.insertBefore(ppr, p.firstChild);
  }
  const viejo = hijo(ppr, "numPr");
  if (viejo) ppr.removeChild(viejo);
  const numpr = doc.createElementNS(W, "w:numPr");
  const ilvl = doc.createElementNS(W, "w:ilvl");
  ilvl.setAttributeNS(W, "w:val", "0");
  const numid = doc.createElementNS(W, "w:numId");
  numid.setAttributeNS(W, "w:val", "0");
  numpr.appendChild(ilvl);
  numpr.appendChild(numid);

  const pstyle = hijo(ppr, "pStyle");
  if (pstyle) insertarDespues(pstyle, numpr);
  else ppr.insertBefore(numpr, ppr.firstChild);
}

function llenarCamposSimples(body, datos) {
  for (const p of iterarParrafos(body)) {
    const texto = nfc(textoParrafo(p));
    if (!texto.includes("[")) continue;
    for (const [marcador, clave] of Object.entries(MARCADORES_SIMPLES)) {
      if (texto.includes(marcador)) reemplazarEnParrafo(p, marcador, datos[clave] ?? PENDIENTE);
    }
  }
}

function buscarEncabezado(body, palabras, ocurrencia) {
  let encontrados = 0;
  for (const tbl of iterarTablas(body)) {
    const filas = hijos(tbl, "tr");
    for (let i = 0; i < filas.length; i++) {
      const textos = celdasUnicas(filas[i]).map((tc) => sinTildes(textoCelda(tc)));
      if (palabras.every((p) => textos.some((t) => t === p || t.startsWith(p)))) {
        if (encontrados === ocurrencia) return [tbl, i];
        encontrados++;
      }
    }
  }
  return [null, null];
}

// Las filas de reserva vacías de la plantilla no deben conservarse como registros del acta.
function esFilaReservaVacia(tr, nCols) {
  const celdas = celdasUnicas(tr);
  if (celdas.length !== nCols || celdas.some((tc) => textoCelda(tc).trim())) return false;
  return !["drawing", "pict", "object", "sdt", "fldChar", "instrText", "br", "tab", "vMerge", "sectPr"].some(
    (local) => tr.getElementsByTagNameNS(W, local).length
  );
}

// La división de filas permite que los textos extensos continúen en la página siguiente.
function permitirDivisionFila(tr) {
  const propiedades = hijo(tr, "trPr");
  if (propiedades) for (const restriccion of hijos(propiedades, "cantSplit")) quitar(restriccion);
}

function ajustarPaginacion(body) {
  for (const tabla of iterarTablas(body)) {
    for (const fila of hijos(tabla, "tr")) {
      const encabezado = hijo(hijo(fila, "trPr") || fila, "tblHeader");
      const esEncabezado = encabezado && !["0", "false", "off"].includes(encabezado.getAttributeNS(W, "val"));
      if (!esEncabezado) permitirDivisionFila(fila);
    }
  }
}

function llenarTabla(body, config, filasDatos, nombre, aviso) {
  const [palabras, nCols, ocurrencia] = config;
  const [tbl, idx] = buscarEncabezado(body, palabras, ocurrencia);
  if (!tbl) {
    aviso(`No se encontró en la plantilla la tabla de ${nombre}.`);
    return;
  }
  const filas = hijos(tbl, "tr");

  // Reutilización del formato de las filas con marcadores para insertar los datos del acta.
  let filasModelo = [];
  for (const tr of filas.slice(idx + 1)) {
    const txt = nfc(celdasUnicas(tr).map(textoCelda).join(" "));
    if (RE_MARCADOR_RESTANTE.test(txt) || (filasModelo.length && esFilaReservaVacia(tr, nCols))) filasModelo.push(tr);
    else break;
  }
  if (!filasModelo.length) filasModelo = filas.length > idx + 1 ? [filas[idx + 1]] : [];
  if (!filasModelo.length) {
    aviso(`La tabla de ${nombre} no tiene una fila modelo para copiar.`);
    return;
  }

  const modeloTr = filasModelo[0].cloneNode(true);
  let ancla = anteriorElemento(filasModelo[0]);
  for (const tr of filasModelo) quitar(tr);

  for (const valores of filasDatos) {
    const nuevoTr = insertarDespues(ancla, modeloTr.cloneNode(true));
    ancla = nuevoTr;
    const celdas = celdasUnicas(nuevoTr);
    if (celdas.length !== nCols) {
      aviso(`La tabla de ${nombre} tiene ${celdas.length} columnas en la plantilla ` +
            `(se esperaban ${nCols}).`);
    }
    celdas.forEach((tc, i) => {
      if (i < valores.length) ponerTextoCelda(tc, valores[i]);
    });
  }
}

function llenarLista(body, regex, items, nombre, aviso, textoSiVacio = null) {
  const modelos = [...iterarParrafos(body)].filter((p) => regex.test(nfc(textoParrafo(p))));
  if (!modelos.length) {
    aviso(`No se encontraron en la plantilla los marcadores de ${nombre}.`);
    return;
  }
  const primero = modelos[0];
  const numeracionLiteral = RE_NUMERO_LITERAL.test(textoParrafo(primero));
  const modeloElem = primero.cloneNode(true);

  const sinItems = !items.length;
  if (sinItems) items = [textoSiVacio || NO_APLICA];
  const quitarPunto = (s) => s.replace(/\.+$/, "");
  const unicoSinNumero =
    sinItems || (items.length === 1 && quitarPunto(items[0]) === quitarPunto(SIN_DECISIONES));

  let ancla = primero;
  const creados = [];
  for (let i = 0; i < items.length; i++) {
    ancla = clonarParrafoDespues(ancla, modeloElem);
    creados.push(ancla);
  }
  for (const p of modelos) quitar(p);

  creados.forEach((p, i) => {
    const item = items[i];
    if (unicoSinNumero) {
      quitarNumeracion(p);
      ponerTextoParrafo(p, item);
    } else if (numeracionLiteral) {
      ponerTextoParrafo(p, `${i + 1}. ${item}`);
    } else {
      ponerTextoParrafo(p, item);
    }
  });
}

function capturarModeloVineta(body) {
  for (const p of iterarParrafos(body)) {
    if (nfc(textoParrafo(p)).includes(MODELO_VINETA)) {
      const m = textoParrafo(p).match(RE_VINETA_LITERAL);

      // Los párrafos que agrupan todas las instrucciones se excluyen como modelos de viñeta.
      if (!m && !hijo(hijo(p, "pPr") || p, "numPr")) return [null, null];
      return [p.cloneNode(true), m ? m[1] : null];
    }
  }
  return [null, null];
}

function llenarDesarrollo(body, bloques, modeloV, vinetaLiteral, aviso) {
  const modelo = [...iterarParrafos(body)].find((p) =>
    nfc(textoParrafo(p)).includes(MARCADOR_DESARROLLO));
  if (!modelo) {
    aviso("No se encontró en la plantilla el marcador del DESARROLLO DE LA AGENDA.");
    return;
  }
  const modeloP = modelo.cloneNode(true);
  let ancla = modelo;
  for (const [tipo, texto] of bloques) {
    let nuevo;
    let contenido;
    if (tipo === "v" && modeloV) {
      nuevo = clonarParrafoDespues(ancla, modeloV);
      contenido = vinetaLiteral ? `${vinetaLiteral} ${texto}` : texto;
    } else if (tipo === "v") {
      nuevo = clonarParrafoDespues(ancla, modeloP);
      contenido = `• ${texto}`;
    } else {
      nuevo = clonarParrafoDespues(ancla, modeloP);
      contenido = texto;
    }
    ponerTextoParrafo(nuevo, contenido);
    ancla = nuevo;
  }
  quitar(modelo);
}

function eliminarInstructivos(body) {
  for (const p of [...iterarParrafos(body)]) {
    const texto = nfc(textoParrafo(p)).replace(RE_VINETA_LITERAL, "").trim();
    if (!texto) continue;
    for (const instr of TEXTOS_INSTRUCTIVOS) {
      const coincide = instr.endsWith(".") ? texto === instr : texto.startsWith(instr);
      if (coincide) {
        eliminarParrafo(p);
        break;
      }
    }
  }
}

function revisarRestantes(body, aviso) {
  const restantes = new Set();
  for (const p of iterarParrafos(body)) {
    for (const m of nfc(textoParrafo(p)).matchAll(new RegExp(SRC_MARCADOR_RESTANTE, "g"))) {
      restantes.add(m[0]);
    }
  }
  for (const m of [...restantes].sort()) {
    aviso(`Quedó un marcador sin reemplazar en el documento: ${m}`);
  }
}

function serializar(doc, Serializer) {
  let xml = new Serializer().serializeToString(doc);
  if (!xml.startsWith("<?xml")) {
    xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + xml;
  }
  return xml;
}

function parsearXml(xml, Parser, nombre) {
  const doc = new Parser().parseFromString(xml.replace(/^\uFEFF/, ""), "application/xml");
  if (doc.getElementsByTagName("parsererror").length) {
    throw new Error(`No se pudo leer ${nombre} dentro de la plantilla.`);
  }
  return doc;
}

function ponerTitulo(coreXml, titulo, Parser, Serializer) {
  const doc = parsearXml(coreXml, Parser, "docProps/core.xml");
  let el = doc.getElementsByTagNameNS(DC, "title")[0];
  if (!el) {
    el = doc.createElementNS(DC, "dc:title");
    doc.documentElement.appendChild(el);
  }
  while (el.firstChild) el.removeChild(el.firstChild);
  el.appendChild(doc.createTextNode(titulo));
  return serializar(doc, Serializer);
}

/**
 * Generación del documento a partir de la plantilla, con devolución del archivo, los avisos y los datos interpretados.
 * opciones.salida controla el formato de salida (blob por defecto).
 * Las pruebas proporcionan DOMParser y XMLSerializer mediante opciones.
 */
export async function generarActa(contenidoTxt, plantilla, opciones = {}) {
  const Parser = opciones.DOMParser || globalThis.DOMParser;
  const Serializer = opciones.XMLSerializer || globalThis.XMLSerializer;
  const avisos = [];
  const aviso = (msg) => avisos.push(msg);

  const acta = parsearTxt(nfc(contenidoTxt), aviso);

  let zip;
  try {
    zip = await JSZip.loadAsync(plantilla);
  } catch {
    throw new Error("La plantilla no es un archivo .docx válido.");
  }
  const docFile = zip.file("word/document.xml");
  if (!docFile) throw new Error("La plantilla no es un archivo .docx válido.");
  const doc = parsearXml(await docFile.async("string"), Parser, "word/document.xml");
  const body = doc.getElementsByTagNameNS(W, "body")[0];
  if (!body) throw new Error("La plantilla no tiene cuerpo de documento.");

  // Conservación del modelo de viñeta antes de eliminar las instrucciones de la plantilla.
  const [modeloV, vinetaLiteral] = capturarModeloVineta(body);
  eliminarInstructivos(body);

  llenarCamposSimples(body, acta.datos);
  llenarTabla(body, TABLA_PARTICIPANTES, acta.participantes, "PARTICIPANTES", aviso);
  llenarLista(body, RE_AGENDA, acta.agenda, "AGENDA", aviso);
  llenarTabla(body, TABLA_PENDIENTES, acta.pendientes, "ACTIVIDADES PENDIENTES", aviso);
  llenarDesarrollo(body, acta.desarrollo, modeloV, vinetaLiteral, aviso);
  llenarTabla(body, TABLA_COMPROMISOS, acta.compromisos, "ACTIVIDADES A REALIZAR", aviso);
  llenarTabla(body, TABLA_SEGUIMIENTO, acta.seguimiento, "SEGUIMIENTO DIRECTIVO", aviso);
  llenarLista(body, RE_DECISIONES, acta.decisiones, "DECISIONES", aviso, SIN_DECISIONES);
  ajustarPaginacion(body);
  revisarRestantes(body, aviso);

  zip.file("word/document.xml", serializar(doc, Serializer));

  const tema = acta.datos.TEMA || "";
  const core = zip.file("docProps/core.xml");
  if (tema && tema !== PENDIENTE && core) {
    zip.file("docProps/core.xml",
      ponerTitulo(await core.async("string"), `Acta CMC-F-05 - ${tema}`, Parser, Serializer));
  }

  const archivo = await zip.generateAsync({
    type: opciones.salida || "blob",
    mimeType: MIME_DOCX,
    compression: "DEFLATE",
  });
  return { archivo, avisos, acta };
}

// El nombre del documento incorpora el nombre base del archivo de entrada y la fecha y hora locales.
export function nombreSalida(nombreTxt) {
  const base = nombreTxt.replace(/\.[^.]*$/, "");
  const d = new Date();
  const dos = (n) => String(n).padStart(2, "0");
  const sello = `${d.getFullYear()}${dos(d.getMonth() + 1)}${dos(d.getDate())}` +
                `_${dos(d.getHours())}${dos(d.getMinutes())}`;
  return `Acta_CMC-F-05_${base}_${sello}.docx`;
}
