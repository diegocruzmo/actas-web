import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import JSZip from "jszip";
import { DOMParser, XMLSerializer } from "@xmldom/xmldom";
import { generarActa, leerTxt, nombreSalida } from "../src/lib/llenarActa.js";

const plantilla = await readFile(new URL("../public/Plantilla_Acta_Reunion_CMC-F-05.docx", import.meta.url));
// Implementaciones de las API XML utilizadas por Node para verificar la generación con la plantilla institucional.
const opciones = { DOMParser, XMLSerializer, salida: "uint8array" };
const contenido = await readFile(new URL("./fixtures/acta-completa.txt", import.meta.url), "utf8");
async function generar(texto = contenido) {
  const resultado = await generarActa(texto, plantilla, opciones);
  const zip = await JSZip.loadAsync(resultado.archivo);
  const xml = await zip.file("word/document.xml").async("string");
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  return { ...resultado, zip, xml, doc, texto: doc.documentElement.textContent };
}
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

test("genera un DOCX completo con la plantilla real y conserva sus recursos", async () => {
  const r = await generar();
  assert.deepEqual(r.avisos, []);
  assert.match(r.texto, /Acta de reunión No. 42/);
  assert.match(r.texto, /Planeación & seguimiento <2026>/);
  assert.match(r.texto, /Se cierra la sesión a las 09:00/);
  assert.match(r.texto, /Ana Pérez/);
  assert.match(r.texto, /Aprobar el plan de trabajo/);
  assert.doesNotMatch(r.texto, /\[[^\]]+\]|Estructura sugerida|Nota de uso/);
  const original = await JSZip.loadAsync(plantilla);
  const archivos = (zip) => Object.keys(zip.files).filter((nombre) => !zip.files[nombre].dir).sort();
  assert.deepEqual(archivos(r.zip), archivos(original));
  for (const nombre of archivos(original).filter((n) => !["word/document.xml", "docProps/core.xml"].includes(n))) {
    assert.deepEqual(await r.zip.file(nombre).async("uint8array"), await original.file(nombre).async("uint8array"));
  }
  assert.match(await r.zip.file("docProps/core.xml").async("string"), /Planeación &amp; seguimiento &lt;2026&gt;/);
});

test("mantiene viñetas cuando las instrucciones están en un único párrafo", async () => {
  const r = await generar();
  const parrafos = [...r.doc.getElementsByTagNameNS(W, "p")].map((p) => p.textContent);
  assert.ok(parrafos.includes("• Se solicitó actualizar el informe."));
});

test("las tablas Markdown no desplazan columnas ni incluyen encabezados", async () => {
  const r = await generar(contenido.replace("Ana Pérez | Coordinadora\nLuis Díaz | Asesor", "| Asistente | Cargo |\n| --- | :---: |\n| Ana Pérez | Coordinadora |\n| Luis Díaz | Asesor |"));
  assert.deepEqual(r.acta.participantes, [["Ana Pérez", "Coordinadora"], ["Luis Díaz", "Asesor"]]);
  assert.deepEqual(r.avisos, []);
});

test("rechaza TXT vacíos, texto libre y secciones sin datos", async () => {
  for (const texto of ["", "  \n", "Transcripción de una reunión sin estructura", "###AGENDA###\n\n###FIN###"]) {
    await assert.rejects(generarActa(texto, plantilla, opciones), /vacío|secciones|datos/);
  }
});

test("conserva datos de secciones repetidas y genera un aviso", async () => {
  const r = await generar(contenido.replace("###AGENDA###", "###PARTICIPANTES###\nMarta Ruiz | Directora\n###AGENDA###"));
  assert.equal(r.acta.participantes.length, 3);
  assert.ok(r.avisos.some((a) => a.includes("repetida")));
});

test("acepta datos incompletos con avisos y valores pendientes", async () => {
  const r = await generar("###DATOS_GENERALES###\nTEMA: Reunión de prueba");
  assert.ok(r.avisos.length > 0);
  assert.match(r.texto, /\[PENDIENTE DE CONFIRMAR\]/);
  assert.match(r.texto, /NO SE IDENTIFICARON DECISIONES EXPRESAS/);
  assert.doesNotMatch(r.texto, /\[NOMBRE COMPLETO|\[ACTIVIDAD|Nota de uso/);
});

test("separa las tablas de pendientes y compromisos", async () => {
  const r = await generar();
  const tablas = [...r.doc.getElementsByTagNameNS(W, "tbl")];
  const pendientes = tablas.find((t) => t.textContent.includes("Validar cifras"));
  const compromisos = tablas.find((t) => t.textContent.includes("Actualizar informe"));
  assert.ok(pendientes && compromisos && pendientes !== compromisos);
});

test("rechaza una plantilla corrupta", async () => {
  await assert.rejects(generarActa(contenido, new Uint8Array([1,2,3]), opciones), /docx válido/);
});

test("decodifica UTF-8, BOM y Windows-1252 sin perder acentos", () => {
  assert.equal(leerTxt(new TextEncoder().encode("\ufeffReunión de planeación")), "Reunión de planeación");
  assert.equal(leerTxt(new Uint8Array([82,101,117,110,105,243,110])), "Reunión");
});

test("nombra el documento a partir del TXT", () => {
  assert.match(nombreSalida("reunion.directiva.TXT"), /^Acta_CMC-F-05_reunion\.directiva_\d{8}_\d{4}\.docx$/);
});

test("elimina las filas vacías de reserva sin eliminar registros de las tablas", async () => {
  const r = await generar();
  const tablas = [...r.doc.getElementsByTagNameNS(W, "tbl")];
  for (const [indice, filasDatos] of [[1, r.acta.participantes], [3, r.acta.pendientes], [5, r.acta.compromisos], [6, r.acta.seguimiento]]) {
    const filas = [...tablas[indice].childNodes].filter((n) => n.namespaceURI === W && n.localName === "tr");
    assert.equal(filas.length, filasDatos.length + 2);
    for (const [i, valores] of filasDatos.entries()) {
      const celdas = [...filas[i + 2].childNodes].filter((n) => n.namespaceURI === W && n.localName === "tc");
      assert.deepEqual(celdas.map((n) => n.textContent), valores);
    }
  }
});

test("permite dividir las filas de contenido sin alterar la protección de los encabezados", async () => {
  const r = await generar();
  let encabezados = 0;
  for (const fila of r.doc.getElementsByTagNameNS(W, "tr")) {
    if (fila.getElementsByTagNameNS(W, "tblHeader").length) {
      encabezados++;
      assert.ok(fila.getElementsByTagNameNS(W, "cantSplit").length > 0);
    } else {
      assert.equal(fila.getElementsByTagNameNS(W, "cantSplit").length, 0);
    }
  }
  assert.ok(encabezados > 0);
});

test("conserva íntegro el desarrollo extenso y permite su continuación entre páginas", async () => {
  const parrafos = Array.from({ length: 30 }, (_, i) => `Asunto ${i + 1}: ${"Se revisaron los avances y las solicitudes recibidas. ".repeat(15)}`.trim());
  const texto = contenido.replace(/###DESARROLLO###[\s\S]*?###COMPROMISOS###/, `###DESARROLLO###\n${parrafos.join("\n\n")}\n###COMPROMISOS###`);
  const r = await generar(texto);
  const desarrollo = [...r.doc.getElementsByTagNameNS(W, "tbl")][4];
  const filas = desarrollo.getElementsByTagNameNS(W, "tr");
  assert.equal(filas[1].getElementsByTagNameNS(W, "cantSplit").length, 0);
  const textos = [...filas[1].getElementsByTagNameNS(W, "p")].map((p) => p.textContent);
  assert.deepEqual(textos, parrafos);
  assert.deepEqual(r.avisos, []);
});
