import test from 'node:test'
import assert from 'node:assert/strict'
import { readFile } from 'node:fs/promises'
import JSZip from 'jszip'
import { DOMParser, XMLSerializer } from '@xmldom/xmldom'
import { esDocumentoAdmitido, leerDocumento, leerDocx } from '../src/lib/leerDocumento.js'
import { generarActa } from '../src/lib/llenarActa.js'

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
const opciones = { DOMParser, XMLSerializer, salida: 'uint8array' }
const plantilla = await readFile(new URL('../public/Plantilla_Acta_Reunion_CMC-F-05.docx', import.meta.url))
const contenido = await readFile(new URL('./fixtures/acta-completa.txt', import.meta.url), 'utf8')
const escapar = (s) => s.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;')
const parrafo = (s) => `<w:p><w:r><w:t xml:space="preserve">${escapar(s)}</w:t></w:r></w:p>`

async function crearDocx(body) {
  const zip = await JSZip.loadAsync(plantilla)
  zip.file('word/document.xml', `<w:document xmlns:w="${W}"><w:body>${body}<w:sectPr/></w:body></w:document>`)
  return zip.generateAsync({ type: 'uint8array' })
}
const archivo = (name, bytes) => ({ name, arrayBuffer: async () => bytes })

test('admite TXT y DOCX sin distinguir mayúsculas y rechaza DOC y otros formatos', async () => {
  for (const nombre of ['reunion.txt', 'acta.DOCX', 'ACTA.TXT']) assert.ok(esDocumentoAdmitido(nombre))
  for (const nombre of ['acta.doc', 'acta.docm', 'acta.pdf', 'acta.docx.exe']) {
    assert.equal(esDocumentoAdmitido(nombre), false)
    await assert.rejects(leerDocumento(archivo(nombre, new Uint8Array())), /Se admiten archivos TXT y Word/)
  }
})

test('extrae párrafos fragmentados, hipervínculos, saltos y caracteres XML sin incluir texto eliminado', async () => {
  const bytes = await crearDocx(`<w:p><w:r><w:t>###DATOS_</w:t></w:r><w:r><w:t>GENERALES###</w:t><w:br/><w:t>TEMA: Planeación &amp; revisión &lt;2026&gt;</w:t></w:r><w:del><w:r><w:t>eliminado</w:t></w:r></w:del></w:p><w:sdt><w:sdtContent><w:p><w:hyperlink><w:r><w:t>###AGENDA###</w:t></w:r></w:hyperlink></w:p></w:sdtContent></w:sdt>`)
  const texto = await leerDocx(bytes, opciones)
  assert.equal(texto, '###DATOS_GENERALES###\nTEMA: Planeación & revisión <2026>\n###AGENDA###')
})

test('convierte filas de tablas Word en columnas delimitadas para el generador', async () => {
  const bytes = await crearDocx(parrafo('###PARTICIPANTES###') + `<w:tbl><w:tr><w:tc>${parrafo('Asistente')}</w:tc><w:tc>${parrafo('Cargo')}</w:tc></w:tr><w:tr><w:tc>${parrafo('Ana Pérez')}</w:tc><w:tc>${parrafo('Coordinadora')}${parrafo('Regional')}</w:tc></w:tr></w:tbl>`)
  const texto = await leerDocx(bytes, opciones)
  assert.equal(texto, '###PARTICIPANTES###\nAsistente | Cargo\nAna Pérez | Coordinadora Regional')
  const acta = await generarActa(texto, plantilla, opciones)
  assert.deepEqual(acta.acta.participantes, [['Ana Pérez', 'Coordinadora Regional']])
})

test('genera los mismos datos y avisos desde TXT y desde Word con contenido estructurado', async () => {
  const bytes = await crearDocx(contenido.split('\n').map(parrafo).join(''))
  const txt = await leerDocumento(archivo('acta.txt', new TextEncoder().encode(contenido)), opciones)
  const word = await leerDocumento(archivo('acta.DOCX', bytes), opciones)
  const desdeTxt = await generarActa(txt, plantilla, opciones)
  const desdeWord = await generarActa(word, plantilla, opciones)
  assert.deepEqual(desdeWord.acta, desdeTxt.acta)
  assert.deepEqual(desdeWord.avisos, [])
})

test('rechaza Word corrupto, sin cuerpo o sin texto', async () => {
  await assert.rejects(leerDocx(new Uint8Array([1, 2, 3]), opciones), /docx válido/)
  const zip = new JSZip()
  zip.file('otro.xml', '<documento/>')
  await assert.rejects(leerDocx(await zip.generateAsync({ type: 'uint8array' }), opciones), /docx válido/)
  zip.file('word/document.xml', `<w:document xmlns:w="${W}"/>`)
  await assert.rejects(leerDocx(await zip.generateAsync({ type: 'uint8array' }), opciones), /No se pudo leer/)
  await assert.rejects(leerDocx(await crearDocx(parrafo('')), opciones), /no contiene texto/)
})

test('un lote mixto conserva las actas válidas y permite identificar archivos inválidos', async () => {
  const bytes = await crearDocx(contenido.split('\n').map(parrafo).join(''))
  const lote = [archivo('acta.txt', new TextEncoder().encode(contenido)), archivo('acta.docx', bytes), archivo('corrupto.docx', new Uint8Array([0]))]
  const resultados = await Promise.allSettled(lote.map(async (a) => generarActa(await leerDocumento(a, opciones), plantilla, opciones)))
  assert.deepEqual(resultados.map((r) => r.status), ['fulfilled', 'fulfilled', 'rejected'])
})
