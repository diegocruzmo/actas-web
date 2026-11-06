import JSZip from 'jszip'
import { leerTxt } from './llenarActa.js'

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

export const esDocumentoAdmitido = (nombre) => /\.(txt|docx)$/i.test(nombre)

// Extracción del contenido visible, sin incluir texto eliminado mediante control de cambios.
function textoVisible(nodo) {
  if (nodo.namespaceURI === W) {
    if (['del', 'moveFrom'].includes(nodo.localName)) return ''
    if (nodo.localName === 't') return nodo.textContent
    if (nodo.localName === 'tab') return '\t'
    if (['br', 'cr'].includes(nodo.localName)) return '\n'
  }
  return Array.from(nodo.childNodes || []).map(textoVisible).join('')
}

// Conservación del orden de párrafos y tablas, incluidas las agrupaciones de contenido de Word.
function bloques(contenedor) {
  const lineas = []
  for (const nodo of Array.from(contenedor.childNodes)) {
    if (nodo.namespaceURI === W && ['del', 'moveFrom'].includes(nodo.localName)) continue
    if (nodo.namespaceURI === W && nodo.localName === 'p') {
      lineas.push(textoVisible(nodo))
    } else if (nodo.namespaceURI === W && nodo.localName === 'tbl') {
      for (const fila of Array.from(nodo.childNodes).filter((n) => n.namespaceURI === W && n.localName === 'tr')) {
        const celdas = Array.from(fila.childNodes).filter((n) => n.namespaceURI === W && n.localName === 'tc')
        lineas.push(celdas.map((celda) => bloques(celda).join(' ').trim()).join(' | '))
      }
    } else if (nodo.nodeType === 1) {
      lineas.push(...bloques(nodo))
    }
  }
  return lineas
}

export async function leerDocx(buffer, opciones = {}) {
  let zip
  try {
    zip = await JSZip.loadAsync(buffer)
  } catch {
    throw new Error('El archivo no es un documento Word .docx válido o está protegido con contraseña.')
  }
  const archivo = zip.file('word/document.xml')
  if (!archivo) throw new Error('El archivo no contiene un documento Word .docx válido.')
  const Parser = opciones.DOMParser || globalThis.DOMParser
  let doc
  try {
    doc = new Parser().parseFromString((await archivo.async('string')).replace(/^\uFEFF/, ''), 'application/xml')
  } catch {
    throw new Error('No se pudo leer el contenido del documento Word.')
  }
  const body = doc.getElementsByTagNameNS(W, 'body')[0]
  if (!body || doc.getElementsByTagName('parsererror').length) {
    throw new Error('No se pudo leer el contenido del documento Word.')
  }
  const contenido = bloques(body).join('\n').normalize('NFC')
  if (!contenido.trim()) throw new Error('El documento Word no contiene texto para procesar.')
  return contenido
}

export async function leerDocumento(archivo, opciones = {}) {
  if (!esDocumentoAdmitido(archivo.name)) {
    throw new Error('Se admiten archivos TXT y Word (.docx). Los archivos .doc deben guardarse como .docx.')
  }
  const buffer = await archivo.arrayBuffer()
  return /\.docx$/i.test(archivo.name) ? leerDocx(buffer, opciones) : leerTxt(buffer)
}
