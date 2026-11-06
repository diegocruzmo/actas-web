import { useEffect, useRef, useState } from 'react'
import JSZip from 'jszip'
import {
  generarActa,
  nombreSalida,
  PLANTILLA_POR_DEFECTO
} from './lib/llenarActa'
import { esDocumentoAdmitido, leerDocumento } from './lib/leerDocumento.js'

function Icono({ tipo, ...props }) {
  const trazos = {
    archivo: (
      <>
        <path d='M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8Z' />
        <path d='M14 2v6h6M8 13h8M8 17h5' />
      </>
    ),
    subir: <path d='M12 16V3m-5 5 5-5 5 5M4 16v5h16v-5' />,
    descargar: <path d='M12 3v12m-5-5 5 5 5-5M4 16v5h16v-5' />,
    seguro: (
      <>
        <path d='m12 2 8 4v6c0 5-8 10-8 10S4 17 4 12V6Z' />
        <path d='m8 12 3 3 5-6' />
      </>
    )
  }
  return (
    <svg
      width='22'
      height='22'
      viewBox='0 0 24 24'
      fill='none'
      stroke='currentColor'
      strokeWidth='1.6'
      strokeLinecap='round'
      strokeLinejoin='round'
      aria-hidden='true'
      {...props}
    >
      {trazos[tipo]}
    </svg>
  )
}

export default function App() {
  const [archivos, setArchivos] = useState([])
  const [procesando, setProcesando] = useState(false)
  const [progreso, setProgreso] = useState(0)
  const [resultados, setResultados] = useState([])
  const [paquete, setPaquete] = useState(null)
  const [error, setError] = useState('')
  const [arrastrando, setArrastrando] = useState(false)
  const urls = useRef([])
  const ocupado = useRef(false)
  const montado = useRef(true)
  const inputRef = useRef(null)

  // Liberación de las URL de descarga temporales al iniciar un lote o desmontar la vista.
  function liberarUrls() {
    urls.current.forEach((url) => URL.revokeObjectURL(url))
    urls.current = []
  }
  useEffect(() => {
    montado.current = true
    return () => {
      montado.current = false
      liberarUrls()
    }
  }, [])
  function crearUrl(blob) {
    const url = URL.createObjectURL(blob)
    urls.current.push(url)
    return url
  }
  function agregarArchivos(seleccion) {
    if (ocupado.current) return
    const nuevos = Array.from(seleccion || [])
    const invalidos = nuevos.filter((a) => !esDocumentoAdmitido(a.name))
    setError(
      invalidos.length
        ? `Se admiten archivos .txt y .docx. Los archivos .doc deben guardarse como .docx. No se agregaron: ${invalidos.map((a) => a.name).join(', ')}.`
        : ''
    )
    // Validación de archivos duplicados por nombre, tamaño y fecha de modificación.
    setArchivos((anteriores) => {
      const lista = [...anteriores]
      for (const a of nuevos.filter((a) => esDocumentoAdmitido(a.name))) {
        if (
          !lista.some(
            (b) =>
              b.name === a.name &&
              b.size === a.size &&
              b.lastModified === a.lastModified
          )
        )
          lista.push(a)
      }
      return lista
    })
  }
  async function onGenerar(e) {
    e.preventDefault()
    // El bloqueo por referencia impide el inicio simultáneo de lotes durante la actualización del estado de React.
    if (ocupado.current) return
    if (!archivos.length) {
      setError('Seleccione al menos un archivo TXT o Word (.docx) para generar las actas.')
      return
    }
    ocupado.current = true
    setProcesando(true)
    setError('')
    setProgreso(0)
    setResultados([])
    setPaquete(null)
    liberarUrls()
    try {
      // La plantilla institucional publicada con la aplicación se utiliza para todos los documentos del lote.
      const resp = await fetch(
        `${import.meta.env.BASE_URL}${PLANTILLA_POR_DEFECTO}`
      )
      if (!resp.ok)
        throw new Error(
          'No se pudo cargar la plantilla institucional. Verifica que esté disponible en la carpeta public e inténtalo de nuevo.'
        )
      const plantilla = await resp.arrayBuffer()
      // La respuesta puede contener HTML con estado HTTP 200 si la plantilla no está disponible.
      const zipPlantilla = await JSZip.loadAsync(plantilla).catch(() => null)
      if (!zipPlantilla?.file('word/document.xml'))
        throw new Error(
          'La plantilla institucional no es un documento Word válido. Verifica el archivo de la carpeta public.'
        )
      const lote = []
      const zip = new JSZip()
      const nombres = new Set()
      // Procesamiento independiente de cada documento para conservar los resultados ante errores en otros archivos.
      for (const [indice, entrada] of archivos.entries()) {
        if (!montado.current) return
        try {
          const contenido = await leerDocumento(entrada)
          const { archivo, avisos } = await generarActa(contenido, plantilla)
          if (!montado.current) return
          const original = nombreSalida(entrada.name)
          let nombre = original
          let consecutivo = 2
          // Asignación de nombres únicos para impedir la sobrescritura de documentos en el archivo ZIP.
          while (nombres.has(nombre.toLowerCase()))
            nombre = original.replace(/\.docx$/, `_${consecutivo++}.docx`)
          nombres.add(nombre.toLowerCase())
          const bytes = await archivo.arrayBuffer()
          if (!montado.current) return
          zip.file(nombre, bytes)
          lote.push({
            origen: entrada.name,
            nombre,
            url: crearUrl(archivo),
            avisos
          })
        } catch (err) {
          lote.push({
            origen: entrada.name,
            error: err?.message || 'No se pudo generar esta acta.'
          })
        }
        if (!montado.current) return
        setResultados([...lote])
        setProgreso(indice + 1)
      }
      // La descarga ZIP se habilita cuando existen al menos dos documentos generados.
      if (lote.filter((r) => !r.error).length > 1) {
        try {
          const blob = await zip.generateAsync({
            type: 'blob',
            compression: 'DEFLATE'
          })
          if (montado.current)
            setPaquete({ url: crearUrl(blob), nombre: 'Actas_CMC-F-05.zip' })
        } catch {
          if (montado.current)
            setError(
              'Las actas están listas, pero no se pudo crear el ZIP. Puedes descargarlas individualmente.'
            )
        }
      }
    } catch (err) {
      if (montado.current)
        setError(err?.message || 'No se pudieron generar las actas.')
    } finally {
      ocupado.current = false
      if (montado.current) setProcesando(false)
    }
  }
  const generadas = resultados.filter((r) => !r.error).length
  return (
    <div className='app'>
      <div className='franja-pais' aria-hidden='true' />
      <header className='cabecera'>
        <div className='contenedor cabecera-interior'>
          <img
            className='logo'
            src={`${import.meta.env.BASE_URL}logo-mintrabajo.svg`}
            alt='Ministerio del Trabajo de Colombia'
          />
          <div className='cabecera-separador' aria-hidden='true' />
          <div className='identidad'>
            <span>MINISTERIO DEL TRABAJO</span>
            <strong>Gestión de actas</strong>
          </div>
          <span className='etiqueta-cabecera'>Gestión documental</span>
        </div>
      </header>
      <main className='contenedor'>
        <div className='ruta'>
          Gestión documental <span aria-hidden='true'>/</span>
          <span>Actas de reunión</span>
        </div>
        <section className='introduccion'>
          <div>
            <p className='sobre-titulo'>DOCUMENTOS INSTITUCIONALES</p>
            <h1>Generador de actas de reunión</h1>
            <p className='descripcion'>
              Convierte los archivos TXT o Word (.docx) del agente en actas con el
              formato institucional. Carga varios archivos y genéralos en un
              solo paso.
            </p>
          </div>
          <span className='codigo-formato'>
            FORMATO <strong>CMC-F-05</strong>
          </span>
        </section>
        <form
          className='tarjeta carga'
          onSubmit={onGenerar}
          aria-busy={procesando}
        >
          <div className='titulo-seccion'>
            <span className='paso'>01</span>
            <div>
              <h2>Carga tus archivos</h2>
              <p>Un archivo TXT o Word (.docx) por cada acta de reunión.</p>
            </div>
          </div>
          <div
            className={`zona-carga ${arrastrando ? 'arrastrando' : ''}`}
            onDragOver={(e) => {
              e.preventDefault()
              if (!procesando) setArrastrando(true)
            }}
            onDragLeave={() => setArrastrando(false)}
            onDrop={(e) => {
              e.preventDefault()
              setArrastrando(false)
              agregarArchivos(e.dataTransfer.files)
            }}
          >
            <span className='icono-subir'>
              <Icono tipo='subir' width='30' height='30' />
            </span>
            <h3>Todo listo para tu próxima acta</h3>
            <p>Arrastra tus archivos aquí o selecciónalos desde tu equipo.</p>
            <input
              ref={inputRef}
              className='input-archivo'
              id='txt'
              type='file'
              multiple
              accept='.txt,.docx,text/plain,application/vnd.openxmlformats-officedocument.wordprocessingml.document'
              disabled={procesando}
              aria-label='Seleccionar archivos TXT o Word'
              onChange={(e) => {
                agregarArchivos(e.target.files)
                e.target.value = ''
              }}
            />
            <button
              className='boton-secundario'
              type='button'
              disabled={procesando}
              onClick={() => inputRef.current?.click()}
            >
              <Icono tipo='subir' />
              Seleccionar archivos
            </button>
            <small>Formatos admitidos: .txt y .docx · Selección múltiple</small>
            <small>Los documentos deben conservar las secciones del agente (###DATOS_GENERALES###, ###PARTICIPANTES###, etc.).</small>
          </div>
          {archivos.length > 0 && (
            <div className='seleccion'>
              <div className='seleccion-titulo'>
                <strong>
                  {archivos.length}{' '}
                  {archivos.length === 1
                    ? 'archivo seleccionado'
                    : 'archivos seleccionados'}
                </strong>
                <button
                  className='boton-texto'
                  type='button'
                  disabled={procesando}
                  onClick={() => setArchivos([])}
                >
                  Quitar todos
                </button>
              </div>
              <ul className='lista-archivos'>
                {archivos.map((a, i) => (
                  <li key={`${a.name}-${a.lastModified}-${a.size}`}>
                    <span className='icono-archivo'>
                      <Icono tipo='archivo' />
                    </span>
                    <span className='detalle-archivo'>
                      <strong>{a.name}</strong>
                      <small>
                        {new Intl.NumberFormat('es-CO', {
                          maximumFractionDigits: 1
                        }).format(a.size / 1024)}{' '}
                        KB · {a.name.split('.').pop().toUpperCase()}
                      </small>
                    </span>
                    <button
                      className='quitar'
                      type='button'
                      disabled={procesando}
                      aria-label={`Quitar ${a.name}`}
                      onClick={() =>
                        setArchivos((prev) => prev.filter((_, j) => j !== i))
                      }
                    >
                      ×
                    </button>
                  </li>
                ))}
              </ul>
            </div>
          )}
          <div className='pie-carga'>
            <span>La plantilla institucional se aplica automáticamente.</span>
            <button
              className='boton-primario'
              type='submit'
              disabled={procesando}
            >
              <Icono tipo='archivo' />
              {procesando
                ? `Generando ${progreso} de ${archivos.length}…`
                : archivos.length > 1
                  ? 'Generar actas'
                  : 'Generar acta'}
              <span aria-hidden='true'>→</span>
            </button>
          </div>
          <div aria-live='polite' aria-atomic='true'>
            {procesando && (
              <div className='progreso'>
                <progress value={progreso} max={archivos.length} />
                <span>
                  {progreso} de {archivos.length} archivos procesados
                </span>
              </div>
            )}
          </div>
          {error && (
            <div className='mensaje error' role='alert'>
              {error}
            </div>
          )}
        </form>
        {resultados.length > 0 && (
          <section className='tarjeta resultados' aria-label='Resultados'>
            <div className='encabezado-resultados'>
              <div>
                <p className='sobre-titulo'>TUS DOCUMENTOS</p>
                <h2>
                  {procesando ? 'Generando documentos…' : 'Proceso finalizado'}
                </h2>
                <p role='status'>
                  {generadas}{' '}
                  {generadas === 1 ? 'acta generada' : 'actas generadas'} de{' '}
                  {resultados.length}{' '}
                  {resultados.length === 1
                    ? 'archivo procesado'
                    : 'archivos procesados'}
                  .
                </p>
              </div>
              {paquete && (
                <a
                  className='boton-primario'
                  href={paquete.url}
                  download={paquete.nombre}
                >
                  <Icono tipo='descargar' />
                  Descargar todas (.zip)
                </a>
              )}
            </div>
            <ul className='lista-resultados'>
              {resultados.map((r, i) => (
                <li key={i}>
                  <div className='resultado-linea'>
                    <Icono tipo='archivo' />
                    <div className='detalle-archivo'>
                      <strong>{r.origen}</strong>
                      <span className={r.error ? 'estado-error' : 'estado-ok'}>
                        {r.error
                          ? 'No se pudo generar'
                          : r.avisos.length
                            ? 'Generada · Requiere revisión'
                            : 'Generada correctamente'}
                      </span>
                    </div>
                    {!r.error && (
                      <a
                        className='descarga-individual'
                        aria-label={`Descargar ${r.nombre}`}
                        href={r.url}
                        download={r.nombre}
                      >
                        <Icono tipo='descargar' />
                        <span>
                          Descargar <span className='sr-only'>{r.nombre}</span>
                        </span>
                      </a>
                    )}
                  </div>
                  {r.error ? (
                    <p className='mensaje error'>{r.error}</p>
                  ) : (
                    r.avisos.length > 0 && (
                      <details>
                        <summary>{r.avisos.length} avisos para revisar</summary>
                        <ul>
                          {r.avisos.map((a, j) => (
                            <li key={j}>{a}</li>
                          ))}
                        </ul>
                      </details>
                    )
                  )}
                </li>
              ))}
            </ul>
          </section>
        )}
        <div className='privacidad'>
          <Icono tipo='seguro' />
          <p>
            <strong>Tus documentos se quedan contigo.</strong> El procesamiento
            ocurre en tu navegador. Los archivos no se envían a ningún servidor.
          </p>
        </div>
      </main>
      <footer>
        <div className='contenedor'>
          <span>
            Ministerio del Trabajo <span className='separador-footer'>|</span>{' '}
            República de Colombia
          </span>
          <span>Actas de reunión · CMC-F-05</span>
        </div>
      </footer>
    </div>
  )
}
