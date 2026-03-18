# 📚 UTP - GrammarScan

Aplicativo web en Streamlit para descarga masiva, transformación y revisión automatizada de ortografía y gramática en documentos académicos y administrativos.

[![Python](https://img.shields.io/badge/Python-3.8%2B-blue.svg)](https://www.python.org/)
[![Streamlit](https://img.shields.io/badge/Streamlit-App-FF4B4B.svg)](https://streamlit.io/)
[![Estado](https://img.shields.io/badge/Estado-Producción-success.svg)]()

---

## Descripción general

**UTP - GrammarScan** permite descargar archivos desde un Excel con URLs, procesar documentos PDF, Word, PowerPoint, TXT o ZIP, analizar el contenido con **LanguageTool**, detectar **modismos argentinos** de forma opcional y generar un **reporte Excel** con incidencias, sugerencias y contexto.

En términos simples, la herramienta ayuda a responder esta necesidad:

- Recibir muchos documentos desde distintas fuentes.
- Estandarizarlos para poder analizarlos.
- Revisar ortografía y gramática de forma automatizada.
- Entregar un reporte listo para revisión académica o administrativa.
  
<img width="1743" height="835" alt="image" src="https://github.com/user-attachments/assets/b7914ef8-18c4-4047-a582-52e7b17bbe95" />

---

## ¿Qué problema resuelve?

La revisión manual de documentos suele ser lenta, repetitiva y difícil de escalar cuando se trabaja con muchos archivos.

Este aplicativo resuelve ese problema porque:

- Automatiza la descarga de documentos desde un reporte Excel que contiene las URLs de los archivos a descargar.
- Convierte PDFs a Word para facilitar su procesamiento.
- Analiza ortografía y gramática con reglas robustas.
- Filtra bibliografía para reducir falsos positivos.
- Genera un Excel final con resultados detallados.

---

## ¿Para quién está pensado?

Este aplicativo está orientado a personas que necesitan **revisar calidad de redacción** sin intervenir en código, por ejemplo:

- Docentes
- Coordinadores académicos
- Equipos de calidad
- Personal administrativo
- Revisores de contenido

---

## ¿Qué puede hacer el aplicativo?

### Funcionalidades principales

- **Descarga masiva desde Excel** usando una columna `url`.
- **Carga manual** de archivos PDF, DOCX, PPTX, TXT y ZIP.
- **Conversión de PDF a Word** para normalizar documentos antes del análisis.
- **Análisis ortográfico y gramatical** con LanguageTool.
- **Detección opcional de modismos argentinos** mediante `modismos_ar.xlsx`.
- **Filtrado de bibliografía y referencias** para disminuir ruido en los resultados.
- **Exportación a Excel** con detalle de incidencias y resumen consolidado.
- **Descarga automática del Excel final** desde la interfaz.
- **Persistencia de flujo por sesión** usando `st.session_state`.

---

## Flujo funcional del usuario

### Flujo resumido

```text
Excel con URLs o archivos directos
        ↓
Descarga masiva de documentos
        ↓
Conversión de PDFs a Word
        ↓
Preparación de documentos para análisis
        ↓
Validación ortográfica y gramatical
        ↓
Reporte final en Excel
```

### Flujo de uso dentro del aplicativo

1. El usuario puede cargar un **Excel con URLs** de documentos.
2. La app descarga automáticamente todos los archivos compatibles.
3. Si existen PDFs, la app los transforma a **DOCX**.
4. El usuario también puede subir documentos manualmente o en ZIP.
5. El sistema analiza los archivos admitidos.
6. Se muestran métricas e incidencias detectadas.
7. Se genera un **Excel final** con resultados y resumen.
8. El archivo Excel se descarga de forma automatica desde la interfaz.

---

## 🏗️ Arquitectura

La arquitectura del aplicativo se puede entender en **4 capas**:

### 1. Capa de interfaz
Es la pantalla que ve el usuario en Streamlit.

Aquí se muestran:
- El Módulo Home
- El Módulo Report GrammarScan
- Expanders
- Métricas
- Barras de progreso
- Botones de descarga
- Tablas de resultados.

### 2. Capa de procesamiento documental
Es la parte que trabaja directamente con los archivos.

Incluye:
- Descarga masiva desde URLs.
- Lectura de PDF, DOCX, PPTX y TXT.
- Extracción de texto.
- Conversión PDF → DOCX.
- Expansión de ZIP.
- Conteo lógico de páginas.
- Filtros de bibliografía.

### 3. Capa de análisis y salida
Es donde ocurre la validación lingüística y la generación del resultado final.

Incluye:
- LanguageTool.
- Detección de modismos.
- Limpieza de incidencias inválidas.
- Enriquecimiento de metadata.
- Exportación a Excel.

---

## 🧱 Arquitectura técnica

```text
┌──────────────────────────────────────────────┐
│ Interfaz Streamlit                           │
│ - Home                                       │
│ - Report GrammarScan                         │
│ - Sidebar / métricas / progreso / descargas │
└───────────────────────┬──────────────────────┘
                        │
                        ▼
┌──────────────────────────────────────────────┐
│ Orquestación de sesión                       │
│ - init_session_state()                       │
│ - reset_report_broken_pipeline()             │
│ - reset_grammarscan_state()                  │
│ - reset_full_pipeline()                      │
└───────────────────────┬──────────────────────┘
                        │
                        ▼
┌──────────────────────────────────────────────┐
│ Procesamiento documental                     │
│ - _run_descarga_masiva_streamlit()           │
│ - PDFBatchProcessor                          │
│ - expand_uploaded_files()                    │
│ - extract_pages()                            │
│ - read_pdf_pages() / read_docx_pages()       │
│ - read_pptx_slides() / read_txt_pages()      │
└───────────────────────┬──────────────────────┘
                        │
                        ▼
┌──────────────────────────────────────────────┐
│ Análisis lingüístico                         │
│ - get_language_tool()                        │
│ - analyze_text()                             │
│ - analyze_file()                             │
│ - detect_modismos_in_pages()                 │
│ - detect_bibliography_pages()                │
└───────────────────────┬──────────────────────┘
                        │
                        ▼
┌──────────────────────────────────────────────┐
│ Reportería y salida                          │
│ - process_grammarscan_files()                │
│ - _enrich_grammarscan_with_name_linkclass()  │
│ - to_excel_bytes()                           │
│ - descarga manual y automática del Excel     │
└──────────────────────────────────────────────┘
```

---

## Pipeline operativo real

### Etapa A. Descarga masiva desde Excel
Entrada esperada:
- Un Excel con la columna `url`.

La app:
- Filtra extensiones permitidas.
- Valida límites en Streamlit Cloud.
- Descarga documentos con reintentos.
- Genera ZIP de archivos descargados.
- Genera CSV de descargas fallidas cuando aplica.
  
<img width="1453" height="806" alt="image" src="https://github.com/user-attachments/assets/7687c6d0-ef27-4ba9-bec9-48b43f625a78" />

### Etapa B. Carga directa de documentos
El usuario puede subir:
- PDF
- DOCX
- PPTX
- ZIP con documentos dentro.

Los formatos `.doc` y `.ppt` se identifican, pero se marcan como no soportados para análisis directo.

<img width="1646" height="843" alt="image" src="https://github.com/user-attachments/assets/ed3bb6e2-33ff-4c14-a2a7-a2d82747e0fb" />

### Etapa C. Transformación PDF → Word
Los PDFs pueden convertirse a DOCX para integrarlos al flujo posterior.

### Etapa D. Análisis GrammarScan
Se procesan los archivos válidos con:
- Extracción de texto.
- Fltrado de bibliografía.
- Control de URLs dentro del texto.
- Revisión con LanguageTool.
- Detección opcional de modismos.

<img width="1648" height="815" alt="image" src="https://github.com/user-attachments/assets/8177e7e0-5f66-44e0-9155-10d2c0a29d14" />

### Etapa E. Exportación
La aplicación genera:
- **Resultados**: detalle de incidencias,
- **ResumenIncidencias**,
- **ResumenCompleto**.

---

## Formatos soportados

### Entrada

| Tipo | Extensión | Estado |
|------|-----------|--------|
| PDF | `.pdf` | Soportado |
| Word | `.docx` | Soportado |
| PowerPoint | `.pptx` | Soportado |
| Texto | `.txt` | Soportado |
| ZIP | `.zip` | Soportado |
| Word legado | `.doc` | Detectado, no recomendado |
| PowerPoint legado | `.ppt` | Detectado, no recomendado |

### Salida

| Tipo | Archivo |
|------|---------|
| Excel final | `UTP_GrammarScan_Resultados.xlsx` |

---

## Componentes técnicos clave

### `init_session_state()`
Inicializa el estado persistente del flujo para que la app recuerde resultados, cargas y pasos ya ejecutados dentro de la sesión.

### `PDFBatchProcessor`
Clase encargada de procesar PDFs y convertir su contenido a DOCX.

### `process_grammarscan_files()`
Función orquestadora del análisis principal. Recorre archivos lógicos, aplica reglas, arma métricas y construye los DataFrames finales.

### `analyze_file()`
Analiza un archivo individual. Extrae páginas, aplica filtros, ejecuta LanguageTool y arma incidencias por archivo.

### `detect_bibliography_pages()`
Reduce falsos positivos detectando páginas o fragmentos que parecen bibliografía o referencias.

### `to_excel_bytes()`
Convierte los resultados a un Excel descargable con hojas de resultados y resumen.

---

## Archivos importantes del proyecto

```text
app.py                              # Aplicación principal Streamlit.
modismos_ar.xlsx                    # Diccionario de modismos argentinos.
custom_dictionary_ignore_list.txt   # Lista de términos a ignorar para reglas morfológicas.
requirements.txt                    # Dependencias del proyecto.
.streamlit/                         # Configuración para despliegue en Streamlit.
```

---

## Requisitos para ejecución

### Requisitos base

- Python 3.8 o superior.
- Java instalado y accesible en el sistema.
- Dependencias Python del proyecto (requirements.txt)

### Dependencias principales usadas por la app

- **streamlit**: framework para construir la interfaz web interactiva del aplicativo.
- **pandas**: librería para leer, transformar y exportar datos en estructuras tipo tabla.
- **pdfplumber**: usada para extraer texto y tablas desde archivos PDF.
- **PyMuPDF (`fitz`)**: permite leer, analizar y manipular PDFs de forma rápida y eficiente.
- **python-docx**: utilizada para crear y editar documentos de Word (`.docx`).
- **python-pptx**: permite generar y modificar presentaciones de PowerPoint (`.pptx`).
- **openpyxl**: permite leer y editar archivos Excel existentes.
- **lxml**: facilita el procesamiento y análisis de estructuras XML y HTML.
- **xlsxwriter**: se utiliza para generar archivos Excel con formato.
- **requests**: librería para realizar consultas HTTP a servicios externos o APIs.
- **language_tool_python**: usada para revisar gramática, ortografía y estilo del texto automáticamente.

---

## Recomendaciones de uso

### Para usuarios funcionales

- Usa un Excel limpio con columna `url`.
- Si el lote es grande, trabaja por bloques.
- Usa ZIP cuando tengas muchos archivos manuales.
- Activa la exclusión de bibliografía para obtener resultados más limpios.
- Activa modismos argentinos solo cuando realmente lo necesites.

### Para despliegue en Streamlit Cloud

- Se recomienda trabajar entre **500 y 700 registros** por lote.
- Para cargas mayores, conviene dividir el Excel o ejecutar localmente el app.
- Esta recomendación ayuda a evitar problemas de memoria del contenedor.

---

## Consideraciones y límites actuales

- Los formatos `.doc` y `.ppt` no son ideales para análisis directo.
- Los documentos de más de **100 páginas o diapositivas** no se analizan con LanguageTool y se reportan en resumen como no analizados.
- La calidad del análisis depende de que el texto del documento pueda extraerse correctamente.
- Streamlit Cloud puede quedarse corto para lotes muy grandes.

---

## Solución rápida de problemas

### Java no detectado
Verifica que Java esté instalado y disponible en el `PATH`.

### No se analizan documentos grandes
Revisa si el archivo supera las 100 páginas o diapositivas, este tipo de archivos no son tratados como documentos académicos.

### El Excel que contiene las URLs no carga
Confirma que exista la columna `url` y que el archivo sea `.xlsx` o `.xls`.

### No se descargan todos los documentos
Descarga el CSV de fallidos para revisar qué URLs no pudieron procesarse.

### No aparecen incidencias
Puede suceder si:
- El archivo no tiene texto extraíble.
- El contenido fue filtrado como bibliografía.
- El documento no contiene incidencias detectables.
- El archivo excede el umbral de páginas soportado para análisis (archivos > 100 páginas)

---

## 👨‍💻 Autor

**José Luis Antunez Condezo**

Proyecto orientado a automatización documental, análisis lingüístico y flujos de revisión académica.
