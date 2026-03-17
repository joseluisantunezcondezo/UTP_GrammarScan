# 📚 UTP - GrammarScan

> Aplicativo web en Streamlit para descarga masiva, transformación y revisión automatizada de ortografía y gramática en documentos académicos y administrativos.

[![Python](https://img.shields.io/badge/Python-3.8%2B-blue.svg)](https://www.python.org/)
[![Streamlit](https://img.shields.io/badge/Streamlit-App-FF4B4B.svg)](https://streamlit.io/)
[![Estado](https://img.shields.io/badge/Estado-Producción-success.svg)]()

---

## 📌 Descripción general

**UTP - GrammarScan** es una aplicación pensada para simplificar la revisión documental en entornos académicos.
Permite descargar archivos desde un Excel con URLs, procesar documentos PDF, Word, PowerPoint, TXT o ZIP, analizar el contenido con **LanguageTool local**, detectar **modismos argentinos** de forma opcional y generar un **reporte Excel** con incidencias, sugerencias y contexto.

En términos simples, la herramienta ayuda a responder esta necesidad:

- Recibir muchos documentos desde distintas fuentes.
- Estandarizarlos para poder analizarlos.
- Revisar ortografía y gramática de forma automatizada.
- Entregar un reporte listo para revisión académica o administrativa.

---

## 🎯 ¿Qué problema resuelve?

La revisión manual de documentos suele ser lenta, repetitiva y difícil de escalar cuando se trabaja con muchos archivos.

Este aplicativo resuelve ese problema porque:

- automatiza la descarga de documentos desde un Excel,
- convierte PDFs a Word para facilitar su procesamiento,
- analiza ortografía y gramática con reglas robustas,
- filtra bibliografía para reducir falsos positivos,
- genera un Excel final con resultados detallados.

---

## 👥 ¿Para quién está pensado?

### Usuarios no técnicos

Este aplicativo está orientado a personas que necesitan **revisar calidad de redacción** sin intervenir en código, por ejemplo:

- docentes,
- coordinadores académicos,
- equipos de calidad,
- personal administrativo,
- revisores de contenido.

### Usuarios técnicos

También está pensado para quienes necesiten **mantener, desplegar o mejorar** la solución:

- desarrolladores Python,
- analistas de automatización,
- responsables de despliegue en Streamlit,
- equipos de soporte técnico.

---

## ✅ ¿Qué puede hacer el aplicativo?

### Funcionalidades principales

- **Descarga masiva desde Excel** usando una columna `url`.
- **Carga manual** de archivos PDF, DOCX, PPTX, TXT y ZIP.
- **Conversión de PDF a Word** para normalizar documentos antes del análisis.
- **Análisis ortográfico y gramatical** con LanguageTool local.
- **Detección opcional de modismos argentinos** mediante `modismos_ar.xlsx`.
- **Filtrado de bibliografía y referencias** para disminuir ruido en los resultados.
- **Exportación a Excel** con detalle de incidencias y resumen consolidado.
- **Descarga automática del Excel final** desde la interfaz.
- **Persistencia de flujo por sesión** usando `st.session_state`.

---

## 🧭 Flujo funcional del usuario

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
2. La app descarga automáticamente archivos compatibles.
3. Si existen PDFs, la app los transforma a **DOCX**.
4. El usuario también puede subir documentos manualmente o en ZIP.
5. El sistema analiza los archivos admitidos.
6. Se muestran métricas e incidencias detectadas.
7. Se genera un **Excel final** con resultados y resumen.
8. El archivo Excel se descarga desde la interfaz.

---

## 🏗️ Arquitectura explicada de forma sencilla

La arquitectura del aplicativo se puede entender en **4 capas**:

### 1. Capa de interfaz
Es la pantalla que ve el usuario en Streamlit.

Aquí se muestran:
- el módulo Home,
- el módulo Report GrammarScan,
- expanders,
- métricas,
- barras de progreso,
- botones de descarga,
- tablas de resultados.

### 2. Capa de orquestación
Es la lógica que coordina el flujo completo usando `st.session_state`.

Esta capa decide:
- qué parte del proceso ya fue ejecutada,
- cuándo reiniciar el pipeline,
- cuándo reutilizar resultados previos,
- cuándo relanzar análisis automáticamente.

### 3. Capa de procesamiento documental
Es la parte que trabaja directamente con los archivos.

Incluye:
- descarga masiva desde URLs,
- lectura de PDF, DOCX, PPTX y TXT,
- extracción de texto,
- conversión PDF → DOCX,
- expansión de ZIP,
- conteo lógico de páginas,
- filtros de bibliografía.

### 4. Capa de análisis y salida
Es donde ocurre la validación lingüística y la generación del resultado final.

Incluye:
- LanguageTool local,
- detección de modismos,
- limpieza de incidencias inválidas,
- enriquecimiento de metadata,
- exportación a Excel.

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

## 🧩 Módulos del aplicativo

## 1. Home
Muestra el propósito del sistema, sus funcionalidades, el flujo de trabajo y recomendaciones generales de uso.

## 2. Report GrammarScan
Concentra el flujo operativo principal:

- carga del Excel con URLs,
- descarga masiva,
- descarga de ZIPs consolidados,
- carga directa de documentos,
- conversión de PDF a Word,
- análisis ortográfico y gramatical,
- exportación automática a Excel.

---

## 🔄 Pipeline operativo real

### Etapa A. Descarga masiva desde Excel
Entrada esperada:
- un Excel con la columna `url`.

La app:
- filtra extensiones permitidas,
- valida límites en Streamlit Cloud,
- descarga documentos con reintentos,
- genera ZIP de archivos descargados,
- genera CSV de descargas fallidas cuando aplica.

### Etapa B. Carga directa de documentos
El usuario puede subir:
- PDF,
- DOCX,
- PPTX,
- ZIP con documentos dentro.

Los formatos `.doc` y `.ppt` se identifican, pero se marcan como no soportados para análisis directo.

### Etapa C. Transformación PDF → Word
Los PDFs pueden convertirse a DOCX para integrarlos al flujo posterior.

### Etapa D. Análisis GrammarScan
Se procesan los archivos válidos con:
- extracción de texto,
- filtrado de bibliografía,
- control de URLs dentro del texto,
- revisión con LanguageTool,
- detección opcional de modismos.

### Etapa E. Exportación
La aplicación genera:
- **Resultados**: detalle de incidencias,
- **ResumenIncidencias**,
- **ResumenCompleto**.

---

## 📂 Formatos soportados

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
| ZIP de descarga | `Descarga_Masiva_Documentos_*.zip` |
| CSV de fallidos | `descargas_fallidas.csv` |

---

## 🧠 Componentes técnicos clave

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

## 🗂️ Archivos importantes del proyecto

```text
app.py                              # Aplicación principal Streamlit
modismos_ar.xlsx                    # Diccionario de modismos argentinos
custom_dictionary_ignore_list.txt   # Lista de términos a ignorar para reglas morfológicas
requirements.txt                    # Dependencias del proyecto
.streamlit/                         # Configuración para despliegue en Streamlit
```

---

## ⚙️ Requisitos para ejecución

### Requisitos base

- Python 3.8 o superior
- Java instalado y accesible en el sistema
- Dependencias Python del proyecto

### Dependencias principales usadas por la app

- `streamlit`
- `pandas`
- `pdfplumber`
- `PyMuPDF` (`fitz`)
- `python-docx`
- `python-pptx`
- `requests`
- `language_tool_python`

---

## 🚀 Ejecución local

```bash
python -m venv venv
```

### Windows

```bash
venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

### Linux / macOS

```bash
source venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

---

## 🖥️ Recomendaciones de uso

### Para usuarios funcionales

- Usa un Excel limpio con columna `url`.
- Si el lote es grande, trabaja por bloques.
- Usa ZIP cuando tengas muchos archivos manuales.
- Activa la exclusión de bibliografía para obtener resultados más limpios.
- Activa modismos argentinos solo cuando realmente lo necesites.

### Para despliegue en Streamlit Cloud

- Se recomienda trabajar entre **500 y 700 registros** por lote.
- Para cargas mayores, conviene dividir el Excel o ejecutar localmente.
- Esta recomendación ayuda a evitar problemas de memoria del contenedor.

### Para mantenimiento técnico

- Mantén `modismos_ar.xlsx` versionado.
- Mantén actualizada la lista `custom_dictionary_ignore_list.txt`.
- Evita mezclar lógica de UI con lógica de negocio en nuevas mejoras.
- Conserva el uso de `session_state` para no romper el flujo persistente.

---

## ⚠️ Consideraciones y límites actuales

- Los formatos `.doc` y `.ppt` no son ideales para análisis directo.
- Los documentos de más de **100 páginas o diapositivas** no se analizan con LanguageTool y se reportan en resumen como no analizados.
- La calidad del análisis depende de que el texto del documento pueda extraerse correctamente.
- PDFs escaneados pueden requerir OCR fuera de este flujo.
- Streamlit Cloud puede quedarse corto para lotes muy grandes.

---

## 🔐 Seguridad y manejo de archivos

El aplicativo está orientado a procesamiento controlado por sesión.

Buenas prácticas recomendadas:
- no conservar archivos temporales más tiempo del necesario,
- no subir información sensible a entornos públicos sin validación previa,
- ejecutar localmente cuando el volumen o la sensibilidad documental sea alto.

---

## 🛠️ Solución rápida de problemas

### Java no detectado
Verifica que Java esté instalado y disponible en el `PATH`.

### No se analizan documentos grandes
Revisa si el archivo supera las 100 páginas o diapositivas.

### El Excel no carga
Confirma que exista la columna `url` y que el archivo sea `.xlsx` o `.xls`.

### No se descargan todos los documentos
Descarga el CSV de fallidos para revisar qué URLs no pudieron procesarse.

### No aparecen incidencias
Puede suceder si:
- el archivo no tiene texto extraíble,
- el contenido fue filtrado como bibliografía,
- el documento no contiene incidencias detectables,
- el archivo excede el umbral de páginas soportado para análisis.

---

## 🧪 Recomendaciones de documentación para GitHub

Si vas a mantener este proyecto en producción, el README debería conservar siempre estas secciones:

1. **Qué hace la herramienta** en lenguaje simple.
2. **Flujo funcional** para usuarios no técnicos.
3. **Arquitectura técnica** para desarrolladores.
4. **Requisitos y ejecución** local o cloud.
5. **Formatos soportados y límites**.
6. **Archivos clave del proyecto**.
7. **Problemas frecuentes y solución**.
8. **Próximos cambios o roadmap** solo cuando estén confirmados.

---

## 📈 Próximas mejoras sugeridas

- incorporar OCR para PDFs escaneados,
- separar módulos de UI, procesamiento y exportación en archivos independientes,
- agregar pruebas automáticas para las funciones críticas,
- documentar mejor los archivos auxiliares de diccionario y exclusiones,
- incorporar un diagrama de arquitectura visual en `/docs`.

---

## 📄 Licencia

Define aquí la licencia que aplicará al proyecto, por ejemplo MIT, Apache 2.0 o uso interno institucional.

---

## 👨‍💻 Autor

**José Luis Antunez Condezo**

Proyecto orientado a automatización documental, análisis lingüístico y flujos de revisión académica.
