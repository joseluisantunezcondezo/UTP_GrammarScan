# 📚 UTP - GrammarScan

> **Sistema inteligente de análisis gramatical y procesamiento masivo de documentos académicos**

[![Python](https://img.shields.io/badge/Python-3.8%2B-blue.svg)](https://www.python.org/)
[![Streamlit](https://img.shields.io/badge/Streamlit-1.28%2B-FF4B4B.svg)](https://streamlit.io/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Status](https://img.shields.io/badge/Status-Production-success.svg)]()

---

## 📖 Índice

- [¿Qué es UTP - GrammarScan?](#-qué-es-utp---grammarscan)
- [Características Principales](#-características-principales)
- [Casos de Uso](#-casos-de-uso)
- [Arquitectura del Sistema](#-arquitectura-del-sistema)
- [Instalación](#-instalación)
- [Configuración](#-configuración)
- [Guía de Uso](#-guía-de-uso)
- [Formatos Soportados](#-formatos-soportados)
- [API y Componentes](#-api-y-componentes)
- [Mejores Prácticas](#-mejores-prácticas)
- [Solución de Problemas](#-solución-de-problemas)
- [Contribuir](#-contribuir)
- [Roadmap](#-roadmap)
- [Licencia](#-licencia)

---

## 🎯 ¿Qué es UTP - GrammarScan?

**GrammarScan** es una aplicación web diseñada para la **Universidad Tecnológica del Perú (UTP)** que automatiza el análisis gramatical y la gestión de documentos académicos. El sistema permite procesar grandes volúmenes de documentos, detectar errores gramaticales, identificar modismos regionales y generar reportes detallados en formato Excel.

### 🌟 Problema que Resuelve

La revisión manual de documentos académicos es:
- ⏰ **Lenta**: Requiere horas de trabajo manual
- 🔍 **Inconsistente**: Depende de la experiencia del revisor
- 📊 **No escalable**: Difícil de procesar grandes volúmenes
- 📝 **Sin trazabilidad**: No genera reportes automáticos

**GrammarScan soluciona estos problemas mediante:**
- ✅ Análisis automatizado con LanguageTool
- ✅ Procesamiento paralelo de múltiples documentos
- ✅ Detección de modismos regionales personalizables
- ✅ Reportes Excel con trazabilidad completa
- ✅ Pipeline de descarga y análisis integrado

---

## ⚡ Características Principales

### 🔍 Análisis Gramatical Avanzado

- **Motor de análisis**: Integración con [LanguageTool](https://languagetool.org/)
- **Idiomas soportados**: Español, Inglés, y más
- **Tipos de detección**:
  - Errores ortográficos
  - Errores gramaticales
  - Sugerencias de estilo
  - Detección de modismos regionales (configurable)
  
### 📥 Descarga Masiva Inteligente

- **Desde Excel**: Procesa URLs desde archivos Excel
- **Reintentos automáticos**: Hasta 7 intentos con backoff exponencial
- **Manejo de errores**: Continúa procesando incluso si algunas descargas fallan
- **Formatos soportados**: PDF, DOCX, PPTX, DOC, PPT
- **Límite Cloud**: 700 URLs (configurable para ejecución local)

### 🔄 Pipeline Automatizado

```
Excel con URLs → Descarga Masiva → Conversión → Análisis → Reporte Excel
```

1. **Entrada**: Excel con columna de URLs
2. **Descarga**: Sistema paralelo con reintentos
3. **Conversión**: Documentos a formatos analizables
4. **Análisis**: LanguageTool + Detección de modismos
5. **Salida**: Excel con resultados detallados

### 📊 Reportes Detallados

Genera archivos Excel con dos hojas:
- **Resultados**: Incidencias por documento con contexto
- **Resumen**: Estadísticas agregadas por archivo

### 🎨 Interfaz Moderna

- Diseño responsive con Streamlit
- Cards visuales con estados de progreso
- Descarga automática de resultados
- Feedback en tiempo real

---

## 💼 Casos de Uso

### 1️⃣ Revisión de Trabajos Académicos

**Escenario**: Docente necesita revisar 50 ensayos de estudiantes

**Proceso**:
1. Recopila los ensayos en formato PDF/DOCX
2. Sube los archivos al sistema (puede subir ZIP)
3. Configura idioma español y activa detección de modismos
4. Obtiene reporte Excel con todas las incidencias
5. Filtra por tipo de error para dar feedback específico

**Beneficio**: Ahorra 80% del tiempo de revisión manual

### 2️⃣ Auditoría de Contenido Educativo

**Escenario**: Institución necesita auditar 500 documentos almacenados en URLs

**Proceso**:
1. Exporta lista de URLs desde su sistema a Excel
2. Carga el Excel en el módulo de descarga masiva
3. Sistema descarga y analiza automáticamente
4. Obtiene reporte consolidado con métricas de calidad

**Beneficio**: Centraliza auditoría de todo el contenido

### 3️⃣ Estandarización de Lenguaje

**Escenario**: Universidad quiere eliminar modismos regionales de material didáctico

**Proceso**:
1. Configura archivo `modismos_ar.xlsx` con expresiones a detectar
2. Procesa documentos con detección de modismos activada
3. Revisa sugerencias de reemplazo en el reporte
4. Actualiza documentos según estándares institucionales

**Beneficio**: Mantiene consistencia en todo el material académico

---

## 🏗️ Arquitectura del Sistema

### Vista de Alto Nivel

```
┌─────────────────────────────────────────────────────────────┐
│                     INTERFAZ WEB (Streamlit)                │
│  ┌──────────────┐  ┌──────────────┐  ┌──────────────┐      │
│  │    Home      │  │  GrammarScan │  │   Reportes   │      │
│  └──────────────┘  └──────────────┘  └──────────────┘      │
└───────────────────────────┬─────────────────────────────────┘
                            │
                            ▼
┌─────────────────────────────────────────────────────────────┐
│                    CAPA DE PROCESAMIENTO                     │
│  ┌──────────────────────────────────────────────────────┐   │
│  │  Pipeline de Descarga Masiva                         │   │
│  │  • ThreadPoolExecutor para descargas paralelas       │   │
│  │  • Sistema de reintentos con backoff exponencial     │   │
│  │  • Validación de extensiones y tamaños               │   │
│  └──────────────────────────────────────────────────────┘   │
│  ┌──────────────────────────────────────────────────────┐   │
│  │  Extracción de Texto                                 │   │
│  │  • PDF: pdfplumber + PyMuPDF (fitz)                 │   │
│  │  • DOCX: python-docx                                 │   │
│  │  • PPTX: python-pptx                                 │   │
│  │  • TXT: lectura directa                              │   │
│  └──────────────────────────────────────────────────────┘   │
│  ┌──────────────────────────────────────────────────────┐   │
│  │  Motor de Análisis                                   │   │
│  │  • LanguageTool (análisis gramatical)                │   │
│  │  • Sistema de detección de modismos (regex patterns) │   │
│  │  • Normalización de texto                            │   │
│  │  • Paginación virtual (50 líneas/página)             │   │
│  └──────────────────────────────────────────────────────┘   │
└───────────────────────────┬─────────────────────────────────┘
                            │
                            ▼
┌─────────────────────────────────────────────────────────────┐
│                    CAPA DE DATOS                             │
│  ┌──────────────┐  ┌──────────────┐  ┌──────────────┐      │
│  │  DataFrames  │  │ Excel Output │  │  ZIP Temp    │      │
│  │  (Pandas)    │  │ (openpyxl)   │  │  Storage     │      │
│  └──────────────┘  └──────────────┘  └──────────────┘      │
└─────────────────────────────────────────────────────────────┘
```

### Componentes Principales

#### 1. **GrammarScan Core** (`process_grammarscan_files`)
- **Responsabilidad**: Orquestar todo el proceso de análisis
- **Entrada**: Lista de archivos (UploadedFile o LogicalFileSource)
- **Salida**: DataFrames con resultados y resúmenes
- **Características**:
  - Procesamiento paralelo con ThreadPoolExecutor
  - Manejo robusto de errores por archivo
  - Métricas de procesamiento en tiempo real

#### 2. **Text Extractor** (`extract_text_from_...`)
- **Responsabilidad**: Extraer texto de diferentes formatos
- **Métodos especializados**:
  - `extract_text_from_pdf()`: PDF → texto plano
  - `extract_text_from_docx()`: DOCX → texto + tablas
  - `extract_text_from_pptx()`: PPTX → texto de slides
  - `extract_text_from_txt()`: TXT → texto directo
- **Características**:
  - Normalización de espacios en blanco
  - Manejo de caracteres especiales
  - Extracción de tablas (DOCX)

#### 3. **LanguageTool Client** (`call_languagetool`)
- **Responsabilidad**: Comunicación con API de LanguageTool
- **Características**:
  - División automática de textos largos (30,000 caracteres)
  - Manejo de errores de red
  - Soporte multi-idioma
- **Endpoint**: Configurable (local o cloud)

#### 4. **Modismos Detector** (`detect_modismos`)
- **Responsabilidad**: Detectar expresiones regionales
- **Fuente**: Archivo Excel `modismos_ar.xlsx`
- **Tipos de patrones**:
  - Literales: Búsqueda exacta
  - Regex: Patrones complejos
- **Salida**: Lista de coincidencias con contexto

#### 5. **Bulk Downloader** (`download_urls_from_excel`)
- **Responsabilidad**: Descarga masiva paralela
- **Características**:
  - Hasta 7 reintentos por URL
  - Backoff exponencial
  - Validación de extensiones
  - Límite de tamaño (200 MB)
  - Progreso en tiempo real

#### 6. **Excel Reporter** (`to_excel_bytes`)
- **Responsabilidad**: Generar reportes Excel
- **Estructura**:
  - Hoja "Resultados": Detalle de incidencias
  - Hoja "Resumen": Estadísticas por archivo
- **Formato**: XLSX con estilos y filtros

---

## 🚀 Instalación

### Prerrequisitos

- **Python**: 3.8 o superior
- **Java**: 8 o superior (requerido por LanguageTool)
- **Memoria RAM**: Mínimo 2 GB (4 GB recomendado)
- **Espacio en disco**: 500 MB para dependencias

### Instalación Local

#### 1. Clonar el repositorio

```bash
git clone https://github.com/joseluisantunezcondezo/UTP_GrammarScan.git
cd UTP_GrammarScan
```

#### 2. Crear entorno virtual

```bash
# Windows
python -m venv venv
venv\Scripts\activate

# Linux/Mac
python3 -m venv venv
source venv/bin/activate
```

#### 3. Instalar dependencias

```bash
pip install -r requirements.txt
```

#### 4. Verificar instalación de Java

```bash
java -version
```

Si Java no está instalado:
- **Windows**: Descargar desde [java.com](https://www.java.com/)
- **Linux**: `sudo apt-get install default-jre`
- **Mac**: `brew install java`

#### 5. Ejecutar la aplicación

```bash
streamlit run app.py
```

La aplicación se abrirá automáticamente en `http://localhost:8501`

### Instalación con Docker (Opcional)

```bash
# Construir imagen
docker build -t utp-grammarscan .

# Ejecutar contenedor
docker run -p 8501:8501 utp-grammarscan
```

---

## ⚙️ Configuración

### Archivo `modismos_ar.xlsx`

Este archivo configura los modismos a detectar. Debe contener las siguientes columnas:

| Columna | Tipo | Descripción | Ejemplo |
|---------|------|-------------|---------|
| `modismo` | Texto | Expresión a detectar | "che" |
| `tipo` | literal/regex | Tipo de patrón | "literal" |
| `patron` | Texto | Patrón regex (opcional) | `\bche\b` |
| `sugerencia` | Texto | Reemplazo sugerido | "oye" |
| `comentario` | Texto | Explicación | "Modismo argentino" |

#### Ejemplo de contenido:

```
modismo,tipo,patron,sugerencia,comentario
che,literal,,oye,Modismo argentino informal
boludo,literal,,compañero,Modismo argentino coloquial
vos,regex,\bvos\b,tú,Pronombre argentino
```

### Variables de Entorno

#### Para Streamlit Cloud

Crear archivo `.streamlit/secrets.toml`:

```toml
[general]
is_streamlit_cloud = true
max_bulk_urls = 700

[languagetool]
api_url = "https://api.languagetool.org/v2/check"
# O usar instancia local:
# api_url = "http://localhost:8081/v2/check"
```

#### Para ejecución local

Crear archivo `.env`:

```bash
IS_STREAMLIT_CLOUD=false
MAX_BULK_URLS=2000
LANGUAGETOOL_API_URL=http://localhost:8081/v2/check
```

### LanguageTool Local (Recomendado para producción)

Para mejor rendimiento y privacidad, instalar LanguageTool local:

```bash
# Descargar LanguageTool
wget https://languagetool.org/download/LanguageTool-stable.zip
unzip LanguageTool-stable.zip

# Iniciar servidor
cd LanguageTool-*/
java -cp languagetool-server.jar org.languagetool.server.HTTPServer --port 8081
```

Actualizar en el código (línea ~300):

```python
LT_API_URL = "http://localhost:8081/v2/check"
```

---

## 📘 Guía de Uso

### Módulo 1: Home

Pantalla de bienvenida con información general del sistema.

### Módulo 2: Report GrammarScan

#### Flujo Completo (con descarga masiva)

**Paso 1: Configurar descarga masiva**

1. Prepare un archivo Excel con columna `link` o `url`
2. Cada fila debe contener una URL válida a documento
3. Suba el Excel en el Paso 1

**Paso 2: Configurar parámetros de descarga**

- **Workers paralelos**: 4-8 (más rápido pero consume más recursos)
- **Reintentos**: 7 (por defecto)
- **Timeout**: 30 segundos

**Paso 3: Ejecutar descarga**

- Haga clic en "Descargar archivos"
- Observe el progreso en tiempo real
- Descargue el ZIP generado

**Paso 4: Configurar análisis gramatical**

- **Idioma**: Español (es) o Inglés (en)
- **Máx. caracteres**: 30,000 por llamada
- **Workers**: 4 (recomendado)
- **Excluir bibliografía**: Activado (recomendado)
- **Analizar modismos**: Según necesidad

**Paso 5: Subir documentos**

- Los documentos descargados aparecen automáticamente
- También puede subir archivos manualmente
- Soporta archivos individuales o ZIP

**Paso 6: Procesar documentos**

- El análisis inicia automáticamente al subir archivos
- Visualice métricas en tiempo real:
  - Total de documentos
  - Con incidencias
  - Sin incidencias
  - Errores

**Paso 7: Descargar reporte Excel**

- Se genera automáticamente al finalizar
- Contiene 2 hojas:
  - **Resultados**: Todas las incidencias detectadas
  - **Resumen**: Estadísticas por archivo

#### Flujo Rápido (solo análisis)

1. Vaya directo al Paso 6
2. Suba sus archivos (PDF, DOCX, PPTX, TXT, ZIP)
3. Configure parámetros de análisis
4. Obtenga resultados y descargue Excel

---

## 📄 Formatos Soportados

### Entrada

| Formato | Extensión | Notas |
|---------|-----------|-------|
| PDF | `.pdf` | Extrae texto y tablas |
| Word | `.docx` | Extrae texto, tablas y estilos |
| PowerPoint | `.pptx` | Extrae texto de slides |
| Texto plano | `.txt` | Lectura directa |
| ZIP | `.zip` | Descomprime y procesa contenido |
| Word Legacy | `.doc` | A través de conversión |
| PowerPoint Legacy | `.ppt` | A través de conversión |

### Salida

| Formato | Nombre de archivo | Contenido |
|---------|-------------------|-----------|
| Excel | `UTP_GrammarScan_Resultados.xlsx` | Resultados + Resumen |
| ZIP | `bulk_download_TIMESTAMP.zip` | Archivos descargados |

---

## 🔧 API y Componentes

### Funciones Principales

#### `process_grammarscan_files()`

Procesa múltiples archivos y genera análisis gramatical.

**Parámetros**:
```python
def process_grammarscan_files(
    ups: List[UploadedFile],
    lang_code: str = "es",
    max_chars_call: int = 30000,
    workers: int = 4,
    excluir_biblio: bool = True,
    analizar_modismos: bool = False
) -> Tuple[pd.DataFrame, pd.DataFrame, Dict, float]:
```

**Retorna**:
- `final_df`: DataFrame con incidencias detalladas
- `resumen_df`: DataFrame con estadísticas por archivo
- `metrics`: Diccionario con métricas de procesamiento
- `elapsed`: Tiempo total de procesamiento

**Ejemplo de uso**:
```python
results, summary, metrics, time = process_grammarscan_files(
    ups=uploaded_files,
    lang_code="es",
    max_chars_call=30000,
    workers=4,
    excluir_biblio=True,
    analizar_modismos=True
)
```

#### `download_urls_from_excel()`

Descarga archivos desde URLs en Excel.

**Parámetros**:
```python
def download_urls_from_excel(
    excel_bytes: bytes,
    max_workers: int = 4,
    max_retries: int = 7,
    progress_callback: Optional[Callable] = None
) -> Tuple[List[str], Dict]:
```

**Retorna**:
- Lista de rutas de archivos descargados
- Diccionario con estadísticas de descarga

#### `call_languagetool()`

Realiza llamada a LanguageTool API.

**Parámetros**:
```python
def call_languagetool(
    text: str,
    lang_code: str = "es",
    max_chars: int = 30000
) -> List[Dict]:
```

**Retorna**:
- Lista de incidencias detectadas

### Estructuras de Datos

#### ModismoPattern (DataClass)

```python
@dataclass
class ModismoPattern:
    modismo: str          # Expresión a buscar
    tipo: str             # "literal" o "regex"
    patron: str           # Patrón regex
    sugerencia: str       # Reemplazo sugerido
    comentario: str       # Explicación
    regex: re.Pattern     # Patrón compilado
```

#### LogicalFileSource (DataClass)

```python
@dataclass
class LogicalFileSource:
    display_name: str                    # Nombre para mostrar
    ext: str                             # Extensión (.pdf, .docx, etc.)
    read_bytes: Callable[[], bytes]      # Función para leer contenido
```

---

## ✅ Mejores Prácticas

### Para Usuarios

#### Descarga Masiva

1. **Lotes pequeños**: Procese máximo 500-700 URLs por ejecución
2. **Excel limpio**: Asegúrese de que la columna de URLs no tenga celdas vacías
3. **URLs válidas**: Verifique que las URLs sean accesibles
4. **Formato correcto**: Use extensiones permitidas (.pdf, .docx, .pptx)

#### Análisis Gramatical

1. **Idioma correcto**: Seleccione el idioma del documento
2. **Excluir bibliografía**: Active para ignorar referencias
3. **Modismos**: Solo active si necesita detectarlos (consume más tiempo)
4. **Workers**: Use 4 para balance entre velocidad y recursos

### Para Desarrolladores

#### Código Limpio

```python
# ✅ CORRECTO: Usar funciones auxiliares
def process_document(file_source: LogicalFileSource) -> pd.DataFrame:
    text = extract_text(file_source)
    results = analyze_text(text)
    return format_results(results)

# ❌ INCORRECTO: Código monolítico
def process_document(file_source):
    # 500 líneas de código mezclando extracción, análisis y formato
    ...
```

#### Manejo de Errores

```python
# ✅ CORRECTO: Manejo específico por tipo de error
try:
    results = call_languagetool(text)
except requests.Timeout:
    logger.warning(f"Timeout para {file_name}")
    results = []
except Exception as e:
    logger.error(f"Error inesperado: {e}")
    raise

# ❌ INCORRECTO: Captura genérica que oculta problemas
try:
    results = call_languagetool(text)
except:
    pass
```

#### Performance

```python
# ✅ CORRECTO: Procesamiento paralelo
with ThreadPoolExecutor(max_workers=4) as executor:
    futures = {executor.submit(process_doc, doc): doc for doc in docs}
    results = [future.result() for future in as_completed(futures)]

# ❌ INCORRECTO: Procesamiento secuencial
results = [process_doc(doc) for doc in docs]
```

---

## 🔍 Solución de Problemas

### Problema: "Java no encontrado"

**Síntoma**: Error al iniciar aplicación
```
FileNotFoundError: Java no está instalado
```

**Solución**:
1. Instalar Java JRE 8+
2. Verificar con `java -version`
3. Agregar Java al PATH del sistema

### Problema: Memoria insuficiente en Streamlit Cloud

**Síntoma**: Aplicación se detiene al procesar muchos archivos

**Solución**:
1. Reducir número de URLs a 500-700
2. Dividir Excel en múltiples archivos
3. Ejecutar en local para lotes grandes

### Problema: LanguageTool muy lento

**Síntoma**: Análisis toma mucho tiempo

**Solución**:
1. Instalar LanguageTool local (no usar API pública)
2. Reducir `max_chars_call` a 15000
3. Desactivar análisis de modismos si no es necesario

### Problema: Documentos PDF no se procesan

**Síntoma**: PDF aparece como "sin texto"

**Solución**:
1. Verificar que el PDF no sea imagen escaneada
2. Usar OCR si el PDF es imagen
3. Convertir PDF a DOCX y volver a intentar

### Problema: Excel de salida corrupto

**Síntoma**: No se puede abrir el archivo descargado

**Solución**:
1. Verificar que el navegador completó la descarga
2. Intentar descarga manual (botón en Paso 8)
3. Limpiar caché del navegador

---

## 🤝 Contribuir

### Cómo Contribuir

1. **Fork** el repositorio
2. Crea una **rama** para tu feature (`git checkout -b feature/nueva-funcionalidad`)
3. **Commit** tus cambios (`git commit -m 'Agrega nueva funcionalidad'`)
4. **Push** a la rama (`git push origin feature/nueva-funcionalidad`)
5. Abre un **Pull Request**

### Guía de Estilo

- **Código**: Seguir PEP 8
- **Nombres**: Usar español para variables de negocio, inglés para técnicas
- **Documentación**: Docstrings en todas las funciones públicas
- **Tests**: Incluir tests para nuevas funcionalidades

### Áreas de Contribución

- 🐛 **Bug fixes**: Reportar y corregir errores
- ✨ **Features**: Nuevas funcionalidades
- 📚 **Documentación**: Mejorar README y comentarios
- 🎨 **UI/UX**: Mejorar interfaz de usuario
- ⚡ **Performance**: Optimizaciones de velocidad

---

## 🗺️ Roadmap

### v2.0 (Q2 2026)

- [ ] Soporte para OCR en PDFs escaneados
- [ ] Análisis de plagio con algoritmos de similitud
- [ ] API REST para integración con otros sistemas
- [ ] Dashboard de métricas históricas
- [ ] Soporte para más idiomas (francés, alemán, portugués)

### v2.1 (Q3 2026)

- [ ] Integración con Google Drive y OneDrive
- [ ] Análisis de legibilidad (índice Flesch)
- [ ] Detección de sesgo en lenguaje
- [ ] Exportación a formatos adicionales (JSON, CSV)
- [ ] Modo batch por línea de comandos

### v3.0 (Q4 2026)

- [ ] Integración con modelos de IA (GPT) para sugerencias contextuales
- [ ] Editor inline para correcciones
- [ ] Sistema de roles y permisos
- [ ] Base de datos para historial de análisis
- [ ] Aplicación móvil (iOS/Android)

---

## 📊 Métricas del Proyecto

- **Líneas de código**: ~4,970
- **Módulos**: 2 (Home, GrammarScan)
- **Formatos soportados**: 7 (PDF, DOCX, PPTX, TXT, ZIP, DOC, PPT)
- **Idiomas**: 2+ (Español, Inglés, extensible)
- **Máximo procesamiento**: 700 URLs (Cloud) / ilimitado (Local)

---

## 🙏 Agradecimientos

- **LanguageTool**: Motor de análisis gramatical
- **Streamlit**: Framework de aplicación web
- **Universidad Tecnológica del Perú**: Patrocinio y casos de uso

---

## 📞 Contacto

- **Desarrollador**: José Luis Antunez Condezo
- **Email**: joseluisantunezcondezo@utp.edu.pe
- **GitHub**: [@joseluisantunezcondezo](https://github.com/joseluisantunezcondezo)
- **Proyecto**: [UTP_GrammarScan](https://github.com/joseluisantunezcondezo/UTP_GrammarScan)

---

## 📜 Licencia

Este proyecto está bajo la Licencia MIT. Ver el archivo [LICENSE](LICENSE) para más detalles.

---

## 🌟 Star History

Si este proyecto te fue útil, ¡no olvides darle una ⭐ en GitHub!

[![Star History Chart](https://api.star-history.com/svg?repos=joseluisantunezcondezo/UTP_GrammarScan&type=Date)](https://star-history.com/#joseluisantunezcondezo/UTP_GrammarScan&Date)

---

**Hecho con ❤️ para la comunidad académica de UTP**
