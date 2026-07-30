# SARIMAX IFRS 9 — Model Dashboard & Generator

> **Aplicación Streamlit para la validación, diagnóstico, comparación y documentación metodológica de modelos SARIMAX aplicados a proyecciones de cartera bajo IFRS 9.**

---

## Tabla de Contenidos

1. [Descripción General](#descripción-general)
2. [Arquitectura y Stack Tecnológico](#arquitectura-y-stack-tecnológico)
3. [Funcionalidades Principales](#funcionalidades-principales)
4. [Estructura del Código](#estructura-del-código)
5. [Formato de Entrada (Excel)](#formato-de-entrada-excel)
6. [Sistema de Scoring (A–D / 0–10)](#sistema-de-scoring-ad--010)
7. [Flujo de Datos](#flujo-de-datos)
8. [Instalación y Ejecución](#instalación-y-ejecución)
9. [Dependencias](#dependencias)
10. [Notas Técnicas](#notas-técnicas)
11. [Licencia](#licencia)

---

## Descripción General

`app.py` es una aplicación monolítica construida con **Streamlit** que actúa como **dashboard interactivo** y **herramienta de gobierno de modelos** para el motor SARIMAX (Seasonal AutoRegressive Integrated Moving Average with eXogenous regressors) usado en la estimación de provisiones bajo **IFRS 9**.

El sistema permite:

- **Ingestar** archivos Excel con múltiples modelos SARIMAX exportados por el motor de generación.
- **Parsear** automáticamente cada hoja como un modelo independiente, extrayendo endógenas, exógenas, coeficientes, residuos, factores FWL y pruebas estadísticas.
- **Evaluar** la calidad de cada modelo mediante un sistema de scoring A–D y un score global ponderado (0–10).
- **Visualizar** predicciones, factores FWL, distribución de residuos y coeficientes con **Plotly**.
- **Comparar** hasta 3 modelos en paralelo.
- **Generar** documentación metodológica automática en **PDF** (ReportLab) y **Excel** (openpyxl) con gráficas estáticas (matplotlib).
- **Configurar** y descargar notebooks Jupyter listos para ejecutar en Google Colab.


---

## Arquitectura y Stack Tecnológico

| Capa | Tecnología | Uso |
|------|-----------|-----|
| **Frontend / UI** | Streamlit + HTML/CSS injectado | Layout responsive, sidebar, tabs, botones, forms |
| **Visualización interactiva** | Plotly (graph_objects) | Gráficas de predicciones, FWL, histogramas, barras de coeficientes |
| **Visualización estática (PDF)** | Matplotlib (Agg backend) | Gráficas de macros, FWL e histórico embebidas en PDF/Excel |
| **Documentos** | ReportLab (PDF), openpyxl (Excel), FPDF (legacy) | Generación de documentos metodológicos descargables |
| **Parsing / ETL** | pandas, openpyxl | Lectura y normalización de archivos Excel multihoja |
| **Estadística** | scipy.stats | Jarque-Bera, asimetría, curtosis de residuos |
| **Serialización** | json, base64 | Metadata embebida en Excel, notebooks para Colab |
| **Persistencia** | JSON file (`.sidebar_prefs.json`) | Preferencias de sidebar entre sesiones |

---

## Funcionalidades Principales

### 1. Dashboard de Modelos SARIMAX

- **Carga de archivos `.xlsx`**: Cada hoja se interpreta como un modelo independiente.
- **Sidebar inteligente**:
  - Filtros por score de Ljung-Box, Jarque-Bera y Heterocedasticidad.
  - Ordenamiento por nombre, pruebas aprobadas o score global.
  - Marcado de modelos como **favoritos**.
  - Navegación con flechas de teclado (← / →) y reset con Escape.
- **Vistas**:
  - **Resumen de corrida**: KPIs globales, distribución de scores, top/bottom 5.
  - **Vista de favoritos**: Grid de modelos marcados.
  - **Detalle de modelo**: 5 tabs (Resumen, Visualización, Predicciones, Diagnósticos, Comparar).

### 2. Sistema de Diagnóstico y Scoring

Cada modelo se evalúa con tres pruebas estadísticas sobre sus residuos:

| Prueba | Hipótesis nula (H₀) | Peso en Score Global |
|--------|---------------------|----------------------|
| **Ljung-Box** | No hay autocorrelación en residuos | 40 % |
| **Jarque-Bera** | Los residuos son normales | 30 % |
| **Heterocedasticidad (ARCH/LM)** | Varianza homogénea (homocedasticidad) | 30 % |

**Score por prueba (A–D):**

- **A**: p > 0.10 → cumplimiento óptimo
- **B**: 0.05 < p ≤ 0.10 → cumplimiento aceptable
- **C**: 0.01 < p ≤ 0.05 → señal de alerta, requiere monitoreo
- **D**: p ≤ 0.01 → incumplimiento significativo

**Score Global (0–10):**

```
Score Global = 0.40 × Score_Ljung + 0.30 × Score_Jarque + 0.30 × Score_Hetero
```

- **≥ 7.0**: BUENO — modelo robusto para uso oficial.
- **5.0 – 6.9**: REGULAR — utilizable con seguimiento.
- **< 5.0**: DEFICIENTE — requiere reespecificación.

### 3. Visualización de Datos

- **Predicciones**: Series temporales de escenarios Base, Adverso y Optimista.
- **Exógenas**: Gráficas individuales por variable exógena con los 3 escenarios.
- **FWL (Forward-Looking)**:
  - FWL a 12 meses (líneas por escenario).
  - FWL ponderado configurable (pesos Base/Adverso/Optimista).
  - FWL anual resumido en tabla pivot.
- **Residuos**: Histograma con campana normal teórica superpuesta + estadísticas descriptivas.
- **Coeficientes**: Barras horizontales con color por signo (verde positivo, rojo negativo).

### 4. Generador de Documentos Metodológicos

Al marcar un modelo como **"Modelo Final"**, el sistema genera automáticamente:

- **PDF metodológico** (ReportLab) con:
  - Portada con metadata del modelo.
  - Resumen ejecutivo (gráfica histórica, FWL, diagnósticos, coeficientes, estructura).
  - Metodología SARIMAX detallada (9 bloques: carga, estacionariedad, endógena, lags, dummies, combinaciones, forecast, ordenamiento, exportación, reporte CSV).
  - Configuración del motor (país, cartera, tipo de endógena, VIF, rangos FWL, etc.).
  - Variables utilizadas con significancia.
  - Gráfica de variables exógenas (matplotlib estático).
  - Diagnósticos con interpretación textual por score (A–D).
  - Score global con conclusión automática.
- **Excel metodológico** (openpyxl) con hojas: Metadatos, Exógenas, Diagnósticos, Coeficientes, Gráfica.

### 5. Generador de Notebooks

Pestaña **"Generador"** que permite configurar parámetros del motor SARIMAX y descargar:

- `Generacion_Variacion_{pais}_{cartera}.ipynb` — notebook de preparación de datos.
- `Motor_Sarimax_{cartera}_{pais}.ipynb` — notebook del motor econométrico.

Con opción de abrir directamente en **Google Colab** vía URL con payload base64.

### 6. Metadata y Trazabilidad

- **Lectura de metadata embebida**: El parser lee propiedades custom del workbook Excel (`sarimax_meta_*`) para extraer contexto de la corrida (país, cartera, parámetros del motor, etc.).
- **Fallback manual**: Si no hay metadata embebida, el usuario puede completar un formulario en el sidebar.

---

## Estructura del Código

```
app.py
├── PALETA CORPORATIVA BANCA          # Constantes de color, mapeos de países/carteras
├── TEXTOS DOCUMENTO METODOLÓGICO     # Strings largos para el PDF
├── PREFERENCIAS DEL SIDEBAR          # Persistencia JSON de filtros y estado
├── PARSER                            # Lectura y normalización de Excel multihoja
├── UTILIDADES                        # Helpers estadísticos y de formato
├── SCORE SYSTEM (A–D Grading)        # Lógica de scoring y badges visuales
├── METADATA                          # Extracción de KPIs de contexto
├── DOCUMENTO METODOLÓGICO MODELO FINAL
│   ├── recolectar_datos_documento()
│   ├── graficar_macros_estatico()    # matplotlib → PNG temporal
│   ├── graficar_fwl_estatico()
│   ├── graficar_historico_estatico()
│   └── generar_pdf_metodologico()    # ReportLab, ~300 líneas
├── GENERADOR DE PDF CON REPORTLAB    # (contenido extenso, ver función arriba)
├── DIAGNÓSTICOS                      # Renderizado de pruebas, métricas, leyendas
├── GENERADOR DE NOTEBOOKS            # UI de configuración y descarga
├── SESSION STATE                     # Inicialización de st.session_state
└── APP PRINCIPAL                     # Layout Streamlit: tabs, sidebar, lógica de navegación
```

---

## Formato de Entrada (Excel)

El archivo `.xlsx` esperado contiene **una hoja por modelo SARIMAX**. Cada hoja debe incluir las siguientes secciones (detectadas por nombre de columna):

| Sección | Columnas esperadas | Descripción |
|---------|-------------------|-------------|
| **Endógena** | `fecha`, `BASE`, `ADVERSO`, `OPTIMISTA` | Serie temporal de la variable dependiente en 3 escenarios. |
| **Exógenas** | `VAR_BASE`, `VAR_ADVERSO`, `VAR_OPTIMISTA` | Variables explicativas por escenario (sufijo automático). |
| **FWL 12M** | `fecha`, `FWL_BASE`, `FWL_ADVERSO`, `FWL_OPTIMISTA` | Factor Forward-Looking a 12 meses. |
| **FWL Anual** | `Año`, `Escenario`, `Factor FWL` | Resumen anual del factor FWL. |
| **Residuos** | `Obs`, `Residuo` | Residuos individuales del ajuste. |
| **Resumen Residuos** | `Estadistico`, `Valor` | Estadísticas descriptivas de residuos. |
| **Coeficientes** | `Variable`, `Coeficiente`, `P_value` | Tabla de coeficientes estimados. |
| **Pruebas** | `Prueba`, `Estadistico`, `P_value` | Resultados de Ljung-Box, Jarque-Bera, ARCH/LM. |

> El parser es tolerante a mayúsculas/minúsculas y detecta sufijos `_BASE`, `_ADVERSO`, `_OPTIMISTA` para agrupar exógenas.

---

## Sistema de Scoring (A–D / 0–10)

### Conversión letra → numérico

| Letra | Valor | Interpretación |
|-------|-------|----------------|
| A | 10.0 | Óptimo |
| B | 7.5 | Aceptable |
| C | 5.0 | Revisar |
| D | 2.5 | No cumple |

### Pesos del Score Global

```python
PESOS_SCORE_GLOBAL = {
    'ljung_box': 0.40,
    'jarque_bera': 0.30,
    'heterocedasticidad': 0.30
}
```

### Clasificación final

| Score Global | Clasificación | Acción recomendada |
|--------------|---------------|-------------------|
| ≥ 7.0 | **BUENO** | Uso oficial aprobado. |
| 5.0 – 6.9 | **REGULAR** | Monitoreo y ajustes focalizados. |
| < 5.0 | **DEFICIENTE** | Reespecificación obligatoria. |

---

## Flujo de Datos

```
┌─────────────────┐
│  Archivo .xlsx  │
│  (multihoja)    │
└────────┬────────┘
         │
         ▼
┌─────────────────┐     ┌──────────────────────┐
│  parsear_excel  │────▶│  modelos_data (dict) │
│  (openpyxl)     │     │  {hoja: {endogena,   │
└─────────────────┘     │   exogenas, fwl,     │
                        │   coefs, pruebas,    │
                        │   residuos, ...}}    │
                        └──────────┬───────────┘
                                   │
              ┌────────────────────┼────────────────────┐
              ▼                    ▼                    ▼
    ┌─────────────────┐  ┌─────────────────┐  ┌─────────────────┐
    │  Dashboard UI   │  │  Scoring A–D    │  │  Doc. PDF/Excel │
    │  (Streamlit)    │  │  (0–10 global)  │  │  (ReportLab/    │
    │                 │  │                 │  │   openpyxl)     │
    │  • Visualización│  │  • Ljung-Box    │  │                 │
    │  • Comparador   │  │  • Jarque-Bera  │  │  • Portada      │
    │  • Favoritos    │  │  • Heterocedast.│  │  • Metodología  │
    │  • Navegación   │  │                 │  │  • Diagnósticos │
    └─────────────────┘  └─────────────────┘  │  • Conclusión   │
                                              └─────────────────┘
```

---

## Instalación y Ejecución

### 1. Clonar el repositorio

```bash
git clone https://github.com/tu-org/sarimax-ifrs9-dashboard.git
cd sarimax-ifrs9-dashboard
```

### 2. Crear entorno virtual (recomendado)

```bash
python -m venv venv
source venv/bin/activate        # Linux / macOS
venv\Scripts\activate         # Windows
```

### 3. Instalar dependencias

```bash
pip install -r requirements.txt
```

### 4. Ejecutar la aplicación

```bash
streamlit run app.py
```

La aplicación se servirá por defecto en `http://localhost:8501`.

---

## Dependencias

```txt
streamlit>=1.28
pandas>=2.0
numpy>=1.24
plotly>=5.15
scipy>=1.10
openpyxl>=3.1
matplotlib>=3.7
reportlab>=4.0
fpdf>=1.7
```

> Nota: `fpdf` se mantiene por compatibilidad legacy; la generación de PDF actual usa **ReportLab**.

---

## Notas Técnicas

- **Backend matplotlib**: Se fuerza `matplotlib.use("Agg")` para compatibilidad con entornos sin display (Streamlit Cloud, Docker).
- **Gráficas estáticas**: Las imágenes para PDF/Excel se generan en archivos temporales (`tempfile.NamedTemporaryFile`) y se eliminan tras el build.
- **Metadata embebida**: El parser busca propiedades custom del workbook con prefijo `sarimax_meta_*`, divididas en chunks si exceden el límite de caracteres de openpyxl.
- **Atajos de teclado**: Implementados vía `components.html` con JavaScript que intercepta `ArrowLeft`, `ArrowRight` y `Escape` para navegar entre modelos y limpiar filtros.
- **CSS custom**: Inyectado vía `st.markdown(..., unsafe_allow_html=True)` para lograr una UI corporativa (fuentes Inter, paleta Navy/Blue/Green/Red, tabs estilizados, sidebar sticky).
- **Navegación sticky**: Opcional; fija una barra de navegación (Anterior / Siguiente) en la parte inferior de la pantalla.

---

## Licencia

Proyecto interno — Uso exclusivo del equipo de Modelos de Riesgo / IFRS 9.

---

>   
> **Versión**:  V 38.0 2026 
> **Contexto**: Provisión de cartera bajo NIIF 9 / IFRS 9 — Modelos SARIMAX con variables macroeconómicas.
