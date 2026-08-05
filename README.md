# SARIMAX IFRS 9 — Model Dashboard, Generator & Data Consolidator

Aplicacion Streamlit para la generacion, validacion, diagnostico, comparacion, consolidacion y documentacion metodologica de modelos SARIMAX aplicados a proyecciones de cartera bajo IFRS 9.

---

## Tabla de Contenidos

1. [Descripcion General](#descripcion-general)
2. [Arquitectura y Stack Tecnologico](#arquitectura-y-stack-tecnologico)
3. [Funcionalidades Principales](#funcionalidades-principales)
   - [Dashboard de Modelos SARIMAX](#1-dashboard-de-modelos-sarimax)
   - [Sistema de Diagnostico y Scoring](#2-sistema-de-diagnostico-y-scoring)
   - [Visualizacion de Datos](#3-visualizacion-de-datos)
   - [Generador de Documentos Metodologicos](#4-generador-de-documentos-metodologicos)
   - [Generador de Notebooks](#5-generador-de-notebooks)
   - [Concatenador de Impactos FWL](#6-concatenador-de-impactos-fwl)
4. [Estructura del Proyecto](#estructura-del-proyecto)
5. [Formato de Entrada](#formato-de-entrada)
6. [Sistema de Scoring (A–D / 0–10)](#sistema-de-scoring-ad--010)
7. [Flujo de Datos](#flujo-de-datos)
8. [Instalacion y Ejecucion Local](#instalacion-y-ejecucion-local)
9. [Despliegue en Streamlit Cloud](#despliegue-en-streamlit-cloud)
10. [Dependencias](#dependencias)
11. [Paleta Corporativa](#paleta-corporativa)
12. [Notas Tecnicas](#notas-tecnicas)
13. [Desarrollo y Contribuciones](#desarrollo-y-contribuciones)
14. [Contacto y Licencia](#contacto-y-licencia)

---

## Descripcion General

Esta aplicacion monolitica construida con Streamlit actua como dashboard interactivo y herramienta de gobierno de modelos para el motor SARIMAX (Seasonal AutoRegressive Integrated Moving Average with eXogenous regressors) usado en la estimacion de provisiones bajo IFRS 9.

El sistema permite:

- Ingestar archivos Excel con multiples modelos SARIMAX exportados por el motor de generacion.
- Parsear automaticamente cada hoja como un modelo independiente, extrayendo endogenas, exogenas, coeficientes, residuos, factores FWL y pruebas estadisticas.
- Evaluar la calidad de cada modelo mediante un sistema de scoring A–D y un score global ponderado (0–10).
- Visualizar predicciones, factores FWL, distribucion de residuos y coeficientes con Plotly.
- Comparar hasta tres modelos en paralelo.
- Generar documentacion metodologica automatica en PDF (ReportLab) y Excel (openpyxl) con graficas estaticas (matplotlib).
- Configurar y descargar notebooks Jupyter listos para ejecutar en Google Colab.
- Consolidar multiples archivos CSV de impacto FWL por pais, fecha y escenario.

---

## Arquitectura y Stack Tecnologico

| Capa | Tecnologia | Uso |
|------|-----------|-----|
| Frontend / UI | Streamlit + HTML/CSS injectado | Layout responsive, sidebar, tabs, botones, formularios |
| Visualizacion interactiva | Plotly (graph_objects) | Graficas de predicciones, FWL, histogramas, barras de coeficientes |
| Visualizacion estatica (PDF/Excel) | Matplotlib (Agg backend) | Graficas de macros, FWL e historico embebidas en documentos |
| Documentos | ReportLab (PDF), openpyxl (Excel), FPDF (legacy) | Generacion de documentos metodologicos descargables |
| Parsing / ETL | pandas, openpyxl | Lectura y normalizacion de archivos Excel multihoja y CSV |
| Estadistica | scipy.stats, statsmodels | Jarque-Bera, asimetria, curtosis de residuos, pruebas Ljung-Box, ARCH/LM |
| Serializacion | json, base64 | Metadata embebida en Excel, notebooks para Colab |
| Persistencia | JSON file (`.sidebar_prefs.json`) | Preferencias de sidebar entre sesiones |

---

## Funcionalidades Principales

### 1. Dashboard de Modelos SARIMAX

- **Carga de archivos `.xlsx`**: Cada hoja se interpreta como un modelo independiente.
- **Sidebar inteligente**:
  - Filtros por score de Ljung-Box, Jarque-Bera y Heterocedasticidad.
  - Ordenamiento por nombre, pruebas aprobadas o score global.
  - Marcado de modelos como favoritos.
  - Navegacion con flechas de teclado (izquierda / derecha) y reset con Escape.
- **Vistas**:
  - **Resumen de corrida**: KPIs globales, distribucion de scores, top y bottom 5.
  - **Vista de favoritos**: Grid de modelos marcados.
  - **Detalle de modelo**: 5 pestanas (Resumen, Visualizacion, Predicciones, Diagnosticos, Comparar).

### 2. Sistema de Diagnostico y Scoring

Cada modelo se evalua con tres pruebas estadisticas sobre sus residuos:

| Prueba | Hipotesis nula (H0) | Peso en Score Global |
|--------|---------------------|----------------------|
| **Ljung-Box** | No hay autocorrelacion en residuos | 40 % |
| **Jarque-Bera** | Los residuos son normales | 30 % |
| **Heterocedasticidad (ARCH/LM)** | Varianza homogenea (homocedasticidad) | 30 % |

**Score por prueba (A–D):**

- **A**: p > 0.10 — cumplimiento optimo
- **B**: 0.05 < p <= 0.10 — cumplimiento aceptable
- **C**: 0.01 < p <= 0.05 — senal de alerta, requiere monitoreo
- **D**: p <= 0.01 — incumplimiento significativo

**Score Global (0–10):**

```
Score Global = 0.40 * Score_Ljung + 0.30 * Score_Jarque + 0.30 * Score_Hetero
```

- **>= 7.0**: BUENO — modelo robusto para uso oficial.
- **5.0 – 6.9**: REGULAR — utilizable con seguimiento.
- **< 5.0**: DEFICIENTE — requiere reespecificacion.

### 3. Visualizacion de Datos

- **Predicciones**: Series temporales de escenarios Base, Adverso y Optimista.
- **Exogenas**: Graficas individuales por variable exogena con los 3 escenarios.
- **FWL (Forward-Looking)**:
  - FWL a 12 meses (lineas por escenario).
  - FWL ponderado configurable (pesos Base/Adverso/Optimista).
  - FWL anual resumido en tabla pivot.
- **Residuos**: Histograma con campana normal teorica superpuesta y estadisticas descriptivas.
- **Coeficientes**: Barras horizontales con color por signo (verde positivo, rojo negativo).

### 4. Generador de Documentos Metodologicos

Al marcar un modelo como **"Modelo Final"**, el sistema genera automaticamente:

- **PDF metodologico** (ReportLab) con:
  - Portada con metadata del modelo.
  - Resumen ejecutivo (grafica historica, FWL, diagnosticos, coeficientes, estructura).
  - Metodologia SARIMAX detallada (9 bloques: carga, estacionariedad, endogena, lags, dummies, combinaciones, forecast, ordenamiento, exportacion, reporte CSV).
  - Configuracion del motor (pais, cartera, tipo de endogena, VIF, rangos FWL, etc.).
  - Variables utilizadas con significancia.
  - Grafica de variables exogenas (matplotlib estatico).
  - Diagnosticos con interpretacion textual por score (A–D).
  - Score global con conclusion automatica.
- **Excel metodologico** (openpyxl) con hojas: Metadatos, Exogenas, Diagnosticos, Coeficientes, Grafica.

### 5. Generador de Notebooks

Pestana **"Generador"** que permite configurar parametros del motor SARIMAX y descargar:

- `Generacion_Variacion_{pais}_{cartera}.ipynb` — notebook de preparacion de datos.
- `Motor_Sarimax_{cartera}_{pais}.ipynb` — notebook del motor econometrico.

Con opcion de abrir directamente en **Google Colab** via URL con payload base64.

Soporte para multiples paises: Colombia (CO), Panama (PA), Costa Rica (CR).

### 6. Concatenador de Impactos FWL

- Carga de multiples archivos CSV de impacto FWL.
- Deteccion automatica de producto desde nombre de archivo.
- Normalizacion de nombres de columnas.
- Consolidacion de archivos por `[PAIS, date, Scenario_name]`.
- Descarga de resultado en CSV o Excel.
- Generacion de estadisticas de variables de impacto.

Patron de archivos soportados:

```
Impacto_consumo_CO.csv
Impacto_tarjeta_CO.csv
Impacto_vivienda_CO.csv
Impacto_pymes_CO.csv
Impacto_corporativo_CO.csv
```

Alias de productos soportados:

| Producto | Alias detectados |
|----------|-----------------|
| consumo | cons, consumo |
| corporativo | corporativo, corp |
| pymes | pyme, pymes |
| tarjeta | tarjeta |
| vehiculo | vehiculo |
| vivienda | vivi, vivienda |

---

## Estructura del Proyecto

```
notebooks-sarimax/
├── app.py                          # Aplicacion principal Streamlit
├── notebook_generator.py           # Logica de generacion de notebooks
├── concatenador_streamlit.py       # Motor de consolidacion de CSVs
├── ui_concatenador.py              # Interfaz del concatenador
├── requirements.txt                # Dependencias del proyecto
├── .devcontainer/                  # Configuracion de contenedor de desarrollo
├── Concatenador_Resultados_CO.ipynb
├── Generacion_Variacion_2_.ipynb
├── Motor_Sarimax_Vivi_CO_1_.ipynb
└── README.md
```

---

## Formato de Entrada

### Archivo Excel de Modelos (.xlsx)

El archivo esperado contiene **una hoja por modelo SARIMAX**. Cada hoja debe incluir las siguientes secciones (detectadas por nombre de columna):

| Seccion | Columnas esperadas | Descripcion |
|---------|-------------------|-------------|
| **Endogena** | `fecha`, `BASE`, `ADVERSO`, `OPTIMISTA` | Serie temporal de la variable dependiente en 3 escenarios. |
| **Exogenas** | `VAR_BASE`, `VAR_ADVERSO`, `VAR_OPTIMISTA` | Variables explicativas por escenario (sufijo automatico). |
| **FWL 12M** | `fecha`, `FWL_BASE`, `FWL_ADVERSO`, `FWL_OPTIMISTA` | Factor Forward-Looking a 12 meses. |
| **FWL Anual** | `Ano`, `Escenario`, `Factor FWL` | Resumen anual del factor FWL. |
| **Residuos** | `Obs`, `Residuo` | Residuos individuales del ajuste. |
| **Resumen Residuos** | `Estadistico`, `Valor` | Estadisticas descriptivas de residuos. |
| **Coeficientes** | `Variable`, `Coeficiente`, `P_value` | Tabla de coeficientes estimados. |
| **Pruebas** | `Prueba`, `Estadistico`, `P_value` | Resultados de Ljung-Box, Jarque-Bera, ARCH/LM. |

El parser es tolerante a mayusculas/minusculas y detecta sufijos `_BASE`, `_ADVERSO`, `_OPTIMISTA` para agrupar exogenas.

### Archivos CSV de Impacto FWL

Los archivos CSV deben contener las siguientes columnas obligatorias:

- `PAIS`
- `date`
- `Scenario_name`
- `Impacto_[producto]` (al menos una columna)

El formato de fecha por defecto es `%d%b%Y` (ejemplo: `01mar26`), configurable por el usuario.

---

## Sistema de Scoring (A–D / 0–10)

### Conversion letra a numerico

| Letra | Valor | Interpretacion |
|-------|-------|----------------|
| A | 10.0 | Optimo |
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

### Clasificacion final

| Score Global | Clasificacion | Accion recomendada |
|--------------|---------------|-------------------|
| >= 7.0 | **BUENO** | Uso oficial aprobado. |
| 5.0 – 6.9 | **REGULAR** | Monitoreo y ajustes focalizados. |
| < 5.0 | **DEFICIENTE** | Reespecificacion obligatoria. |

---

## Flujo de Datos

```
+------------------+
|  Archivo .xlsx   |
|   (multihoja)    |
+---------+--------+
          |
          v
+------------------+     +----------------------+
|  parsear_excel   |---->|  modelos_data (dict) |
|   (openpyxl)     |     |  {hoja: {endogena,   |
+------------------+     |   exogenas, fwl,     |
                         |   coefs, pruebas,    |
                         |   residuos, ...}}    |
                         +-----------+----------+
                                     |
               +---------------------+---------------------+
               |                     |                     |
               v                     v                     v
    +------------------+  +------------------+  +------------------+
    |  Dashboard UI    |  |  Scoring A–D     |  |  Doc. PDF/Excel  |
    |   (Streamlit)    |  |   (0–10 global)  |  |   (ReportLab/    |
    |                  |  |                  |  |    openpyxl)     |
    |  Visualizacion   |  |  Ljung-Box       |  |                  |
    |  Comparador      |  |  Jarque-Bera     |  |  Portada         |
    |  Favoritos       |  |  Heterocedast.   |  |  Metodologia     |
    |  Navegacion      |  |                  |  |  Diagnosticos    |
    +------------------+  +------------------+  |  Conclusion      |
                                                +------------------+
```

---

## Instalacion y Ejecucion Local

### Requisitos del sistema

- Python 3.8 o superior
- pip

### Pasos

```bash
git clone https://github.com/tu_usuario/notebooks-sarimax.git
cd notebooks-sarimax

python -m venv venv
source venv/bin/activate
# En Windows: venv\Scripts\activate

pip install -r requirements.txt

streamlit run app.py
```

La aplicacion se abrira en `http://localhost:8501`.

Para ejecucion con nivel de debug:

```bash
streamlit run app.py --logger.level=debug
```

---

## Despliegue en Streamlit Cloud

1. Sube el repositorio a GitHub.
2. Accede a [https://share.streamlit.io](https://share.streamlit.io).
3. Selecciona el repositorio.
4. Configura `app.py` como archivo principal.
5. El despliegue se ejecutara automaticamente.

La aplicacion estara disponible en:

```
https://[usuario]-notebooks-sarimax-main-[hash].streamlit.app
```

Los despliegues subsecuentes ocurren automaticamente con cada push a la rama principal.

---

## Dependencias

```
streamlit>=1.28
pandas>=2.0.0
numpy>=1.24.0
scipy>=1.11.0
openpyxl>=3.10
fpdf2>=2.7.0
matplotlib>=3.7.0
reportlab>=4.0.0
plotly>=5.18.0
pyarrow>=15.0.0
statsmodels>=0.14.0
rich==13.7.0
```

Nota: `fpdf` se mantiene por compatibilidad legacy; la generacion de PDF actual utiliza ReportLab.

---

## Paleta Corporativa

| Color | Hex | Uso |
|-------|-----|-----|
| Navy | `#001a4d` | Encabezados, elementos principales |
| Blue | `#0052cc` | Acentos, enlaces, elementos interactivos |
| Green | `#34a853` | Indicadores positivos, coeficientes positivos |
| Red | `#d33b27` | Alertas, coeficientes negativos |
| Tipografia | Inter | Texto general de la interfaz |

---

## Notas Tecnicas

- **Backend matplotlib**: Se fuerza `matplotlib.use("Agg")` para compatibilidad con entornos sin display (Streamlit Cloud, Docker).
- **Graficas estaticas**: Las imagenes para PDF y Excel se generan en archivos temporales (`tempfile.NamedTemporaryFile`) y se eliminan tras el build.
- **Metadata embebida**: El parser busca propiedades custom del workbook con prefijo `sarimax_meta_*`, divididas en chunks si exceden el limite de caracteres de openpyxl. Si no hay metadata embebida, el usuario puede completar un formulario de fallback en el sidebar.
- **Atajos de teclado**: Implementados via `components.html` con JavaScript que intercepta `ArrowLeft`, `ArrowRight` y `Escape` para navegar entre modelos y limpiar filtros.
- **CSS custom**: Inyectado via `st.markdown(..., unsafe_allow_html=True)` para lograr una UI corporativa (fuentes Inter, paleta Navy/Blue/Green/Red, tabs estilizados, sidebar sticky).
- **Navegacion sticky**: Opcional; fija una barra de navegacion (Anterior / Siguiente) en la parte inferior de la pantalla.
- **Cache de sesion**: El sistema de cache de sesion optimiza la carga de modelos repetidos.
- **Concatenador**: El merge utiliza estrategia outer para capturar todos los productos. Las columnas de impacto se ordenan alfabeticamente en la salida.

---

## Desarrollo y Contribuciones

Para agregar nuevas funcionalidades:

```bash
git checkout -b feature/nueva-funcion
git commit -m "Add: descripcion de cambios"
git push origin feature/nueva-funcion
```

Abra un Pull Request con descripcion detallada de los cambios.

---

## Contacto y Licencia

**Responsable**: Juan Jose Moreno  
**Rol**: Financial Risk Analyst - IFRS 9 Credit Risk Modeling

Ultima actualizacion: Agosto 2026  
Version: V 38.0 2026  
Licencia: Privado - Uso Interno  
Contexto: Provision de cartera bajo NIIF 9 / IFRS 9 — Modelos SARIMAX con variables macroeconomicas.
