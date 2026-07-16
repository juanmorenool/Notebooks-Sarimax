# 🧪 SARIMAX Notebook Generator

Herramienta Streamlit para generar automáticamente notebooks de SARIMAX IFRS 9 con parámetros personalizados.

## 🚀 Instalación rápida

### 1. Clonar o descargar el repositorio
```bash
git clone <repo-url>
cd sarimax-generator
```

### 2. Instalar dependencias
```bash
pip install streamlit
```

### 3. Ejecutar la aplicación
```bash
streamlit run app.py
```

La app abrirá en `http://localhost:8501`

---

## 📋 Archivos necesarios

La app necesita estos dos notebooks como **templates**:
- `Generacion_Variacion__2_.ipynb` (Generador)
- `Motor_Sarimax_Vivi_CO__1_.ipynb` (Motor)

Coloca ambos en el **mismo directorio** que `app.py`.

---

## 🎮 Cómo usar

1. **Selecciona País y Cartera** (columna izquierda)
2. **Configura modo endógena**: actual o media móvil
3. **Usa fechas predeterminadas** o personaliza (2018-10-01 a 2025-03-01 por defecto)
4. **Selecciona tipo de modelo**: total o logit
5. **Ajusta máximo de lags**: 1-12 (default 8)
6. **Haz clic en "Generar Notebooks"**
7. **Descarga el ZIP** con ambos notebooks generados

---

## 📁 Estructura del proyecto

```
sarimax-generator/
├── app.py                              # App principal de Streamlit
├── notebook_generator.py                # Lógica de generación
├── Generacion_Variacion__2_.ipynb      # Template del generador
├── Motor_Sarimax_Vivi_CO__1_.ipynb     # Template del motor
└── README.md                           # Este archivo
```

---

## ⚙️ Parámetros editables en la UI

### Generador (ambos notebooks)
- **País**: Colombia (CO), Panamá (PA), Costa Rica (CR)
- **Cartera**: vivienda, consumo, tarjeta, vehículo, corporativo, pymes, comercial, hipotecas
- **Modo endógena**: actual | media móvil
- **Ventana media móvil**: 1-12 meses (si modo = media móvil)
- **Fechas**: Personalizables o predeterminadas

### Motor
- **Tipo de modelo**: total | logit
- **Max Lags**: 1-12

---

## ⚠️ Parámetros NO editables en la UI

Los siguientes parámetros deben editarse **directamente en Colab** después de descargar:
- Signos exógenas (positivo/negativo)
- VIF máximo
- Umbral de sensibilidad
- Factor FWL (debe estar entre 1.0-1.2)

---

## 🔄 Flujo de uso típico

1. **Genera los notebooks** con la herramienta Streamlit
2. **Descarga el ZIP** con ambos .ipynb
3. **Sube a Google Colab**
4. **Ejecuta primero**: `Generacion_Variacion_XX_XX.ipynb`
   - Genera 4 archivos Excel (hist, opt, base, adv)
5. **Ejecuta después**: `Motor_Sarimax_XX_XX.ipynb`
   - Toma los 4 Excel como input
   - Genera archivo de modelos + reporte CSV
6. **Sube los outputs al dashboard** (app.py de revisión)

---

## 📊 Mapeos país/cartera

| Valor UI | Código |
|----------|--------|
| Colombia | CO |
| Panamá | PA |
| Costa Rica | CR |
| vivienda | vivi |
| consumo | cons |
| tarjeta | tc |
| vehículo | vehic |
| corporativo | comercial |
| pymes | pyme |
| hipotecas | hipo |

---

## 🐛 Troubleshooting

### "Falta un template"
✓ Verifica que ambos `.ipynb` estén en el mismo directorio que `app.py`

### "Error al generar notebooks"
✓ Revisa la consola de Streamlit para el mensaje de error exacto
✓ Valida que los valores seleccionados sean válidos

### Descarga no funciona
✓ Intenta actualizar la página
✓ Prueba en otro navegador

---

## 📝 Notas

- Los notebooks generados son **independientes**: cada uno tiene sus parámetros embebidos
- Los nombres de archivos se generan automáticamente (puedes editarlos manualmente si activas la opción)
- Metadatos de trazabilidad se embeben en los Excel del generador (JSON en custom properties)

---

## 🔗 Próximas versiones

- [ ] Carga de templates personalizados
- [ ] Guardado de configuraciones (JSON)
- [ ] Validación automática de combinaciones país/cartera
- [ ] Preview de cambios antes de generar
- [ ] Integración con Google Drive para descargas automáticas

---

**Versión**: 1.0 (MVP)  
**Última actualización**: Julio 2026
