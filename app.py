import streamlit as st
import os
import json
import base64
from pathlib import Path
from notebook_generator import (
    PAIS_MAP, CARTERA_MAP, CARTERA_LABEL,
    generar_notebook_generador,
    generar_notebook_motor,
    guardar_notebook,
    generar_nombres_archivos,
)

# ==================== CONFIG ====================
st.set_page_config(
    page_title="SARIMAX Notebook Generator",
    page_icon="📊",
    layout="wide"
)

st.title("🧪 SARIMAX IFRS 9 · Generador de Notebooks")
st.markdown("Configura parámetros y descarga los notebooks listos para ejecutar.")

# ==================== SIDEBAR CONFIG ====================
st.sidebar.header("⚙️ Configuración")

# Directorio de templates (asumir que están en el mismo directorio)
TEMPLATES_DIR = Path(__file__).parent
TEMPLATE_GENERADOR = TEMPLATES_DIR / "Generacion_Variacion__2_.ipynb"
TEMPLATE_MOTOR = TEMPLATES_DIR / "Motor_Sarimax_Vivi_CO__1_.ipynb"

# Validar que existen los templates
templates_ok = TEMPLATE_GENERADOR.exists() and TEMPLATE_MOTOR.exists()
if not templates_ok:
    st.error(f"❌ Falta un template. Verifica que estos archivos estén en el mismo directorio que app.py:")
    st.code(f"{TEMPLATE_GENERADOR.name}\n{TEMPLATE_MOTOR.name}")
    st.stop()

# ==================== FORMULARIO ====================

col1, col2 = st.columns(2)

with col1:
    st.subheader("Ubicación")
    pais_display = st.selectbox(
        "País",
        options=list(PAIS_MAP.keys()),
        help="Selecciona el país del modelo"
    )
    pais = PAIS_MAP[pais_display]

    cartera_display = st.selectbox(
        "Cartera",
        options=list(CARTERA_MAP.keys()),
        help="Selecciona el portafolio de crédito"
    )
    cartera = CARTERA_MAP[cartera_display]

with col2:
    st.subheader("Modo Endógena")
    modo_endo = st.radio(
        "Cálculo de la endógena",
        options=["actual", "media_movil"],
        help="'actual': valor original | 'media_movil': promedio móvil"
    )

    if modo_endo == "media_movil":
        ventana_mm = st.slider(
            "Ventana de media móvil (meses)",
            min_value=1,
            max_value=12,
            value=3,
            help="Número de meses para el cálculo de promedio móvil"
        )
    else:
        ventana_mm = 3  # No se usa pero lo dejamos


# ==================== CONFIGURACIÓN GENERADOR ====================

st.markdown("---")
st.subheader("📋 Configuración Generador")

col1, col2 = st.columns(2)

with col1:
    usar_fechas_default = st.checkbox(
        "Usar fechas predeterminadas",
        value=True,
        help="Activar para usar: 2018-10-01 a 2025-03-01"
    )

    if usar_fechas_default:
        fecha_inicio = "2018-10-01"
        fecha_fin = "2025-03-01"
        st.info(f"📅 Inicio: {fecha_inicio} | Fin: {fecha_fin}")
    else:
        col_a, col_b = st.columns(2)
        with col_a:
            fecha_inicio = st.text_input(
                "Fecha inicio (YYYY-MM-DD)",
                value="2018-10-01"
            )
        with col_b:
            fecha_fin = st.text_input(
                "Fecha fin histórico (YYYY-MM-DD)",
                value="2025-03-01"
            )

with col2:
    editar_nombres = st.checkbox(
        "Editar nombres de archivos",
        value=False,
        help="Si marcas esto, puedes personalizar los nombres"
    )

    if editar_nombres:
        nombres_sugeridos = generar_nombres_archivos(pais, cartera)
        st.warning("⚠️ Cambiar el formato puede romper el flujo. Procede con cuidado.")

        col_h, col_o = st.columns(2)
        with col_h:
            archivo_hist = st.text_input(
                "Nombre archivo histórico",
                value=nombres_sugeridos["hist"]
            )
            archivo_base = st.text_input(
                "Nombre archivo base",
                value=nombres_sugeridos["base"]
            )
        with col_o:
            archivo_opt = st.text_input(
                "Nombre archivo optimista",
                value=nombres_sugeridos["opt"]
            )
            archivo_adv = st.text_input(
                "Nombre archivo adverso",
                value=nombres_sugeridos["adv"]
            )

        nombres_custom = {
            "hist": archivo_hist,
            "opt": archivo_opt,
            "base": archivo_base,
            "adv": archivo_adv,
        }
    else:
        nombres_custom = None


# ==================== CONFIGURACIÓN MOTOR ====================

st.markdown("---")
st.subheader("⚙️ Configuración Motor SARIMAX")

col1, col2 = st.columns(2)

with col1:
    tipo_modelo = st.selectbox(
        "Tipo de modelo",
        options=["total", "logit"],
        help="'total': escala original | 'logit': transformación logit"
    )

with col2:
    max_lags = st.slider(
        "Máximo de lags",
        min_value=1,
        max_value=12,
        value=8,
        help="Número máximo de rezagos a evaluar en los modelos"
    )

st.info("""
**Nota:** Los siguientes parámetros NO se editan aquí (requieren cambios en Colab):
- Signos exógenas (positivo/negativo)
- VIF máximo
- Umbral de sensibilidad
- Factor FWL (1.0 - 1.2)

Para modificarlos, edita directamente en el notebook.
""")


# ==================== BOTÓN DE GENERACIÓN ====================

st.markdown("---")

if st.button("🚀 Generar Notebooks", use_container_width=True, type="primary"):

    try:
        # Generar nombres de archivos (usamos cartera código para lógica interna)
        archivos = generar_nombres_archivos(pais, cartera)
        if nombres_custom:
            archivos = nombres_custom

        # Generar notebook del GENERADOR
        st.info("Generando notebook del Generador...")
        nb_gen = generar_notebook_generador(
            str(TEMPLATE_GENERADOR),
            pais=pais,
            cartera=cartera,
            fecha_inicio=fecha_inicio,
            fecha_fin=fecha_fin,
            modo_endo=modo_endo,
            ventana_mm=ventana_mm,
            nombres_custom=nombres_custom,
        )

        # Generar notebook del MOTOR
        st.info("Generando notebook del Motor...")
        nb_motor = generar_notebook_motor(
            str(TEMPLATE_MOTOR),
            pais=pais,
            cartera=cartera,
            tipo_modelo=tipo_modelo,
            max_lags=max_lags,
            nombres_archivos=archivos,
        )

        # Nombres de salida: usamos cartera_display para que refleje lo que el usuario seleccionó
        cartera_slug = cartera_display.lower().replace(" ", "_")
        nb_gen_name = f"Generacion_Variacion_{pais}_{cartera_slug}.ipynb"
        nb_motor_name = f"Motor_Sarimax_{cartera_slug}_{pais}.ipynb"

        # Convertir notebooks a JSON (string) para descarga individual
        nb_gen_json = json.dumps(nb_gen, ensure_ascii=False, indent=1)
        nb_motor_json = json.dumps(nb_motor, ensure_ascii=False, indent=1)

        # Mostrar resumen
        st.success("✅ Notebooks generados correctamente")

        st.markdown("### 📦 Resumen de configuración")
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("País", pais)
        with col2:
            st.metric("Cartera", cartera_display)  # ← FIX: muestra el nombre que el usuario seleccionó
        with col3:
            st.metric("Modelo", tipo_modelo)
        with col4:
            st.metric("Max Lags", max_lags)

        st.markdown("### 📥 Descargar o abrir en Google Colab")

        # --- NOTEBOOK GENERADOR ---
        st.markdown("#### 📋 Generador de Variación")

        col_gen1, col_gen2 = st.columns(2)

        with col_gen1:
            st.download_button(
                label=f"💾 Descargar {nb_gen_name}",
                data=nb_gen_json,
                file_name=nb_gen_name,
                mime="application/json",
                use_container_width=True,
            )

        with col_gen2:
            # Codificar el notebook en base64 para el link de Colab
            nb_gen_b64 = base64.b64encode(nb_gen_json.encode("utf-8")).decode("utf-8")
            colab_url_gen = f"https://colab.research.google.com/notebook#data={nb_gen_b64}"
            st.link_button(
                label="🚀 Abrir en Google Colab",
                url=colab_url_gen,
                use_container_width=True,
            )

        # --- NOTEBOOK MOTOR ---
        st.markdown("#### ⚙️ Motor SARIMAX")

        col_mot1, col_mot2 = st.columns(2)

        with col_mot1:
            st.download_button(
                label=f"💾 Descargar {nb_motor_name}",
                data=nb_motor_json,
                file_name=nb_motor_name,
                mime="application/json",
                use_container_width=True,
            )

        with col_mot2:
            # Codificar el notebook en base64 para el link de Colab
            nb_motor_b64 = base64.b64encode(nb_motor_json.encode("utf-8")).decode("utf-8")
            colab_url_motor = f"https://colab.research.google.com/notebook#data={nb_motor_b64}"
            st.link_button(
                label="🚀 Abrir en Google Colab",
                url=colab_url_motor,
                use_container_width=True,
            )

    except Exception as e:
        st.error(f"❌ Error al generar notebooks:\n{str(e)}")
        st.exception(e)


# ==================== PIE ====================

st.markdown("---")
st.markdown("""
**Instrucciones:**
1. Configura los parámetros arriba
2. Haz clic en "Generar Notebooks"
3. Descarga los notebooks o ábrelos directamente en Google Colab
4. Ejecuta en orden: Generador → Motor → Dashboard

**Documentación:** Consulta con tu área de riesgo cuantitativo.
""")
