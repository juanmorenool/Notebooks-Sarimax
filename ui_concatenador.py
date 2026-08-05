"""
Interfaz Streamlit para el concatenador de archivos Impacto.
Integrable directamente en la app principal.
"""

import streamlit as st
import pandas as pd
from concatenador_streamlit import procesar_archivos_impacto, generar_csv_bytes


def render_concatenador(NAVY, BLUE, GREEN, RED, GRAY, LTGRAY, TEXT, MUTED, BG, WHITE, BORDER, TINT):
    """
    Renderiza la interfaz del concatenador como un tab/seccion.

    Args:
        NAVY, BLUE, GREEN, RED, etc.: Variables de color de la paleta corporativa
    """

    # Estilos y helpers (reutilizados de la app principal)
    def section_title(text):
        return f"<p style='font-size:13px;font-weight:700;color:{NAVY};margin:0 0 12px;text-transform:uppercase;letter-spacing:0.5px;'>{text}</p>"

    def divider():
        return f"<div style='height:1px;background:{BORDER};margin:16px 0;'></div>"

    def card_kpi(title, value, subtitle="", accent=NAVY):
        sub = f'<p style="font-size:12px;color:{MUTED};margin:4px 0 0 0;line-height:1.3;">{subtitle}</p>' if subtitle else ''
        return f"""
        <div style="background:{WHITE};border:1px solid {BORDER};border-radius:8px;padding:14px 16px;height:100%;box-sizing:border-box;">
            <p style="font-size:10px;color:{LTGRAY};margin:0 0 6px;text-transform:uppercase;letter-spacing:0.6px;font-weight:600;">{title}</p>
            <p style="font-size:18px;font-weight:700;color:{accent};margin:0;line-height:1.2;">{value}</p>
            {sub}
        </div>
        """

    # =====================================================================
    # HEADER
    # =====================================================================
    st.markdown(f"""
    <div style="margin-bottom:4px;">
        <p style="font-size:20px;font-weight:700;color:{NAVY};margin:0;">Concatenador de Impactos FWL</p>
        <p style="font-size:12px;color:{MUTED};margin:4px 0 0;">Carga multiple archivos CSV de impacto y obten un unico archivo consolidado.</p>
    </div>
    <div style="height:1px;background:{BORDER};margin:12px 0 20px;"></div>
    """, unsafe_allow_html=True)

    # =====================================================================
    # FORMULARIO
    # =====================================================================
    col1, col2 = st.columns(2)

    with col1:
        st.markdown(section_title("Configuracion"), unsafe_allow_html=True)
        pais = st.selectbox(
            "Pais",
            options=["CO", "PA", "CR"],
            help="Codigo del pais (debe coincidir en los archivos)"
        )

        # --- Formato de fecha con opcion predeterminada ---
        usar_fecha_default = st.checkbox(
            "Usar formato de fecha predeterminado",
            value=True,
            help="Activa para usar el formato estandar %d%b%Y (ej: 01mar26)"
        )

        if usar_fecha_default:
            fecha_format = "%d%b%Y"
            st.info(f"Formato activo: {fecha_format}")
        else:
            st.warning(
                "ADVERTENCIA: Cambiar el formato de fecha puede afectar la correcta lectura de los archivos. "
                "Verifique que el formato coincida exactamente con el de sus archivos CSV antes de continuar."
            )
            fecha_format = st.text_input(
                "Formato de fecha personalizado",
                value="%d%b%Y",
                help="Ejemplo: %d%b%Y para '01mar26', %Y-%m-%d para '2026-03-01'"
            )

    with col2:
        st.markdown(section_title("Archivos a cargar"), unsafe_allow_html=True)
        st.markdown("""
        <div style="background:#F0F7FF;border:1px solid #B3D9FF;border-radius:6px;padding:12px;margin-bottom:12px;">
            <p style="font-size:11px;color:#004494;margin:0;"><b>Patron esperado:</b></p>
            <p style="font-size:10px;color:#004494;margin:4px 0 0;">
                Impacto_[producto]_[PAIS].csv<br>
                Ejemplo: Impacto_cons_CO.csv, Impacto_tarjeta_CO.csv
            </p>
        </div>
        """, unsafe_allow_html=True)

    # =====================================================================
    # FILE UPLOADER
    # =====================================================================
    st.markdown(section_title("Subir archivos"), unsafe_allow_html=True)
    archivos_cargados = st.file_uploader(
        "Selecciona los archivos CSV",
        type=["csv"],
        accept_multiple_files=True,
        label_visibility="collapsed",
        help="Selecciona todos los archivos Impacto_* que deseas concatenar"
    )

    # =====================================================================
    # PROCESAMIENTO
    # =====================================================================
    if archivos_cargados:
        with st.spinner("Procesando archivos..."):
            resultado = procesar_archivos_impacto(
                archivos_cargados,
                pais=pais,
                fecha_format=fecha_format
            )

        df_resultado = resultado['df']
        avisos = resultado['avisos']
        errores = resultado['errores']
        archivos_proc = resultado['archivos_procesados']

        # --- Mostrar errores ---
        if errores:
            st.markdown(section_title("Errores"), unsafe_allow_html=True)
            for error in errores:
                st.error(error)

        # --- Mostrar avisos ---
        if avisos:
            st.markdown(section_title("Avisos"), unsafe_allow_html=True)
            for aviso in avisos:
                st.info(aviso)

        # --- Resumen ---
        st.markdown(divider(), unsafe_allow_html=True)
        st.markdown(section_title("Resumen del procesamiento"), unsafe_allow_html=True)

        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.markdown(card_kpi("Archivos cargados", str(len(archivos_cargados))), unsafe_allow_html=True)
        with c2:
            st.markdown(card_kpi("Archivos procesados", str(len(archivos_proc)), accent=GREEN), unsafe_allow_html=True)
        with c3:
            st.markdown(card_kpi("Errores", str(len(errores)), accent=RED if errores else GRAY), unsafe_allow_html=True)
        with c4:
            if df_resultado is not None:
                st.markdown(card_kpi("Filas resultado", str(len(df_resultado)), accent=BLUE), unsafe_allow_html=True)

        # --- Vista previa de datos ---
        if df_resultado is not None and not df_resultado.empty:
            st.markdown(divider(), unsafe_allow_html=True)
            st.markdown(section_title("Vista previa del resultado"), unsafe_allow_html=True)

            col_info1, col_info2 = st.columns(2)
            with col_info1:
                st.metric("Filas", len(df_resultado))
            with col_info2:
                st.metric("Columnas", len(df_resultado.columns))

            st.dataframe(df_resultado.head(20), use_container_width=True)

            # --- Descarga ---
            st.markdown(divider(), unsafe_allow_html=True)
            st.markdown(section_title("Descargar resultado"), unsafe_allow_html=True)

            csv_bytes = generar_csv_bytes(df_resultado)

            col_desc1, col_desc2 = st.columns(2)
            with col_desc1:
                st.download_button(
                    label="Descargar CSV",
                    data=csv_bytes,
                    file_name=f"fwl_{pais}_concatenado.csv",
                    mime="text/csv",
                    use_container_width=True
                )

            with col_desc2:
                # Opcion alternativa: Excel
                excel_buffer = __import__('io').BytesIO()
                with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                    df_resultado.to_excel(writer, index=False, sheet_name='Impactos')
                excel_buffer.seek(0)

                st.download_button(
                    label="Descargar Excel",
                    data=excel_buffer.getvalue(),
                    file_name=f"fwl_{pais}_concatenado.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

            # --- Estadisticas de las columnas ---
            st.markdown(divider(), unsafe_allow_html=True)
            st.markdown(section_title("Estadisticas de variables de impacto"), unsafe_allow_html=True)

            col_impacto = [c for c in df_resultado.columns if c.startswith("Impacto_")]
            if col_impacto:
                stats_df = df_resultado[col_impacto].describe().T
                st.dataframe(stats_df, use_container_width=True)
            else:
                st.caption("No se encontraron columnas de Impacto en el resultado.")

        else:
            if df_resultado is None:
                st.warning("No se pudo generar el resultado debido a errores en el procesamiento.")

    else:
        st.markdown(f"""
        <div style="background:{TINT};border:1px solid {BORDER};border-radius:8px;padding:40px;text-align:center;margin-top:20px;">
            <p style="font-size:13px;color:{NAVY};font-weight:600;margin:0 0 6px;">Ningun archivo seleccionado</p>
            <p style="font-size:11px;color:{MUTED};margin:0;">Sube uno o mas archivos Impacto_*.csv para comenzar.</p>
        </div>
        """, unsafe_allow_html=True)
