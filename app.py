import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from scipy import stats
import json
import os
import base64
import openpyxl
from pathlib import Path

# =============================================================================
# PALETA CORPORATIVA BANCA
# =============================================================================
NAVY    = "#0B2545"
BLUE    = "#1E5AA8"
GREEN   = "#1A6B3E"
RED     = "#B83232"
GRAY    = "#6C757D"
LTGRAY  = "#ADB5BD"
BG      = "#F8F9FA"
WHITE   = "#FFFFFF"
BORDER  = "#DEE2E6"
TEXT    = "#1A1D23"
MUTED   = "#5A6270"
TINT    = "#EDF2F7"

ESTADO_OK      = ("#E8F5E9", "#1A6B3E", "CUMPLE")
ESTADO_WARN    = ("#FFF8E1", "#B8860B", "REVISAR")
ESTADO_FAIL    = ("#FFEBEE", "#B83232", "NO CUMPLE")
ESTADO_NEUTRAL = ("#F5F5F5", "#5A6270", "N/A")

SCORE_COLORS = {
    'A': ("#E8F5E9", "#1A6B3E"),
    'B': ("#E3F2FD", "#1565C0"),
    'C': ("#FFF8E1", "#B8860B"),
    'D': ("#FFEBEE", "#B83232"),
}

SCORE_NUM_MAP = {'A': 10.0, 'B': 7.5, 'C': 5.0, 'D': 2.5}
PESOS_SCORE_GLOBAL = {'ljung_box': 0.40, 'jarque_bera': 0.30, 'heterocedasticidad': 0.30}

PAISES_MAP = {
    'colombia': 'Colombia',
    'panama': 'Panama',
    'panamá': 'Panama',
    'costa rica': 'Costa Rica',
    'co': 'Colombia',
    'pa': 'Panama',
    'cr': 'Costa Rica',
}

BANDERAS_PAISES = {
    'Colombia': 'co',
    'Panama': 'pa',
    'Costa Rica': 'cr',
    'co': 'co', 'pa': 'pa', 'cr': 'cr',
    'CO': 'co', 'PA': 'pa', 'CR': 'cr',
}

CARTERAS_MAP = {
    'vivi': 'Vivienda', 'vivienda': 'Vivienda',
    'cons': 'Consumo', 'consumo': 'Consumo',
    'com': 'Comercial', 'comercial': 'Comercial',
    'micro': 'Microcredito',
}

# --- Mapeos para el generador ---
PAIS_MAP_GEN = {
    'Colombia': 'CO',
    'Panama': 'PA',
    'Costa Rica': 'CR',
}

CARTERA_MAP_GEN = {
    'Vivienda': 'vivienda',
    'Consumo': 'consumo',
    'Comercial': 'comercial',
    'Microcredito': 'microcredito',
    'Corporativo': 'corporativo',
}

CARTERA_LABEL_GEN = {
    'vivienda': 'Vivienda',
    'consumo': 'Consumo',
    'comercial': 'Comercial',
    'microcredito': 'Microcredito',
    'corporativo': 'Corporativo',
}

def inject_css():
    st.markdown(f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    html, body, [class*="css"], .stApp, .stMarkdown, .stDataFrame, .stButton,
    .stSelectbox, .stTextInput, .stNumberInput, .stTabs {{
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif !important;
    }}
    #MainMenu {{ visibility: hidden; }}
    footer {{ visibility: hidden; }}
    header[data-testid="stHeader"] {{ background: transparent; }}
    .block-container {{ padding-top: 0.5rem; padding-bottom: 2rem; max-width: 1400px; }}
    .stApp {{ background-color: {BG}; }}
    section[data-testid="stSidebar"] {{
        background-color: {WHITE}; border-right: 1px solid {BORDER};
        min-width: 260px !important; max-width: 260px !important;
    }}
    section[data-testid="stSidebar"] .block-container {{ padding: 16px; }}
    h1 {{ color: {NAVY} !important; font-weight: 700 !important; font-size: 22px !important; }}
    h2 {{ color: {NAVY} !important; font-weight: 600 !important; font-size: 16px !important; margin-top: 0.2rem !important; }}
    h3 {{ color: {NAVY} !important; font-weight: 600 !important; font-size: 14px !important;
         border-left: 3px solid {BLUE}; padding-left: 10px; margin-top: 0.3rem !important; }}
    .stTabs [data-baseweb="tab-list"] {{ gap: 2px; border-bottom: 1px solid {BORDER}; }}
    .stTabs [data-baseweb="tab"] {{
        background-color: transparent; border-radius: 6px 6px 0 0;
        padding: 8px 16px; color: {MUTED}; font-weight: 500; font-size: 13px;
    }}
    .stTabs [aria-selected="true"] {{
        background-color: {TINT} !important; color: {NAVY} !important; font-weight: 600;
    }}
    div[data-testid="stDataFrame"] {{ border: 1px solid {BORDER}; border-radius: 6px; }}
    section[data-testid="stSidebar"] {{ position: sticky; top: 0; height: 100vh; overflow-y: auto; }}
    .st-key-kb_hidden {{
        position: fixed !important; top: -9999px !important; left: -9999px !important;
        height: 1px !important; width: 1px !important; overflow: hidden !important;
    }}
    section[data-testid="stSidebar"] button[kind="secondary"][class*="st-key-btn_sec_"],
    section[data-testid="stSidebar"] .st-key-btn_sec_contexto button,
    section[data-testid="stSidebar"] .st-key-btn_sec_orden button,
    section[data-testid="stSidebar"] .st-key-btn_sec_exogenas button {{
        background: none !important; border: none !important; box-shadow: none !important;
        padding: 4px 0 !important; text-align: left !important; justify-content: flex-start !important;
        font-size: 13px !important; font-weight: 700 !important; color: {NAVY} !important;
        text-transform: uppercase; letter-spacing: 0.5px;
    }}
    section[data-testid="stSidebar"] .st-key-btn_sec_contexto button:hover,
    section[data-testid="stSidebar"] .st-key-btn_sec_orden button:hover,
    section[data-testid="stSidebar"] .st-key-btn_sec_exogenas button:hover {{
        color: {BLUE} !important;
    }}
    /* Estilo para botones primarios del generador */
    button[kind="primary"] {{
        background-color: {BLUE} !important;
        color: {WHITE} !important;
        border-radius: 6px !important;
        font-weight: 600 !important;
    }}
    button[kind="primary"]:hover {{
        background-color: {NAVY} !important;
    }}
    </style>
    """, unsafe_allow_html=True)

def pill(text, color_bg, color_fg):
    return f'<span style="background:{color_bg};color:{color_fg};font-size:11px;padding:2px 8px;border-radius:4px;font-weight:600;text-transform:uppercase;letter-spacing:0.4px;">{text}</span>'

def tag_ok():    return pill("CUMPLE", ESTADO_OK[0], ESTADO_OK[1])
def tag_warn():  return pill("REVISAR", ESTADO_WARN[0], ESTADO_WARN[1])
def tag_fail():  return pill("NO CUMPLE", ESTADO_FAIL[0], ESTADO_FAIL[1])
def tag_neutral(): return pill("N/A", ESTADO_NEUTRAL[0], ESTADO_NEUTRAL[1])

def card_kpi(title, value, subtitle="", accent=NAVY):
    sub = f'<p style="font-size:12px;color:{MUTED};margin:4px 0 0;line-height:1.3;">{subtitle}</p>' if subtitle else ''
    return f"""
    <div style="background:{WHITE};border:1px solid {BORDER};border-radius:8px;padding:14px 16px;height:100%;box-sizing:border-box;">
        <p style="font-size:10px;color:{LTGRAY};margin:0 0 6px;text-transform:uppercase;letter-spacing:0.6px;font-weight:600;">{title}</p>
        <p style="font-size:18px;font-weight:700;color:{accent};margin:0;line-height:1.2;">{value}</p>
        {sub}
    </div>
    """

def card_metric(label, value, color=TEXT):
    return f"""
    <div style="background:{WHITE};border:1px solid {BORDER};border-radius:6px;padding:12px 14px;">
        <p style="font-size:10px;color:{LTGRAY};margin:0 0 4px;text-transform:uppercase;letter-spacing:0.6px;font-weight:600;">{label}</p>
        <p style="font-size:20px;font-weight:700;color:{color};margin:0;">{value}</p>
    </div>
    """

def divider():
    return f"<div style='height:1px;background:{BORDER};margin:16px 0;'></div>"

def section_title(text):
    return f"<p style='font-size:13px;font-weight:700;color:{NAVY};margin:0 0 12px;text-transform:uppercase;letter-spacing:0.5px;'>{text}</p>"

def obtener_bandera_pais(pais, ancho=20):
    if not pais:
        return ""
    pais_str = str(pais).strip()
    codigo = BANDERAS_PAISES.get(pais_str)
    if not codigo:
        codigo = BANDERAS_PAISES.get(pais_str.lower())
    if not codigo:
        return ""
    return (f'<img src="https://flagcdn.com/w40/{codigo}.png" width="{ancho}" '
            f'style="vertical-align:middle;border-radius:2px;margin-right:6px;box-shadow:0 0 0 1px {BORDER};" '
            f'alt="{pais_str}">')

# =============================================================================
# PREFERENCIAS DEL SIDEBAR (persistentes entre sesiones)
# =============================================================================
PREFS_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".sidebar_prefs.json")

CLAVES_PREFS_SIDEBAR = [
    "sec_contexto", "sec_orden", "sec_exogenas",
    "criterio_ordenamiento",
    "filtro_ljung", "filtro_jarque", "filtro_hetero",
    "filtro_favoritos", "nav_sticky",
]

def cargar_prefs_sidebar():
    try:
        with open(PREFS_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}

def guardar_prefs_sidebar():
    prefs = {clave: st.session_state.get(clave) for clave in CLAVES_PREFS_SIDEBAR}
    try:
        with open(PREFS_PATH, "w", encoding="utf-8") as f:
            json.dump(prefs, f)
    except Exception:
        pass

def encabezado_colapsable(titulo, key):
    if key not in st.session_state:
        st.session_state[key] = True
    abierto = st.session_state[key]
    icono = "-" if abierto else "+"
    if st.button(f"{icono}  {titulo}", key=f"btn_{key}", use_container_width=True):
        st.session_state[key] = not abierto
        guardar_prefs_sidebar()
        st.rerun()
    return st.session_state[key]

# =============================================================================
# PARSER
# =============================================================================
def convertir_fecha(serie):
    if serie is None or len(serie) == 0:
        return serie
    if pd.api.types.is_datetime64_any_dtype(serie):
        return serie
    try:
        numeric = pd.to_numeric(serie, errors='coerce')
        if numeric.notna().sum() > 0 and numeric.dropna().min() > 30000:
            result = pd.to_datetime(numeric, unit='D', origin='1899-12-30', errors='coerce')
            if result.notna().sum() > 0:
                return result
    except:
        pass
    return pd.to_datetime(serie, errors='coerce')

def leer_meta_embebida(file, prefix="sarimax_meta"):
    try:
        file.seek(0)
        wb = openpyxl.load_workbook(file, read_only=True)
        props = wb.custom_doc_props
        n_prop_name = f"{prefix}_n"
        if n_prop_name not in props.names:
            return None
        n_partes = int(props[n_prop_name].value)
        partes = [props[f"{prefix}_{idx:02d}"].value for idx in range(1, n_partes + 1)]
        return json.loads("".join(partes))
    except Exception:
        return None
    finally:
        file.seek(0)

def parsear_excel(file):
    xls = pd.ExcelFile(file)
    modelos = {}
    for sheet_name in xls.sheet_names:
        try:
            df_raw = pd.read_excel(xls, sheet_name=sheet_name, header=None)
        except Exception:
            continue
        if len(df_raw) < 2:
            continue
        headers = [str(v).strip() if pd.notna(v) else "" for v in df_raw.iloc[1].values]
        col_map = {}
        for idx, name in enumerate(headers):
            if name:
                col_map.setdefault(name, []).append(idx)
        modelo = {"nombre": sheet_name}
        exog_cols, exog_names = [], set()
        for idx, name in enumerate(headers):
            nu = name.upper()
            if nu in ['FECHA', 'BASE', 'ADVERSO', 'OPTIMISTA'] or nu.startswith('FWL'):
                continue
            if nu.endswith(('_BASE', '_ADVERSO', '_OPTIMISTA')):
                exog_cols.append(name)
                base_name = name
                for suffix in ['_BASE', '_ADVERSO', '_OPTIMISTA']:
                    if base_name.upper().endswith(suffix):
                        base_name = base_name[:-len(suffix)]
                        break
                exog_names.add(base_name)
        endogena_cols = ['BASE', 'ADVERSO', 'OPTIMISTA']
        cols_seccion1 = ['fecha'] + endogena_cols + exog_cols
        if 'FWL_BASE' in col_map:
            cols_seccion1.extend(['FWL_BASE', 'FWL_ADVERSO', 'FWL_OPTIMISTA'])
        cols_seccion1 = [c for c in cols_seccion1 if c in col_map]
        idx_seccion1 = [col_map[c][0] for c in cols_seccion1]
        df_seccion1 = df_raw.iloc[2:, idx_seccion1].copy()
        df_seccion1.columns = cols_seccion1
        df_seccion1 = df_seccion1.dropna(how='all').reset_index(drop=True)
        if 'fecha' in df_seccion1.columns:
            df_seccion1['fecha'] = convertir_fecha(df_seccion1['fecha'])
        modelo['fecha_endogena'] = df_seccion1
        modelo['endogenas_cols'] = endogena_cols
        modelo['exogenas_cols'] = exog_cols
        modelo['exogenas_nombres'] = sorted(list(exog_names))
        if exog_cols:
            modelo['exogenas'] = df_seccion1[['fecha'] + exog_cols] if 'fecha' in df_seccion1.columns else df_seccion1[exog_cols]
        else:
            modelo['exogenas'] = None
        fwl_cols = [c for c in ['FWL_BASE', 'FWL_ADVERSO', 'FWL_OPTIMISTA'] if c in col_map]
        if fwl_cols and 'fecha' in col_map:
            idx_fwl = [col_map['fecha'][0]] + [col_map[c][0] for c in fwl_cols]
            df_fwl = df_raw.iloc[2:, idx_fwl].copy()
            df_fwl.columns = ['fecha'] + fwl_cols
            df_fwl = df_fwl.dropna(how='all').reset_index(drop=True)
            df_fwl['fecha'] = convertir_fecha(df_fwl['fecha'])
            df_fwl = df_fwl.dropna(subset=['FWL_BASE']).reset_index(drop=True)
            modelo['fwl_12m'] = df_fwl
        else:
            modelo['fwl_12m'] = None
        if 'Ano' in col_map and 'Escenario' in col_map and 'Factor FWL' in col_map:
            idx_anual = [col_map['Ano'][0], col_map['Escenario'][0], col_map['Factor FWL'][0]]
            df_anual = df_raw.iloc[2:, idx_anual].copy()
            df_anual.columns = ['Ano', 'Escenario', 'Factor FWL']
            df_anual = df_anual.dropna(how='all').reset_index(drop=True)
            modelo['fwl_anual'] = df_anual
        else:
            modelo['fwl_anual'] = None
        if 'Obs' in col_map and 'Residuo' in col_map:
            idx_res = [col_map['Obs'][0], col_map['Residuo'][0]]
            df_res = df_raw.iloc[2:, idx_res].copy()
            df_res.columns = ['Obs', 'Residuo']
            df_res = df_res.dropna(how='all').reset_index(drop=True)
            modelo['residuos_ind'] = df_res
        else:
            modelo['residuos_ind'] = None
        if 'Estadistico' in col_map and 'Valor' in col_map:
            idx_est = col_map['Estadistico'][0]
            idx_val = col_map['Valor'][0]
            df_resumen = df_raw.iloc[2:, [idx_est, idx_val]].copy()
            df_resumen.columns = ['Estadistico', 'Valor']
            df_resumen = df_resumen.dropna(how='all').reset_index(drop=True)
            modelo['resumen_residuos'] = df_resumen
        else:
            modelo['resumen_residuos'] = None
        if 'Variable' in col_map and 'Coeficiente' in col_map and 'P_value' in col_map:
            idx_var = col_map['Variable'][0]
            idx_coef = col_map['Coeficiente'][0]
            idx_pval = col_map['P_value'][0]
            df_coef = df_raw.iloc[2:, [idx_var, idx_coef, idx_pval]].copy()
            df_coef.columns = ['Variable', 'Coeficiente', 'P_value']
            df_coef = df_coef.dropna(how='all').reset_index(drop=True)
            modelo['coeficientes'] = df_coef
        else:
            modelo['coeficientes'] = None
        if 'Prueba' in col_map and 'Estadistico' in col_map and 'P_value' in col_map:
            idx_prueba = col_map['Prueba'][0]
            idx_est = col_map['Estadistico'][-1]
            idx_pval = col_map['P_value'][-1]
            df_pruebas = df_raw.iloc[2:, [idx_prueba, idx_est, idx_pval]].copy()
            df_pruebas.columns = ['Prueba', 'Estadistico', 'P_value']
            df_pruebas = df_pruebas.dropna(how='all').reset_index(drop=True)
            modelo['pruebas'] = df_pruebas
        else:
            modelo['pruebas'] = None
        modelo['observaciones'] = len(modelo['fecha_endogena'].dropna(how='all')) if modelo['fecha_endogena'] is not None and not modelo['fecha_endogena'].empty else 0
        modelos[sheet_name] = modelo
    return modelos

# =============================================================================
# UTILIDADES
# =============================================================================
def contar_pruebas_aprobadas(pruebas_df):
    if pruebas_df is None or pruebas_df.empty:
        return 0, 3
    aprobadas = 0
    for _, row in pruebas_df.iterrows():
        prueba = str(row.get('Prueba', '')).lower()
        p_val = row.get('P_value', None)
        if p_val is None or pd.isna(p_val):
            continue
        try: p_val = float(p_val)
        except: continue
        if 'ljung' in prueba or 'box' in prueba:
            if p_val > 0.05: aprobadas += 1
        elif 'jarque' in prueba or 'bera' in prueba:
            if p_val > 0.05: aprobadas += 1
        elif 'hetero' in prueba or 'arch' in prueba:
            if p_val > 0.05: aprobadas += 1
    return aprobadas, 3

def clasificar_variable(var_name):
    var_lower = str(var_name).lower()
    if var_lower.startswith('ar.'): return 'AR'
    elif var_lower.startswith('ma.'): return 'MA'
    elif var_lower == 'intercept': return 'Exogena'
    elif var_lower.startswith('var_'): return 'Exogena'
    elif var_lower == 'sigma2': return 'Varianza'
    return 'Otro'

def contar_ar_ma(coeficientes_df):
    if coeficientes_df is None or coeficientes_df.empty:
        return 0, 0
    ar_count = 0
    ma_count = 0
    for _, row in coeficientes_df.iterrows():
        var = str(row.get('Variable', '')).lower()
        if var.startswith('ar.'):
            ar_count += 1
        elif var.startswith('ma.'):
            ma_count += 1
    return ar_count, ma_count

def generar_campana_normal(residuos, media, std):
    if std == 0 or len(residuos) == 0:
        return [], []
    x = np.linspace(min(residuos), max(residuos), 100)
    y = (1 / (std * np.sqrt(2 * np.pi))) * np.exp(-0.5 * ((x - media) / std) ** 2)
    return x, y

def obtener_significancia_exogenas(coeficientes_df, exogenas_lista):
    if coeficientes_df is None or coeficientes_df.empty:
        return []
    resultados = []
    for exog in exogenas_lista:
        p_val = None
        for _, row in coeficientes_df.iterrows():
            var = str(row.get('Variable', ''))
            if exog in var:
                p_val = row.get('P_value', None)
                break
        if p_val is not None:
            try: p_val = float(p_val)
            except: p_val = None
        if p_val is None:
            resultados.append((exog, None, "Sin datos"))
        elif p_val < 0.05:
            resultados.append((exog, p_val, "Significativa"))
        elif p_val < 0.10:
            resultados.append((exog, p_val, "Marginal"))
        else:
            resultados.append((exog, p_val, "No significativa"))
    return resultados

def calcular_fwl_ponderado(fwl_df, pesos):
    if fwl_df is None or fwl_df.empty:
        return None
    fecha_col = None
    for c in fwl_df.columns:
        if 'fecha' in str(c).lower():
            fecha_col = c
            break
    base_col = [c for c in fwl_df.columns if 'FWL_BASE' in str(c).upper()]
    adv_col = [c for c in fwl_df.columns if 'FWL_ADVERSO' in str(c).upper() or 'FWL_ADVERSA' in str(c).upper()]
    opt_col = [c for c in fwl_df.columns if 'FWL_OPTIMISTA' in str(c).upper()]
    if not base_col or not adv_col or not opt_col:
        return None
    base_col, adv_col, opt_col = base_col[0], adv_col[0], opt_col[0]
    df = fwl_df.copy()
    df['FWL_Ponderado'] = (
        df[base_col].astype(float) * pesos['base'] +
        df[adv_col].astype(float) * pesos['adverso'] +
        df[opt_col].astype(float) * pesos['optimista']
    )
    return df

def resumen_fwl(fwl_df):
    if fwl_df is None or fwl_df.empty or 'FWL_Ponderado' not in fwl_df.columns:
        return {}
    vals = fwl_df['FWL_Ponderado'].dropna().astype(float)
    if len(vals) == 0:
        return {}
    return {'promedio': vals.mean(), 'maximo': vals.max(), 'minimo': vals.min(), 'volatilidad': vals.std()}

# =============================================================================
# SCORE SYSTEM (A-D Grading)
# =============================================================================
def calcular_score(p_val, prueba_nombre=""):
    if p_val is None or pd.isna(p_val):
        return 'N/A', ESTADO_NEUTRAL
    try:
        p_val = float(p_val)
    except:
        return 'N/A', ESTADO_NEUTRAL
    prueba_lower = str(prueba_nombre).lower()
    if 'jarque' in prueba_lower or 'bera' in prueba_lower:
        if p_val > 0.10:
            return 'A', SCORE_COLORS['A']
        elif p_val > 0.05:
            return 'B', SCORE_COLORS['B']
        elif p_val > 0.01:
            return 'C', SCORE_COLORS['C']
        else:
            return 'D', SCORE_COLORS['D']
    else:
        if p_val > 0.10:
            return 'A', SCORE_COLORS['A']
        elif p_val > 0.05:
            return 'B', SCORE_COLORS['B']
        elif p_val > 0.01:
            return 'C', SCORE_COLORS['C']
        else:
            return 'D', SCORE_COLORS['D']

def obtener_score_prueba(pruebas_df, nombre_prueba):
    if pruebas_df is None or pruebas_df.empty:
        return 'N/A', None
    nombre_lower = nombre_prueba.lower()
    for _, row in pruebas_df.iterrows():
        prueba = str(row.get('Prueba', '')).lower()
        p_val = row.get('P_value', None)
        if 'ljung' in nombre_lower and ('ljung' in prueba or 'box' in prueba):
            return calcular_score(p_val, prueba)
        elif 'jarque' in nombre_lower and ('jarque' in prueba or 'bera' in prueba):
            return calcular_score(p_val, prueba)
        elif 'hetero' in nombre_lower and ('hetero' in prueba or 'arch' in prueba):
            return calcular_score(p_val, prueba)
    return 'N/A', None

def obtener_scores_modelo(pruebas_df):
    scores = {}
    scores['ljung_box'] = obtener_score_prueba(pruebas_df, 'ljung')
    scores['jarque_bera'] = obtener_score_prueba(pruebas_df, 'jarque')
    scores['heterocedasticidad'] = obtener_score_prueba(pruebas_df, 'hetero')
    return scores

def calcular_score_global(pruebas_df):
    scores = obtener_scores_modelo(pruebas_df)
    acumulado, peso_total = 0.0, 0.0
    detalle = {}
    for clave, peso in PESOS_SCORE_GLOBAL.items():
        letra, _ = scores.get(clave, ('N/A', None))
        detalle[clave] = letra
        if letra in SCORE_NUM_MAP:
            acumulado += SCORE_NUM_MAP[letra] * peso
            peso_total += peso
    if peso_total == 0:
        return None, detalle
    return round(acumulado / peso_total, 1), detalle

def clasificar_score_global(score):
    if score is None:
        return "N/A", GRAY, ESTADO_NEUTRAL[0]
    if score >= 7:
        return "BUENO", GREEN, ESTADO_OK[0]
    elif score >= 5:
        return "REGULAR", "#B8860B", ESTADO_WARN[0]
    else:
        return "DEFICIENTE", RED, ESTADO_FAIL[0]

def score_global_badge(score, tamano="12px"):
    etiqueta, color, bg = clasificar_score_global(score)
    valor = f"{score:.1f}/10" if score is not None else "N/A"
    return (f'<span style="background:{bg};color:{color};font-size:{tamano};padding:3px 10px;'
            f'border-radius:4px;font-weight:700;">{valor} - {etiqueta}</span>')

def interpretar_prueba(nombre_prueba, p_val, score):
    if score == 'N/A' or p_val is None or (isinstance(p_val, float) and pd.isna(p_val)):
        return "Sin datos disponibles para esta prueba."
    p_str = fmt_pvalor(p_val)
    nl = str(nombre_prueba).lower()
    if 'ljung' in nl or 'box' in nl:
        base = "autocorrelacion en los residuos"
    elif 'jarque' in nl or 'bera' in nl:
        base = "no-normalidad en los residuos"
    else:
        base = "heterocedasticidad (varianza no constante)"
    if score in ['A', 'B']:
        return f"p = {p_str} - sin evidencia significativa de {base}."
    elif score == 'C':
        return f"p = {p_str} - evidencia debil de {base}."
    else:
        return f"p = {p_str} - evidencia de {base}."

def score_badge(score, bg_color, fg_color):
    return f'<span style="background:{bg_color};color:{fg_color};font-size:12px;padding:3px 10px;border-radius:4px;font-weight:700;">{score}</span>'

def render_leyenda_scores():
    html = f"""
    <div style="background:{WHITE};border:1px solid {BORDER};border-radius:8px;padding:16px;margin:16px 0;">
        <p style="font-size:12px;font-weight:700;color:{NAVY};margin:0 0 12px;text-transform:uppercase;letter-spacing:0.5px;">Clasificacion de Scores</p>
        <table style="width:100%;border-collapse:collapse;font-size:12px;">
            <thead>
                <tr style="border-bottom:1px solid {BORDER};">
                    <th style="text-align:left;padding:6px 8px;font-weight:600;color:{NAVY};">Score</th>
                    <th style="text-align:left;padding:6px 8px;font-weight:600;color:{NAVY};">Rango p-valor</th>
                    <th style="text-align:left;padding:6px 8px;font-weight:600;color:{NAVY};">Ljung-Box / Heterocedasticidad</th>
                    <th style="text-align:left;padding:6px 8px;font-weight:600;color:{NAVY};">Jarque-Bera</th>
                </tr>
            </thead>
            <tbody>
                <tr style="border-bottom:1px solid {BORDER};">
                    <td style="padding:6px 8px;">{score_badge('A', SCORE_COLORS['A'][0], SCORE_COLORS['A'][1])}</td>
                    <td style="padding:6px 8px;color:{TEXT};">p &gt; 0.10</td>
                    <td style="padding:6px 8px;color:{GREEN};font-weight:600;">Sin autocorrelacion / heterocedasticidad</td>
                    <td style="padding:6px 8px;color:{GREEN};font-weight:600;">Residuos normales</td>
                </tr>
                <tr style="border-bottom:1px solid {BORDER};">
                    <td style="padding:6px 8px;">{score_badge('B', SCORE_COLORS['B'][0], SCORE_COLORS['B'][1])}</td>
                    <td style="padding:6px 8px;color:{TEXT};">0.05 &lt; p &le; 0.10</td>
                    <td style="padding:6px 8px;color:{BLUE};font-weight:600;">Sin evidencia significativa</td>
                    <td style="padding:6px 8px;color:{BLUE};font-weight:600;">Sin evidencia significativa</td>
                </tr>
                <tr style="border-bottom:1px solid {BORDER};">
                    <td style="padding:6px 8px;">{score_badge('C', SCORE_COLORS['C'][0], SCORE_COLORS['C'][1])}</td>
                    <td style="padding:6px 8px;color:{TEXT};">0.01 &lt; p &le; 0.05</td>
                    <td style="padding:6px 8px;color:#B8860B;font-weight:600;">Evidencia debil de problema</td>
                    <td style="padding:6px 8px;color:#B8860B;font-weight:600;">Evidencia debil de no-normalidad</td>
                </tr>
                <tr>
                    <td style="padding:6px 8px;">{score_badge('D', SCORE_COLORS['D'][0], SCORE_COLORS['D'][1])}</td>
                    <td style="padding:6px 8px;color:{TEXT};">p &le; 0.01</td>
                    <td style="padding:6px 8px;color:{RED};font-weight:600;">Autocorrelacion / heterocedasticidad presente</td>
                    <td style="padding:6px 8px;color:{RED};font-weight:600;">Residuos no normales</td>
                </tr>
            </tbody>
        </table>
        <p style="font-size:10px;color:{MUTED};margin:10px 0 0;line-height:1.4;">
            <b>Ljung-Box:</b> Prueba de autocorrelacion en residuos. H0: no hay autocorrelacion.<br>
            <b>Jarque-Bera:</b> Prueba de normalidad. H0: los residuos siguen distribucion normal.<br>
            <b>Heterocedasticidad (ARCH/LM):</b> Prueba de varianza constante. H0: varianza homogenea.
        </p>
    </div>
    """
    return html

# =============================================================================
# METADATA
# =============================================================================
def extraer_kpis_meta(meta):
    if not meta:
        return {}
    return {
        'pais': PAISES_MAP.get(str(meta.get('pais', '')).lower().strip(), meta.get('pais', 'N/A')),
        'cartera': CARTERAS_MAP.get(str(meta.get('cartera', '')).lower().strip(), meta.get('cartera', 'N/A')),
        'tipo_endogena': meta.get('motor_tipo_endogena', 'N/A'),
        'modo_endogena': meta.get('generador_modo_endogena', 'N/A'),
        'ventana_mm': meta.get('generador_ventana_mm', 'N/A'),
        'vif_max': meta.get('motor_vif_max', 'N/A'),
        'fwl_min': meta.get('motor_fwl_factor_min', '?'),
        'fwl_max': meta.get('motor_fwl_factor_max', '?'),
        'max_exog': meta.get('motor_max_exog_por_modelo', 'N/A'),
        'top_exportar': meta.get('motor_top_exportar', 'N/A'),
    }

# =============================================================================
# PLOTS
# =============================================================================
def aplicar_tema_plotly(fig):
    fig.update_layout(
        font=dict(family="Inter, Arial, sans-serif", size=12, color=TEXT),
        plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
        legend=dict(bgcolor="rgba(0,0,0,0)"),
        margin=dict(t=40, l=10, r=10, b=10),
    )
    fig.update_xaxes(showgrid=True, gridcolor=BORDER, zeroline=False, linecolor=BORDER)
    fig.update_yaxes(showgrid=True, gridcolor=BORDER, zeroline=False, linecolor=BORDER)
    if fig.layout.title and fig.layout.title.text:
        fig.update_layout(title=dict(font=dict(size=14, color=NAVY)))
    return fig

def fig_predicciones(df_end, endogena_cols, exog_df, exog_sel, modelo_nombre):
    fig = go.Figure()
    fecha_col = 'fecha' if 'fecha' in df_end.columns else df_end.columns[0]
    for col in endogena_cols:
        if col not in df_end.columns:
            continue
        col_str = str(col).upper()
        if col_str == 'BASE':
            fig.add_trace(go.Scatter(x=df_end[fecha_col], y=df_end[col], mode='lines', name='Base', line=dict(color=BLUE, width=2)))
        elif col_str in ['ADVERSO', 'ADVERSA']:
            fig.add_trace(go.Scatter(x=df_end[fecha_col], y=df_end[col], mode='lines', name='Adverso', line=dict(color=RED, width=2, dash='dash')))
        elif col_str == 'OPTIMISTA':
            fig.add_trace(go.Scatter(x=df_end[fecha_col], y=df_end[col], mode='lines', name='Optimista', line=dict(color=GREEN, width=2, dash='dot')))
    if exog_df is not None and exog_sel:
        for ex in exog_sel:
            for suffix in ['_BASE', '_ADVERSO', '_OPTIMISTA']:
                col_name = ex + suffix
                if col_name in exog_df.columns:
                    x_vals = exog_df[fecha_col] if fecha_col in exog_df.columns else exog_df.index
                    fig.add_trace(go.Scatter(x=x_vals, y=exog_df[col_name], mode='lines', name=f'{ex}{suffix}', line=dict(width=1.2), yaxis='y2'))
        fig.update_layout(yaxis2=dict(title='Exogenas', overlaying='y', side='right'))
    fig.update_layout(title=f"Predicciones - {modelo_nombre}", xaxis_title="Fecha", yaxis_title="Valor",
                      legend=dict(orientation='v', yanchor='top', y=1, xanchor='left', x=1.05), hovermode='x unified')
    return aplicar_tema_plotly(fig)

def fig_fwl_12m(df_fwl):
    fig = go.Figure()
    fecha_col = None
    for c in df_fwl.columns:
        if 'fecha' in str(c).lower():
            fecha_col = c
            break
    if fecha_col is None:
        fecha_col = df_fwl.columns[0]
    for col in df_fwl.columns:
        col_str = str(col).upper()
        if 'FWL_BASE' in col_str:
            fig.add_trace(go.Scatter(x=df_fwl[fecha_col], y=df_fwl[col], mode='lines', name='Base', line=dict(color=BLUE, width=2)))
        elif 'FWL_ADVERSO' in col_str or 'FWL_ADVERSA' in col_str:
            fig.add_trace(go.Scatter(x=df_fwl[fecha_col], y=df_fwl[col], mode='lines', name='Adverso', line=dict(color=RED, width=2, dash='dash')))
        elif 'FWL_OPTIMISTA' in col_str:
            fig.add_trace(go.Scatter(x=df_fwl[fecha_col], y=df_fwl[col], mode='lines', name='Optimista', line=dict(color=GREEN, width=2, dash='dot')))
    fig.update_layout(title="Factor FWL a 12 Meses", xaxis_title="Fecha", yaxis_title="FWL",
                      legend=dict(orientation='v', yanchor='top', y=1, xanchor='left', x=1.05), hovermode='x unified')
    return aplicar_tema_plotly(fig)

def fig_fwl_ponderado(df_pond):
    fecha_col = None
    for c in df_pond.columns:
        if 'fecha' in str(c).lower():
            fecha_col = c
            break
    if fecha_col is None:
        fecha_col = df_pond.columns[0]
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=df_pond[fecha_col], y=df_pond['FWL_Ponderado'], mode='lines', name='FWL Ponderado',
                              fill='tozeroy', line=dict(color=BLUE, width=2), fillcolor='rgba(30,90,168,0.10)'))
    fig.update_layout(title="Factor FWL Ponderado", xaxis_title="Fecha", yaxis_title="FWL Ponderado")
    return aplicar_tema_plotly(fig)

def fig_histograma_residuos(vals, media, std):
    fig = go.Figure()
    fig.add_trace(go.Histogram(x=vals, nbinsx=20, marker_color=BLUE, opacity=0.7, name='Residuos'))
    x_norm, y_norm = generar_campana_normal(vals, media, std)
    if len(x_norm) > 0:
        bin_width = (vals.max() - vals.min()) / 20 if vals.max() != vals.min() else 1
        y_norm_scaled = y_norm * len(vals) * bin_width
        fig.add_trace(go.Scatter(x=x_norm, y=y_norm_scaled, mode='lines', name='Normal teorica', line=dict(color=RED, width=2)))
    fig.update_layout(xaxis_title="Residuos", yaxis_title="Frecuencia")
    return aplicar_tema_plotly(fig)

def fig_barras_coeficientes(df_coef):
    df = df_coef.copy()
    df['abs'] = df['Coeficiente'].abs()
    df = df.sort_values('abs', ascending=True)
    colors = [GREEN if c >= 0 else RED for c in df['Coeficiente']]
    fig = go.Figure()
    fig.add_trace(go.Bar(y=df['Variable'], x=df['Coeficiente'], orientation='h', marker_color=colors,
                          text=df['Coeficiente'].round(4), textposition='outside'))
    fig.update_layout(title="Coeficientes del Modelo", xaxis_title="Valor", yaxis_title="Variable", showlegend=False)
    return aplicar_tema_plotly(fig)

# =============================================================================
# DIAGNOSTICOS
# =============================================================================
def limpiar_nombre_prueba(nombre):
    nl = str(nombre).lower()
    if 'arch' in nl: return "Heterocedasticidad"
    if 'ljung' in nl or 'box' in nl: return "Ljung-Box"
    if 'jarque' in nl or 'bera' in nl: return "Jarque-Bera"
    return str(nombre)

def evaluar_prueba(prueba, p_val):
    score, (bg, fg) = calcular_score(p_val, prueba)
    if score == 'N/A':
        return "N/A", ESTADO_NEUTRAL, score, bg, fg
    if score == 'A':
        estado = "CUMPLE"
    elif score == 'B':
        estado = "CUMPLE"
    elif score == 'C':
        estado = "REVISAR"
    else:
        estado = "NO CUMPLE"
    return estado, (bg, fg, estado), score, bg, fg

def render_diagnosticos_corporativo(pruebas_df, mostrar_detalle_tecnico=True):
    if pruebas_df is None or pruebas_df.empty:
        st.info("No hay datos de pruebas estadisticas.")
        return
    df = pruebas_df.copy()
    df['Prueba'] = df['Prueba'].apply(limpiar_nombre_prueba)
    st.markdown(section_title("Resumen de Diagnosticos"), unsafe_allow_html=True)
    filas = []
    for _, row in df.iterrows():
        prueba = row['Prueba']
        p_val = row['P_value']
        estado, _, score, bg, fg = evaluar_prueba(prueba, p_val)
        filas.append({
            'Diagnostico': prueba,
            'Score': score,
            'ScoreBg': bg,
            'ScoreFg': fg,
            'P-valor': p_val,
            'Estadistico': row.get('Estadistico', '-'),
        })
    cols = st.columns(len(filas))
    for i, f in enumerate(filas):
        with cols[i]:
            st.markdown(f"""
            <div style="background:{f['ScoreBg']};border:1px solid {BORDER};border-radius:6px;padding:16px;text-align:center;">
                <p style="font-size:10px;color:{MUTED};margin:0 0 6px;text-transform:uppercase;letter-spacing:0.5px;font-weight:600;">{f['Diagnostico']}</p>
                <p style="font-size:28px;font-weight:700;color:{f['ScoreFg']};margin:0;">{f['Score']}</p>
                <p style="font-size:11px;color:{MUTED};margin:4px 0 0;">p = {fmt_pvalor(f['P-valor'])}</p>
            </div>
            """, unsafe_allow_html=True)
            with st.expander("Ver interpretacion"):
                st.markdown(
                    f"<p style='font-size:12px;color:{TEXT};margin:0;'>{interpretar_prueba(f['Diagnostico'], f['P-valor'], f['Score'])}</p>",
                    unsafe_allow_html=True
                )
    if not mostrar_detalle_tecnico:
        return
    st.markdown("<div style='height:8px;'></div>", unsafe_allow_html=True)
    st.markdown(section_title("Detalle Tecnico"), unsafe_allow_html=True)
    df_tec = pd.DataFrame(filas)[['Diagnostico', 'Estadistico', 'P-valor', 'Score']]
    df_tec['P-valor'] = df_tec['P-valor'].apply(fmt_pvalor)
    df_tec['Estadistico'] = df_tec['Estadistico'].apply(fmt_pvalor)
    def color_score(val):
        if val == "A": return f"color: {GREEN}; font-weight: 700;"
        elif val == "B": return f"color: {BLUE}; font-weight: 700;"
        elif val == "C": return f"color: #B8860B; font-weight: 700;"
        elif val == "D": return f"color: {RED}; font-weight: 700;"
        return ""
    styler = df_tec.style
    if hasattr(styler, "map"):
        styler = styler.map(color_score, subset=['Score'])
    else:
        styler = styler.applymap(color_score, subset=['Score'])
    st.dataframe(styler, use_container_width=True, hide_index=True)

def fmt_pvalor(v):
    try:
        if pd.isna(v): return "-"
        vf = float(v)
        return f"{vf:.4f}" if vf >= 0.001 else f"{vf:.2e}"
    except: return str(v)

def render_metricas_diagnostico(pruebas_df):
    if pruebas_df is None or pruebas_df.empty:
        return
    scores = obtener_scores_modelo(pruebas_df)
    score_global, _ = calcular_score_global(pruebas_df)
    etiqueta_g, color_g, _ = clasificar_score_global(score_global)
    c0, c1, c2, c3 = st.columns(4)
    with c0:
        valor_g = f"{score_global:.1f}/10" if score_global is not None else "N/A"
        st.markdown(card_kpi("Score Global", valor_g, etiqueta_g, accent=color_g), unsafe_allow_html=True)
    with c1:
        score, (bg, fg) = scores.get('ljung_box', ('N/A', ESTADO_NEUTRAL))
        st.markdown(card_metric("Ljung-Box (Score)", score, fg if score != 'N/A' else TEXT), unsafe_allow_html=True)
    with c2:
        score, (bg, fg) = scores.get('jarque_bera', ('N/A', ESTADO_NEUTRAL))
        st.markdown(card_metric("Jarque-Bera (Score)", score, fg if score != 'N/A' else TEXT), unsafe_allow_html=True)
    with c3:
        score, (bg, fg) = scores.get('heterocedasticidad', ('N/A', ESTADO_NEUTRAL))
        st.markdown(card_metric("Heterocedast.", score, fg if score != 'N/A' else TEXT), unsafe_allow_html=True)

def es_favorito(nombre):
    return nombre in st.session_state.get("favoritos", set())

def alternar_favorito(nombre):
    if nombre in st.session_state.favoritos:
        st.session_state.favoritos.discard(nombre)
    else:
        st.session_state.favoritos.add(nombre)

def boton_favorito(nombre, key_suffix=""):
    activo = es_favorito(nombre)
    label = "Favorito" if activo else "Marcar favorito"
    if st.button(label, key=f"fav_{key_suffix}_{nombre}", use_container_width=True):
        alternar_favorito(nombre)
        st.rerun()

def construir_opciones_modelos():
    criterio = st.session_state.get("criterio_ordenamiento", "Pruebas aprobadas ↓")
    filtro_ljung = st.session_state.get("filtro_ljung", "Todos")
    filtro_jarque = st.session_state.get("filtro_jarque", "Todos")
    filtro_hetero = st.session_state.get("filtro_hetero", "Todos")
    filtro_favoritos = st.session_state.get("filtro_favoritos", False)
    modelos_con_pruebas = []
    for nombre, datos in st.session_state.modelos_data.items():
        pruebas = datos.get('pruebas')
        apr, tot = contar_pruebas_aprobadas(pruebas)
        scores = obtener_scores_modelo(pruebas)
        pasa_filtro = True
        if filtro_ljung != "Todos":
            score_ljung, _ = scores.get('ljung_box', ('N/A', None))
            if filtro_ljung == "A o B (Cumple)" and score_ljung not in ['A', 'B']:
                pasa_filtro = False
            elif filtro_ljung == "A, B o C" and score_ljung not in ['A', 'B', 'C']:
                pasa_filtro = False
            elif filtro_ljung == "Solo A" and score_ljung != 'A':
                pasa_filtro = False
        if filtro_jarque != "Todos":
            score_jarque, _ = scores.get('jarque_bera', ('N/A', None))
            if filtro_jarque == "A o B (Cumple)" and score_jarque not in ['A', 'B']:
                pasa_filtro = False
            elif filtro_jarque == "A, B o C" and score_jarque not in ['A', 'B', 'C']:
                pasa_filtro = False
            elif filtro_jarque == "Solo A" and score_jarque != 'A':
                pasa_filtro = False
        if filtro_hetero != "Todos":
            score_hetero, _ = scores.get('heterocedasticidad', ('N/A', None))
            if filtro_hetero == "A o B (Cumple)" and score_hetero not in ['A', 'B']:
                pasa_filtro = False
            elif filtro_hetero == "A, B o C" and score_hetero not in ['A', 'B', 'C']:
                pasa_filtro = False
            elif filtro_hetero == "Solo A" and score_hetero != 'A':
                pasa_filtro = False
        if filtro_favoritos and nombre not in st.session_state.get("favoritos", set()):
            pasa_filtro = False
        if pasa_filtro:
            score_global, _ = calcular_score_global(pruebas)
            modelos_con_pruebas.append((nombre, apr, tot, scores, score_global))
    if criterio == "Nombre (A-Z)":
        modelos_con_pruebas.sort(key=lambda x: x[0])
    elif criterio == "Pruebas aprobadas ↓":
        modelos_con_pruebas.sort(key=lambda x: (-x[1], x[0]))
    elif criterio == "Pruebas aprobadas ↑":
        modelos_con_pruebas.sort(key=lambda x: (x[1], x[0]))
    elif criterio == "Score global ↓":
        modelos_con_pruebas.sort(key=lambda x: (-(x[4] if x[4] is not None else -1), x[0]))
    elif criterio == "Score global ↑":
        modelos_con_pruebas.sort(key=lambda x: ((x[4] if x[4] is not None else 999), x[0]))
    else:
        modelos_con_pruebas.sort(key=lambda x: (x[1], x[0]))
    modelos_list = [m[0] for m in modelos_con_pruebas]
    pruebas_dict = {m[0]: (m[1], m[2]) for m in modelos_con_pruebas}
    scores_dict = {m[0]: m[3] for m in modelos_con_pruebas}
    global_dict = {m[0]: m[4] for m in modelos_con_pruebas}
    return modelos_list, pruebas_dict, scores_dict, global_dict

def label_modelo(nombre, pruebas_dict, scores_dict=None, global_dict=None):
    apr, tot = pruebas_dict.get(nombre, (0, 3))
    prefijo = "[F] " if es_favorito(nombre) else ""
    label = f"{prefijo}{nombre}  ({apr}/{tot})"
    if scores_dict and nombre in scores_dict:
        scores = scores_dict[nombre]
        mini_scores = []
        for key in ['ljung_box', 'jarque_bera', 'heterocedasticidad']:
            score, _ = scores.get(key, ('N/A', None))
            if score != 'N/A':
                mini_scores.append(score)
        if mini_scores:
            label += f"  [{'|'.join(mini_scores)}]"
    if global_dict and nombre in global_dict and global_dict[nombre] is not None:
        label += f"  - {global_dict[nombre]:.1f}/10"
    return label

def render_columna_comparacion(nombre):
    datos = st.session_state.modelos_data.get(nombre, {})
    pruebas = datos.get('pruebas')
    score_global, _ = calcular_score_global(pruebas)
    scores = obtener_scores_modelo(pruebas)
    st.markdown(f"<p style='font-weight:700;color:{NAVY};font-size:13px;margin:0 0 6px;'>{nombre}</p>", unsafe_allow_html=True)
    st.markdown(score_global_badge(score_global), unsafe_allow_html=True)
    st.markdown("<div style='height:8px;'></div>", unsafe_allow_html=True)
    sc1, sc2, sc3 = st.columns(3)
    with sc1:
        s, _ = scores.get('ljung_box', ('N/A', ESTADO_NEUTRAL))
        st.markdown(card_metric("Ljung-Box", s), unsafe_allow_html=True)
    with sc2:
        s, _ = scores.get('jarque_bera', ('N/A', ESTADO_NEUTRAL))
        st.markdown(card_metric("Jarque-Bera", s), unsafe_allow_html=True)
    with sc3:
        s, _ = scores.get('heterocedasticidad', ('N/A', ESTADO_NEUTRAL))
        st.markdown(card_metric("Heterocedast.", s), unsafe_allow_html=True)
    st.markdown("<div style='height:8px;'></div>", unsafe_allow_html=True)
    df_fwl_comp = datos.get('fwl_12m')
    if df_fwl_comp is not None and not df_fwl_comp.empty:
        fig = fig_fwl_12m(df_fwl_comp)
        fig.update_layout(height=280, showlegend=False, title="Factor FWL a 12 Meses")
        st.plotly_chart(fig, use_container_width=True, key=f"cmp_fig_{nombre}")
    else:
        st.caption("Sin datos de FWL a 12 meses.")
    coefs = datos.get('coeficientes')
    ar_count, ma_count = contar_ar_ma(coefs) if coefs is not None else (0, 0)
    exogenas = datos.get('exogenas_nombres', [])
    sigs = obtener_significancia_exogenas(coefs, exogenas)
    sig_count = sum(1 for _, _, s in sigs if s == "Significativa")
    mc1, mc2 = st.columns(2)
    with mc1:
        st.markdown(card_metric("AR / MA", f"{ar_count} / {ma_count}", BLUE), unsafe_allow_html=True)
    with mc2:
        st.markdown(card_metric("Exog. significativas", f"{sig_count}/{len(exogenas)}", GREEN), unsafe_allow_html=True)

def fig_distribucion_scores(modelos_data):
    nombre_map = {'ljung_box': 'Ljung-Box', 'jarque_bera': 'Jarque-Bera', 'heterocedasticidad': 'Heterocedasticidad'}
    conteo = {label: {'A': 0, 'B': 0, 'C': 0, 'D': 0} for label in nombre_map.values()}
    for datos in modelos_data.values():
        scores = obtener_scores_modelo(datos.get('pruebas'))
        for clave, label in nombre_map.items():
            letra, _ = scores.get(clave, ('N/A', None))
            if letra in conteo[label]:
                conteo[label][letra] += 1
    fig = go.Figure()
    for letra in ['A', 'B', 'C', 'D']:
        color = SCORE_COLORS[letra][1]
        fig.add_trace(go.Bar(
            name=letra, x=list(nombre_map.values()),
            y=[conteo[label][letra] for label in nombre_map.values()],
            marker_color=color
        ))
    fig.update_layout(barmode='stack', title="Distribucion de Scores por Prueba", yaxis_title="Cantidad de modelos", legend_title_text="Score")
    return aplicar_tema_plotly(fig)

def render_resumen_ejecutivo():
    modelos_data = st.session_state.modelos_data
    filas = []
    for nombre, datos in modelos_data.items():
        pruebas = datos.get('pruebas')
        score_global, _ = calcular_score_global(pruebas)
        coefs = datos.get('coeficientes')
        ar_count, ma_count = contar_ar_ma(coefs) if coefs is not None else (0, 0)
        filas.append({'Modelo': nombre, 'Score': score_global, 'AR': ar_count, 'MA': ma_count})
    df_res = pd.DataFrame(filas)
    total = len(df_res)
    con_score = df_res.dropna(subset=['Score'])
    buenos = int((con_score['Score'] >= 7).sum()) if not con_score.empty else 0
    regulares = int(((con_score['Score'] >= 5) & (con_score['Score'] < 7)).sum()) if not con_score.empty else 0
    deficientes = int((con_score['Score'] < 5).sum()) if not con_score.empty else 0
    st.markdown(f"""
    <div style="margin-bottom:8px;">
        <p style="font-size:20px;font-weight:700;color:{NAVY};margin:0;">Resumen de la corrida</p>
        <p style="font-size:12px;color:{MUTED};margin:4px 0 0;">Vista general de los {total} modelos cargados en el archivo.</p>
    </div>
    <div style="height:1px;background:{BORDER};margin:12px 0 20px;"></div>
    """, unsafe_allow_html=True)
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(card_kpi("Total de modelos", str(total)), unsafe_allow_html=True)
    with c2:
        st.markdown(card_kpi("Buenos (score >= 7)", str(buenos), accent=GREEN), unsafe_allow_html=True)
    with c3:
        st.markdown(card_kpi("Regulares (5 - 7)", str(regulares), accent="#B8860B"), unsafe_allow_html=True)
    with c4:
        st.markdown(card_kpi("Deficientes (< 5)", str(deficientes), accent=RED), unsafe_allow_html=True)
    st.markdown(divider(), unsafe_allow_html=True)
    st.plotly_chart(fig_distribucion_scores(modelos_data), use_container_width=True)
    st.markdown(divider(), unsafe_allow_html=True)
    tcol1, tcol2 = st.columns(2)
    with tcol1:
        st.markdown(section_title("Top 5 - mejor score global"), unsafe_allow_html=True)
        top5 = con_score.sort_values('Score', ascending=False).head(5)[['Modelo', 'Score']]
        st.dataframe(top5, use_container_width=True, hide_index=True)
    with tcol2:
        st.markdown(section_title("Bottom 5 - peor score global"), unsafe_allow_html=True)
        bottom5 = con_score.sort_values('Score', ascending=True).head(5)[['Modelo', 'Score']]
        st.dataframe(bottom5, use_container_width=True, hide_index=True)
    st.markdown(divider(), unsafe_allow_html=True)
    st.markdown(section_title("Estructura promedio de los modelos"), unsafe_allow_html=True)
    ac1, ac2 = st.columns(2)
    with ac1:
        st.markdown(card_metric("Promedio terminos AR", f"{df_res['AR'].mean():.2f}" if total else "0", BLUE), unsafe_allow_html=True)
    with ac2:
        st.markdown(card_metric("Promedio terminos MA", f"{df_res['MA'].mean():.2f}" if total else "0", BLUE), unsafe_allow_html=True)
    st.markdown("<div style='height:24px;'></div>", unsafe_allow_html=True)
    bexp1, bexp2 = st.columns(2)
    with bexp1:
        if st.button("Explorar modelos", key="btn_explorar_modelos", use_container_width=True):
            st.session_state.vista_resumen = False
            st.rerun()
    with bexp2:
        n_fav = len([m for m in st.session_state.get("favoritos", set()) if m in st.session_state.modelos_data])
        if st.button(f"Ver favoritos ({n_fav})", key="btn_ver_favoritos_resumen", use_container_width=True):
            st.session_state.vista_resumen = False
            st.session_state.vista_favoritos = True
            st.rerun()

def render_vista_favoritos():
    favoritos = st.session_state.get("favoritos", set())
    modelos_data = st.session_state.modelos_data
    favoritos_validos = [m for m in favoritos if m in modelos_data]
    st.markdown(f"""
    <div style="margin-bottom:8px;">
        <p style="font-size:20px;font-weight:700;color:{NAVY};margin:0;">Modelos favoritos</p>
        <p style="font-size:12px;color:{MUTED};margin:4px 0 0;">{len(favoritos_validos)} modelo(s) marcados como favoritos.</p>
    </div>
    <div style="height:1px;background:{BORDER};margin:12px 0 20px;"></div>
    """, unsafe_allow_html=True)
    bcol1, bcol2 = st.columns([1, 5])
    with bcol1:
        if st.button("Volver", key="btn_volver_de_favoritos", use_container_width=True):
            st.session_state.vista_favoritos = False
            st.session_state.vista_resumen = True
            st.rerun()
    if not favoritos_validos:
        st.info("Aun no ha marcado ningun modelo como favorito. Abra un modelo y presione 'Marcar favorito' en la barra lateral, o use el boton en cada tarjeta.")
        return
    filas = []
    for nombre in favoritos_validos:
        datos = modelos_data.get(nombre, {})
        pruebas = datos.get('pruebas')
        score_global, _ = calcular_score_global(pruebas)
        filas.append((nombre, score_global))
    filas.sort(key=lambda x: (-(x[1] if x[1] is not None else -1), x[0]))
    n_cols = 3
    for i in range(0, len(filas), n_cols):
        fila_cols = st.columns(n_cols)
        for j, (nombre, score_global) in enumerate(filas[i:i + n_cols]):
            with fila_cols[j]:
                etiqueta_g, color_g, bg_g = clasificar_score_global(score_global)
                datos = modelos_data.get(nombre, {})
                coefs = datos.get('coeficientes')
                ar_count, ma_count = contar_ar_ma(coefs) if coefs is not None else (0, 0)
                obs = datos.get('observaciones', 0)
                st.markdown(f"""
                <div style="background:{WHITE};border:1px solid {BORDER};border-radius:8px;padding:16px;margin-bottom:10px;">
                    <p style="font-size:13px;font-weight:700;color:{NAVY};margin:0 0 8px;">[F] {nombre}</p>
                    {score_global_badge(score_global)}
                    <p style="font-size:11px;color:{MUTED};margin:10px 0 0;">{obs} observaciones - AR {ar_count} / MA {ma_count}</p>
                </div>
                """, unsafe_allow_html=True)
                bc1, bc2 = st.columns(2)
                with bc1:
                    if st.button("Abrir", key=f"abrir_fav_{nombre}", use_container_width=True):
                        st.session_state.modelo_seleccionado = nombre
                        st.session_state.vista_favoritos = False
                        st.session_state.vista_resumen = False
                        st.rerun()
                with bc2:
                    if st.button("Quitar", key=f"quitar_fav_{nombre}", use_container_width=True):
                        alternar_favorito(nombre)
                        st.rerun()

def render_seccion_coeficientes(datos, key_prefix="diag"):
    st.markdown(section_title("Coeficientes del modelo"), unsafe_allow_html=True)
    coefs = datos.get('coeficientes')
    if coefs is not None and not coefs.empty:
        df_coef = coefs.copy()
        df_coef['Tipo'] = df_coef['Variable'].apply(clasificar_variable)
        if 'P_value' in df_coef.columns:
            def fmt_pval(x):
                if pd.isna(x): return "N/A"
                try:
                    xv = float(x)
                    return f"{xv:.4e}" if xv < 0.001 else f"{xv:.4f}"
                except: return str(x)
            df_coef['P-valor'] = df_coef['P_value'].apply(fmt_pval)
        df_display = df_coef[['Tipo', 'Variable', 'Coeficiente', 'P-valor']]
        st.plotly_chart(fig_barras_coeficientes(df_coef), use_container_width=True, key=f"{key_prefix}_coef_bar")
        def color_pval(v):
            try: return f"color: {GREEN}; font-weight: 700;" if float(v) < 0.05 else f"color: {RED};"
            except: return ""
        styler = df_display.style
        if hasattr(styler, "map"):
            styler = styler.map(color_pval, subset=['P-valor'])
        else:
            styler = styler.applymap(color_pval, subset=['P-valor'])
        st.dataframe(styler, use_container_width=True, hide_index=True, key=f"{key_prefix}_coef_tabla")
        ar_count, ma_count = contar_ar_ma(coefs)
        st.markdown(divider(), unsafe_allow_html=True)
        st.markdown(section_title("Estructura del modelo"), unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        with c1:
            st.markdown(card_metric("Terminos AR", str(ar_count), BLUE), unsafe_allow_html=True)
        with c2:
            st.markdown(card_metric("Terminos MA", str(ma_count), BLUE), unsafe_allow_html=True)
    else:
        st.info("No hay datos de coeficientes.")

# =============================================================================
# GENERADOR DE NOTEBOOKS
# =============================================================================

def render_generador():
    st.markdown(f"<p style='font-size:20px;font-weight:700;color:{NAVY};margin:0 0 8px;'>Generador de Notebooks SARIMAX</p>", unsafe_allow_html=True)
    st.markdown(f"<p style='font-size:12px;color:{MUTED};margin:0 0 20px;'>Configura parametros y descarga los notebooks listos para ejecutar.</p>", unsafe_allow_html=True)

    # Directorio de templates
    TEMPLATES_DIR = Path(__file__).parent
    TEMPLATE_GENERADOR = TEMPLATES_DIR / "Generacion_Variacion__2_.ipynb"
    TEMPLATE_MOTOR = TEMPLATES_DIR / "Motor_Sarimax_Vivi_CO__1_.ipynb"

    templates_ok = TEMPLATE_GENERADOR.exists() and TEMPLATE_MOTOR.exists()
    if not templates_ok:
        st.error("Falta un template. Verifica que estos archivos esten en el mismo directorio:")
        st.code(f"{TEMPLATE_GENERADOR.name}\n{TEMPLATE_MOTOR.name}")
        return

    col1, col2 = st.columns(2)

    with col1:
        st.markdown(section_title("Ubicacion"), unsafe_allow_html=True)
        pais_display = st.selectbox(
            "Pais",
            options=list(PAIS_MAP_GEN.keys()),
            help="Selecciona el pais del modelo"
        )
        pais = PAIS_MAP_GEN[pais_display]

        cartera_display = st.selectbox(
            "Cartera",
            options=list(CARTERA_MAP_GEN.keys()),
            help="Selecciona el portafolio de credito"
        )
        cartera = CARTERA_MAP_GEN[cartera_display]

    with col2:
        st.markdown(section_title("Modo Endogena"), unsafe_allow_html=True)
        modo_endo = st.radio(
            "Calculo de la endogena",
            options=["actual", "media_movil"],
            help="'actual': valor original | 'media_movil': promedio movil"
        )

        if modo_endo == "media_movil":
            ventana_mm = st.slider(
                "Ventana de media movil (meses)",
                min_value=1,
                max_value=12,
                value=3,
                help="Numero de meses para el calculo de promedio movil"
            )
        else:
            ventana_mm = 3

    st.markdown(divider(), unsafe_allow_html=True)
    st.markdown(section_title("Configuracion Generador"), unsafe_allow_html=True)

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
            st.info(f"Inicio: {fecha_inicio} | Fin: {fecha_fin}")
        else:
            col_a, col_b = st.columns(2)
            with col_a:
                fecha_inicio = st.text_input(
                    "Fecha inicio (YYYY-MM-DD)",
                    value="2018-10-01"
                )
            with col_b:
                fecha_fin = st.text_input(
                    "Fecha fin historico (YYYY-MM-DD)",
                    value="2025-03-01"
                )

    with col2:
        editar_nombres = st.checkbox(
            "Editar nombres de archivos",
            value=False,
            help="Si marcas esto, puedes personalizar los nombres"
        )

        if editar_nombres:
            st.warning("Cambiar el formato puede romper el flujo. Procede con cuidado.")

            # Generar nombres sugeridos basados en pais y cartera
            nombres_sugeridos = {
                "hist": f"hist_{pais.lower()}_{cartera.lower()}.xlsx",
                "base": f"base_{pais.lower()}_{cartera.lower()}.xlsx",
                "opt": f"opt_{pais.lower()}_{cartera.lower()}.xlsx",
                "adv": f"adv_{pais.lower()}_{cartera.lower()}.xlsx",
            }

            col_h, col_o = st.columns(2)
            with col_h:
                archivo_hist = st.text_input("Nombre archivo historico", value=nombres_sugeridos["hist"])
                archivo_base = st.text_input("Nombre archivo base", value=nombres_sugeridos["base"])
            with col_o:
                archivo_opt = st.text_input("Nombre archivo optimista", value=nombres_sugeridos["opt"])
                archivo_adv = st.text_input("Nombre archivo adverso", value=nombres_sugeridos["adv"])

            nombres_custom = {
                "hist": archivo_hist,
                "opt": archivo_opt,
                "base": archivo_base,
                "adv": archivo_adv,
            }
        else:
            nombres_custom = None

    st.markdown(divider(), unsafe_allow_html=True)
    st.markdown(section_title("Configuracion Motor SARIMAX"), unsafe_allow_html=True)

    col1, col2 = st.columns(2)

    with col1:
        tipo_modelo = st.selectbox(
            "Tipo de modelo",
            options=["total", "logit"],
            help="'total': escala original | 'logit': transformacion logit"
        )

    with col2:
        max_lags = st.slider(
            "Maximo de lags",
            min_value=1,
            max_value=12,
            value=8,
            help="Numero maximo de rezagos a evaluar en los modelos"
        )

    st.info("""
    Los siguientes parametros NO se editan aqui (requieren cambios en Colab):
    - Signos exogenas (positivo/negativo)
    - VIF maximo
    - Umbral de sensibilidad
    - Factor FWL (1.0 - 1.2)
    """)

    st.markdown(divider(), unsafe_allow_html=True)

    if st.button("Generar Notebooks", use_container_width=True, type="primary"):
        try:
            # Importar funciones del generador si existen
            try:
                from notebook_generator import (
                    generar_notebook_generador,
                    generar_notebook_motor,
                    generar_nombres_archivos,
                )
                gen_importado = True
            except ImportError:
                gen_importado = False
                st.error("No se encontro el modulo 'notebook_generator'. Verifica que este en el mismo directorio.")
                return

            # Generar nombres de archivos
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

            # Nombres de salida: usamos cartera_display para que refleje lo que el usuario selecciono
            cartera_slug = cartera_display.lower().replace(" ", "_")
            nb_gen_name = f"Generacion_Variacion_{pais}_{cartera_slug}.ipynb"
            nb_motor_name = f"Motor_Sarimax_{cartera_slug}_{pais}.ipynb"

            # Convertir notebooks a JSON (string) para descarga individual
            nb_gen_json = json.dumps(nb_gen, ensure_ascii=False, indent=1)
            nb_motor_json = json.dumps(nb_motor, ensure_ascii=False, indent=1)

            # Mostrar resumen
            st.success("Notebooks generados correctamente")

            st.markdown(section_title("Resumen de configuracion"), unsafe_allow_html=True)
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.markdown(card_kpi("Pais", pais), unsafe_allow_html=True)
            with col2:
                st.markdown(card_kpi("Cartera", cartera_display), unsafe_allow_html=True)
            with col3:
                st.markdown(card_kpi("Modelo", tipo_modelo), unsafe_allow_html=True)
            with col4:
                st.markdown(card_kpi("Max Lags", str(max_lags)), unsafe_allow_html=True)

            st.markdown(section_title("Descargar o abrir en Google Colab"), unsafe_allow_html=True)

            # --- NOTEBOOK GENERADOR ---
            st.markdown("<p style='font-size:14px;font-weight:600;color:{NAVY};margin:8px 0 4px;'>Generador de Variacion</p>".format(NAVY=NAVY), unsafe_allow_html=True)

            col_gen1, col_gen2 = st.columns(2)

            with col_gen1:
                st.download_button(
                    label=f"Descargar {nb_gen_name}",
                    data=nb_gen_json,
                    file_name=nb_gen_name,
                    mime="application/json",
                    use_container_width=True,
                )

            with col_gen2:
                nb_gen_b64 = base64.b64encode(nb_gen_json.encode("utf-8")).decode("utf-8")
                colab_url_gen = f"https://colab.research.google.com/notebook#data={nb_gen_b64}"
                st.link_button(
                    label="Abrir en Google Colab",
                    url=colab_url_gen,
                    use_container_width=True,
                )

            # --- NOTEBOOK MOTOR ---
            st.markdown("<p style='font-size:14px;font-weight:600;color:{NAVY};margin:8px 0 4px;'>Motor SARIMAX</p>".format(NAVY=NAVY), unsafe_allow_html=True)

            col_mot1, col_mot2 = st.columns(2)

            with col_mot1:
                st.download_button(
                    label=f"Descargar {nb_motor_name}",
                    data=nb_motor_json,
                    file_name=nb_motor_name,
                    mime="application/json",
                    use_container_width=True,
                )

            with col_mot2:
                nb_motor_b64 = base64.b64encode(nb_motor_json.encode("utf-8")).decode("utf-8")
                colab_url_motor = f"https://colab.research.google.com/notebook#data={nb_motor_b64}"
                st.link_button(
                    label="Abrir en Google Colab",
                    url=colab_url_motor,
                    use_container_width=True,
                )

        except Exception as e:
            st.error(f"Error al generar notebooks:\n{str(e)}")
            st.exception(e)

# =============================================================================
# SESSION STATE
# =============================================================================
_prefs_guardadas = cargar_prefs_sidebar()
for key, default in [
    ("uploaded_file", None), ("modelos_data", {}), ("meta_contexto", None),
    ("modelo_seleccionado", None),
    ("criterio_ordenamiento", _prefs_guardadas.get("criterio_ordenamiento", "Pruebas aprobadas ↓")),
    ("exog_sel", {}), ("pred_filtro", "Todas"),
    ("nav_sticky", _prefs_guardadas.get("nav_sticky", True)),
    ("pending_modelo", None),
    ("filtro_ljung", _prefs_guardadas.get("filtro_ljung", "Todos")),
    ("filtro_jarque", _prefs_guardadas.get("filtro_jarque", "Todos")),
    ("filtro_hetero", _prefs_guardadas.get("filtro_hetero", "Todos")),
    ("vista_resumen", True), ("comparar_sel", []),
    ("favoritos", set()),
    ("filtro_favoritos", _prefs_guardadas.get("filtro_favoritos", False)),
    ("vista_favoritos", False),
    ("sec_contexto", _prefs_guardadas.get("sec_contexto", True)),
    ("sec_orden", _prefs_guardadas.get("sec_orden", True)),
    ("sec_exogenas", _prefs_guardadas.get("sec_exogenas", True)),
]:
    if key not in st.session_state:
        st.session_state[key] = default

# =============================================================================
# APP PRINCIPAL
# =============================================================================
st.set_page_config(page_title="SARIMAX IFRS 9", layout="wide")
inject_css()

# =========================================================================
# ATAJOS DE TECLADO (flechas para navegar, Escape para limpiar filtros)
# =========================================================================
with st.container(key="kb_hidden"):
    kb_c1, kb_c2, kb_c3 = st.columns(3)
    with kb_c1:
        kb_prev_clicked = st.button("KB_PREV", key="btn_kb_prev")
    with kb_c2:
        kb_next_clicked = st.button("KB_NEXT", key="btn_kb_next")
    with kb_c3:
        kb_reset_clicked = st.button("KB_RESET", key="btn_kb_reset")

if kb_reset_clicked:
    st.session_state.filtro_ljung_sel = "Todos"
    st.session_state.filtro_jarque_sel = "Todos"
    st.session_state.filtro_hetero_sel = "Todos"
    st.session_state.filtro_ljung = "Todos"
    st.session_state.filtro_jarque = "Todos"
    st.session_state.filtro_hetero = "Todos"
    st.rerun()

if (kb_prev_clicked or kb_next_clicked) and st.session_state.modelos_data and st.session_state.modelo_seleccionado:
    modelos_list_kb, _, _, _ = construir_opciones_modelos()
    if st.session_state.modelo_seleccionado in modelos_list_kb:
        idx_kb = modelos_list_kb.index(st.session_state.modelo_seleccionado)
        if kb_prev_clicked and idx_kb > 0:
            st.session_state.pending_modelo = modelos_list_kb[idx_kb - 1]
            st.rerun()
        elif kb_next_clicked and idx_kb < len(modelos_list_kb) - 1:
            st.session_state.pending_modelo = modelos_list_kb[idx_kb + 1]
            st.rerun()

_KB_SHORTCUT_HTML = """
<script>
(function() {
    if (window.parent.__kbListenerAdded) { return; }
    window.parent.__kbListenerAdded = true;
    function clickHiddenButton(cssKey, fallbackText) {
        const doc = window.parent.document;
        let btn = doc.querySelector('.st-key-' + cssKey + ' button');
        if (!btn) {
            const all = doc.querySelectorAll('button');
            for (const b of all) {
                if (b.innerText && b.innerText.trim() === fallbackText) { btn = b; break; }
            }
        }
        if (btn) { btn.click(); }
    }
    window.parent.document.addEventListener('keydown', function(e) {
        const tag = (e.target && e.target.tagName) || '';
        if (['INPUT', 'TEXTAREA', 'SELECT'].includes(tag)) return;
        if (e.key === 'ArrowLeft') {
            clickHiddenButton('btn_kb_prev', 'KB_PREV');
        } else if (e.key === 'ArrowRight') {
            clickHiddenButton('btn_kb_next', 'KB_NEXT');
        } else if (e.key === 'Escape') {
            clickHiddenButton('btn_kb_reset', 'KB_RESET');
        }
    });
})();
</script>
"""
if hasattr(st, "iframe"):
    st.iframe(_KB_SHORTCUT_HTML, height=1, width=1)
else:
    components.html(_KB_SHORTCUT_HTML, height=0, width=0)

# =========================================================================
# HEADER Y PESTANAS
# =========================================================================
st.markdown(f"""
<div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:4px;">
    <div>
        <p style="font-size:22px;font-weight:700;color:{NAVY};margin:0;">SARIMAX IFRS 9</p>
        <p style="font-size:12px;color:{MUTED};margin:4px 0 0;">Generador y Dashboard de Modelos</p>
    </div>
</div>
<div style="height:1px;background:{BORDER};margin:12px 0 16px;"></div>
""", unsafe_allow_html=True)

tab_gen, tab_dash = st.tabs(["Generador", "Dashboard"])

# =========================================================================
# TAB: GENERADOR
# =========================================================================
with tab_gen:
    render_generador()

# =========================================================================
# TAB: DASHBOARD
# =========================================================================
with tab_dash:
    col_left, col_right = st.columns([1, 4])

    # --- SIDEBAR DEL DASHBOARD ---
    with col_left:
        st.markdown(f"<p style='font-size:12px;font-weight:700;color:{NAVY};margin:0 0 10px;letter-spacing:0.5px;'>CARGAR MODELO</p>", unsafe_allow_html=True)
        uploaded = st.file_uploader("Archivo Excel (.xlsx)", type=["xlsx"], label_visibility="collapsed")
        if uploaded is not None:
            st.session_state.uploaded_file = uploaded
            if not st.session_state.modelos_data or uploaded.name != getattr(st.session_state, 'last_file_name', None):
                with st.spinner("Parseando modelos..."):
                    st.session_state.modelos_data = parsear_excel(uploaded)
                    st.session_state.meta_contexto = leer_meta_embebida(uploaded)
                    st.session_state.last_file_name = uploaded.name
                    st.session_state.vista_resumen = True
                    st.session_state.comparar_sel = []
                st.success(f"Archivo cargado: {uploaded.name}")
                if st.session_state.modelo_seleccionado is None or st.session_state.modelo_seleccionado not in st.session_state.modelos_data:
                    st.session_state.modelo_seleccionado = list(st.session_state.modelos_data.keys())[0]
        if st.session_state.uploaded_file is not None:
            if st.button("Eliminar archivo", key="btn_eliminar", use_container_width=True):
                st.session_state.uploaded_file = None
                st.session_state.modelos_data = {}
                st.session_state.meta_contexto = None
                st.session_state.modelo_seleccionado = None
                st.session_state.last_file_name = None
                st.session_state.exog_sel = {}
                st.session_state.vista_resumen = True
                st.session_state.comparar_sel = []
                st.rerun()
        if st.session_state.modelos_data:
            st.markdown(divider(), unsafe_allow_html=True)
            meta = st.session_state.meta_contexto
            if encabezado_colapsable("Contexto de la corrida", "sec_contexto"):
                if meta:
                    meta_kpis = extraer_kpis_meta(meta)
                    c1, c2 = st.columns(2)
                    with c1:
                        pais_nombre = meta_kpis.get('pais', '-')
                        codigo_iso = BANDERAS_PAISES.get(pais_nombre, pais_nombre).upper()
                        bandera_html = obtener_bandera_pais(pais_nombre)
                        valor_pais = f"{bandera_html}{codigo_iso}"
                        st.markdown(card_kpi("Pais", valor_pais), unsafe_allow_html=True)
                    with c2:
                        st.markdown(card_kpi("Ventana media movil", meta_kpis.get('ventana_mm', '-')), unsafe_allow_html=True)
                    c1, c2 = st.columns(2)
                    with c1:
                        st.markdown(card_kpi("Cartera", meta_kpis.get('cartera', '-'), accent=BLUE), unsafe_allow_html=True)
                    with c2:
                        fwl_range = f"{meta_kpis.get('fwl_min', '?')} - {meta_kpis.get('fwl_max', '?')}"
                        st.markdown(card_kpi("Rango FWL", fwl_range, accent=GREEN), unsafe_allow_html=True)
                    c1, c2 = st.columns(2)
                    with c1:
                        st.markdown(card_kpi("Modo endogena", meta_kpis.get('modo_endogena', '-')), unsafe_allow_html=True)
                    with c2:
                        tipo_endog = meta_kpis.get('tipo_endogena', '-')
                        # Mostrar en mayusculas y con color si es logit
                        if tipo_endog == 'logit':
                            tipo_display = f"<span style='color:#1E5AA8;font-weight:700;'>{tipo_endog.upper()}</span>"
                        elif tipo_endog == 'total':
                            tipo_display = f"<span style='color:#1A6B3E;font-weight:700;'>{tipo_endog.upper()}</span>"
                        else:
                            tipo_display = tipo_endog
                        st.markdown(card_kpi("Tipo endogena", tipo_display), unsafe_allow_html=True)
                else:
                    st.caption("Sin metadata embebida.")
            st.markdown(divider(), unsafe_allow_html=True)
            if encabezado_colapsable("Ordenar y Filtrar", "sec_orden"):
                criterio = st.radio("Ordenar por:", ["Nombre (A-Z)", "Pruebas aprobadas ↓", "Pruebas aprobadas ↑",
                                                      "Score global ↓", "Score global ↑"],
                                    index=["Nombre (A-Z)", "Pruebas aprobadas ↓", "Pruebas aprobadas ↑",
                                           "Score global ↓", "Score global ↑"].index(st.session_state.criterio_ordenamiento),
                                    key="criterio_orden")
                st.session_state.criterio_ordenamiento = criterio
                st.markdown(f"<p style='font-size:10px;color:{MUTED};margin:8px 0 4px;'>Ljung-Box</p>", unsafe_allow_html=True)
                filtro_ljung = st.selectbox("", ["Todos", "A o B (Cumple)", "A, B o C", "Solo A"],
                                             index=["Todos", "A o B (Cumple)", "A, B o C", "Solo A"].index(st.session_state.filtro_ljung),
                                             key="filtro_ljung_sel", label_visibility="collapsed")
                st.session_state.filtro_ljung = filtro_ljung
                st.markdown(f"<p style='font-size:10px;color:{MUTED};margin:8px 0 4px;'>Jarque-Bera</p>", unsafe_allow_html=True)
                filtro_jarque = st.selectbox("", ["Todos", "A o B (Cumple)", "A, B o C", "Solo A"],
                                              index=["Todos", "A o B (Cumple)", "A, B o C", "Solo A"].index(st.session_state.filtro_jarque),
                                              key="filtro_jarque_sel", label_visibility="collapsed")
                st.session_state.filtro_jarque = filtro_jarque
                st.markdown(f"<p style='font-size:10px;color:{MUTED};margin:8px 0 4px;'>Heterocedasticidad</p>", unsafe_allow_html=True)
                filtro_hetero = st.selectbox("", ["Todos", "A o B (Cumple)", "A, B o C", "Solo A"],
                                              index=["Todos", "A o B (Cumple)", "A, B o C", "Solo A"].index(st.session_state.filtro_hetero),
                                              key="filtro_hetero_sel", label_visibility="collapsed")
                st.session_state.filtro_hetero = filtro_hetero
                st.markdown("<div style='height:6px;'></div>", unsafe_allow_html=True)
                filtro_favoritos = st.checkbox("Solo favoritos", key="filtro_favoritos_sel",
                                                value=st.session_state.get("filtro_favoritos", False))
                st.session_state.filtro_favoritos = filtro_favoritos
            st.markdown(divider(), unsafe_allow_html=True)
            modelos_list, pruebas_dict, scores_dict, global_dict = construir_opciones_modelos()
            if st.session_state.pending_modelo is not None and st.session_state.pending_modelo in modelos_list:
                st.session_state.modelo_seleccionado = st.session_state.pending_modelo
                st.session_state["sel_modelo"] = label_modelo(st.session_state.pending_modelo, pruebas_dict, scores_dict, global_dict)
                st.session_state.pending_modelo = None
            opciones = [label_modelo(m, pruebas_dict, scores_dict, global_dict) for m in modelos_list]
            if not modelos_list:
                st.warning("Ningun modelo cumple con los filtros seleccionados.")
                st.session_state.modelo_seleccionado = None
            else:
                idx = modelos_list.index(st.session_state.modelo_seleccionado) if st.session_state.modelo_seleccionado in modelos_list else 0
                seleccion = st.selectbox("Modelo", opciones, index=idx, key="sel_modelo")
                nombre_parseado = seleccion.split("  (")[0]
                if nombre_parseado.startswith("[F] "):
                    nombre_parseado = nombre_parseado[4:]
                st.session_state.modelo_seleccionado = nombre_parseado
            st.markdown(divider(), unsafe_allow_html=True)
            st.toggle("Fijar flechas de navegacion", key="nav_sticky",
                      help="Mantiene los botones Anterior/Siguiente siempre visibles, flotando sobre la pagina al hacer scroll.")
            guardar_prefs_sidebar()
            if st.session_state.modelo_seleccionado:
                datos = st.session_state.modelos_data.get(st.session_state.modelo_seleccionado, {})
                st.markdown(f"<p style='font-size:11px;font-weight:600;color:{NAVY};margin:12px 0 4px;'>MODELO ACTUAL</p>", unsafe_allow_html=True)
                st.markdown(f"<p style='font-size:14px;font-weight:700;color:{NAVY};margin:0;'>{st.session_state.modelo_seleccionado}</p>", unsafe_allow_html=True)
                st.markdown(f"<p style='font-size:11px;color:{MUTED};margin:4px 0 0;'>{datos.get('observaciones', 0)} observaciones</p>", unsafe_allow_html=True)
                st.markdown("<div style='height:6px;'></div>", unsafe_allow_html=True)
                boton_favorito(st.session_state.modelo_seleccionado, key_suffix="sidebar")
                exogenas = datos.get('exogenas_nombres', [])
                if exogenas:
                    st.markdown("<div style='height:8px;'></div>", unsafe_allow_html=True)
                    if encabezado_colapsable("Exogenas", "sec_exogenas"):
                        coefs = datos.get('coeficientes')
                        sigs = obtener_significancia_exogenas(coefs, exogenas)
                        sig_count = sum(1 for _, _, s in sigs if s == "Significativa")
                        st.markdown(f"<p style='font-size:10px;color:{MUTED};margin:0 0 6px;'>{sig_count} de {len(exogenas)} significativas</p>", unsafe_allow_html=True)
                        for ex, pval, status in sigs:
                            color = GREEN if status == "Significativa" else (RED if status == "No significativa" else "#B8860B")
                            label = "SIG" if status == "Significativa" else ("NO SIG" if status == "No significativa" else "MARG")
                            p_txt = f"p={pval:.3f}" if pval is not None else "p=N/A"
                            st.markdown(f'<div style="display:flex;justify-content:space-between;align-items:center;padding:2px 0;font-size:11px;"><span style="color:{TEXT}">{ex}</span><span style="color:{color};font-weight:600;">{label} ({p_txt})</span></div>', unsafe_allow_html=True)

    # --- PANEL PRINCIPAL DEL DASHBOARD ---
    with col_right:
        if not st.session_state.modelos_data:
            st.markdown(f"""
            <div style="background:{WHITE};border:1px solid {BORDER};border-radius:10px;padding:60px 32px;text-align:center;margin-top:40px;">
                <p style="font-size:18px;font-weight:700;color:{NAVY};margin:0 0 8px;">Dashboard SARIMAX</p>
                <p style="font-size:13px;color:{MUTED};margin:0;">Suba un archivo Excel para comenzar el analisis.</p>
            </div>
            """, unsafe_allow_html=True)
        elif st.session_state.vista_favoritos:
            render_vista_favoritos()
        elif st.session_state.vista_resumen:
            render_resumen_ejecutivo()
        elif st.session_state.modelo_seleccionado is None:
            st.markdown(f"""
            <div style="background:{WHITE};border:1px solid {BORDER};border-radius:10px;padding:60px 32px;text-align:center;margin-top:40px;">
                <p style="font-size:18px;font-weight:700;color:{NAVY};margin:0 0 8px;">Sin modelos disponibles</p>
                <p style="font-size:13px;color:{MUTED};margin:0;">Ajuste los filtros de diagnostico para ver modelos.</p>
            </div>
            """, unsafe_allow_html=True)
        else:
            datos = st.session_state.modelos_data.get(st.session_state.modelo_seleccionado, {})
            meta_kpis = extraer_kpis_meta(st.session_state.meta_contexto)
            pais = meta_kpis.get('pais', '-')
            cartera = meta_kpis.get('cartera', '-')
            pais_codigo_hdr = BANDERAS_PAISES.get(pais, pais).upper()
            score_global_hdr, _ = calcular_score_global(datos.get('pruebas'))
            st.markdown(f"""
            <div style="display:flex;align-items:flex-end;gap:16px;margin-bottom:4px;">
                <div style="flex:1;">
                    <p style="font-size:11px;color:{MUTED};margin:0;text-transform:uppercase;letter-spacing:0.5px;font-weight:600;">Modelo seleccionado</p>
                    <p style="font-size:20px;font-weight:700;color:{NAVY};margin:4px 0 0;">{st.session_state.modelo_seleccionado}
                        <span style="margin-left:10px;">{score_global_badge(score_global_hdr)}</span>
                    </p>
                </div>
                <div style="text-align:right;">
                    <p style="font-size:11px;color:{MUTED};margin:0;">{obtener_bandera_pais(pais)}{pais_codigo_hdr} - {cartera} | {meta_kpis.get('tipo_endogena', '').upper()}</p>
                    <p style="font-size:11px;color:{LTGRAY};margin:2px 0 0;">{len(st.session_state.modelos_data)} modelos cargados</p>
                </div>
            </div>
            <div style="height:1px;background:{BORDER};margin:12px 0 16px;"></div>
            """, unsafe_allow_html=True)
            hcol1, hcol2, hcol3 = st.columns([1.4, 1.1, 1])
            with hcol1:
                if st.button("Ver resumen de la corrida", key="btn_ver_resumen", use_container_width=True):
                    st.session_state.vista_resumen = True
                    st.rerun()
            with hcol2:
                n_fav = len([m for m in st.session_state.get("favoritos", set()) if m in st.session_state.modelos_data])
                if st.button(f"Ver favoritos ({n_fav})", key="btn_ver_favoritos_detalle", use_container_width=True):
                    st.session_state.vista_favoritos = True
                    st.rerun()
            with hcol3:
                boton_favorito(st.session_state.modelo_seleccionado, key_suffix="detalle")
            modelos_list, pruebas_dict_nav, scores_dict_nav, global_dict_nav = construir_opciones_modelos()
            current_idx = modelos_list.index(st.session_state.modelo_seleccionado) if st.session_state.modelo_seleccionado in modelos_list else 0
            tab_resumen, tab1, tab2, tab3, tab4 = st.tabs(["Resumen Modelo", "Visualizacion", "Predicciones", "Diagnosticos", "Comparar"])
            # =====================================================================
            # TAB 0: RESUMEN MODELO
            # =====================================================================
            with tab_resumen:
                st.markdown(section_title("Factor FWL a 12 meses"), unsafe_allow_html=True)
                df_fwl_resumen = datos.get('fwl_12m')
                if df_fwl_resumen is not None and not df_fwl_resumen.empty:
                    st.plotly_chart(fig_fwl_12m(df_fwl_resumen), use_container_width=True, key="resumen_fwl12m")
                else:
                    st.info("No hay datos de FWL a 12 meses.")
                st.markdown(divider(), unsafe_allow_html=True)
                render_diagnosticos_corporativo(datos.get('pruebas'), mostrar_detalle_tecnico=False)
                st.markdown(divider(), unsafe_allow_html=True)
                render_seccion_coeficientes(datos, key_prefix="resumen")
            # =====================================================================
            # TAB 1: VISUALIZACION
            # =====================================================================
            with tab1:
                st.markdown(section_title("Exogenas activas"), unsafe_allow_html=True)
                exogenas = datos.get('exogenas_nombres', [])
                modelo_key = st.session_state.modelo_seleccionado
                if modelo_key not in st.session_state.exog_sel:
                    st.session_state.exog_sel[modelo_key] = []
                if exogenas:
                    n_cols = min(len(exogenas), 6)
                    chip_cols = st.columns(n_cols)
                    for i, ex in enumerate(exogenas):
                        with chip_cols[i % n_cols]:
                            activo = ex in st.session_state.exog_sel[modelo_key]
                            label = f"[x] {ex}" if activo else f"[ ] {ex}"
                            if st.button(label, key=f"chip_{modelo_key}_{ex}", use_container_width=True):
                                if activo:
                                    st.session_state.exog_sel[modelo_key].remove(ex)
                                else:
                                    st.session_state.exog_sel[modelo_key].append(ex)
                                st.rerun()
                else:
                    st.caption("Sin exogenas en este modelo.")
                df_end = datos.get('fecha_endogena')
                endogena_cols = datos.get('endogenas_cols', [])
                if df_end is not None and not df_end.empty and endogena_cols:
                    fig = fig_predicciones(df_end, endogena_cols, datos.get('exogenas'), st.session_state.exog_sel.get(modelo_key, []), st.session_state.modelo_seleccionado)
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    st.info("No hay datos de predicciones.")
                st.markdown(divider(), unsafe_allow_html=True)
                st.markdown(section_title("Factor FWL por ano y escenario"), unsafe_allow_html=True)
                df_fwl_anual = datos.get('fwl_anual')
                if df_fwl_anual is not None and not df_fwl_anual.empty:
                    try:
                        df_pivot = df_fwl_anual.pivot(index='Ano', columns='Escenario', values='Factor FWL').reset_index()
                        rename_map = {}
                        for c in df_pivot.columns:
                            c_str = str(c).lower()
                            if 'base' in c_str: rename_map[c] = 'Base'
                            elif 'adverso' in c_str or 'advers' in c_str: rename_map[c] = 'Adverso'
                            elif 'optimista' in c_str: rename_map[c] = 'Optimista'
                        df_pivot = df_pivot.rename(columns=rename_map)
                        st.dataframe(df_pivot, use_container_width=True, hide_index=True)
                    except:
                        st.dataframe(df_fwl_anual, use_container_width=True, hide_index=True)
                else:
                    st.info("No hay datos de Factor FWL por Ano.")
                st.markdown(divider(), unsafe_allow_html=True)
                st.markdown(section_title("Factor FWL a 12 meses"), unsafe_allow_html=True)
                df_fwl = datos.get('fwl_12m')
                if df_fwl is not None and not df_fwl.empty:
                    st.plotly_chart(fig_fwl_12m(df_fwl), use_container_width=True, key="viz_fwl12m")
                else:
                    st.info("No hay datos de FWL a 12 meses.")
                st.markdown(divider(), unsafe_allow_html=True)
                st.markdown(section_title("Factor FWL ponderado"), unsafe_allow_html=True)
                c1, c2, c3 = st.columns(3)
                with c1: peso_base = st.number_input("Peso Base", 0.0, 1.0, 0.33, 0.01, key="pw_base")
                with c2: peso_adverso = st.number_input("Peso Adverso", 0.0, 1.0, 0.33, 0.01, key="pw_adv")
                with c3: peso_optimista = st.number_input("Peso Optimista", 0.0, 1.0, 0.34, 0.01, key="pw_opt")
                suma = peso_base + peso_adverso + peso_optimista
                if abs(suma - 1.0) < 0.001:
                    st.markdown(pill("VALIDO", "#E8F5E9", GREEN), unsafe_allow_html=True)
                elif suma < 1.0:
                    st.markdown(pill(f"{1.0-suma:.2f} DISPONIBLE", "#FFF8E1", "#B8860B"), unsafe_allow_html=True)
                else:
                    st.markdown(pill(f"EXCEDE {suma-1.0:.2f}", "#FFEBEE", RED), unsafe_allow_html=True)
                if df_fwl is not None and not df_fwl.empty and suma <= 1.0:
                    pesos = {'base': peso_base, 'adverso': peso_adverso, 'optimista': peso_optimista}
                    df_pond = calcular_fwl_ponderado(df_fwl, pesos)
                    if df_pond is not None:
                        res = resumen_fwl(df_pond)
                        if res:
                            c1, c2, c3, c4 = st.columns(4)
                            with c1: st.markdown(card_metric("Promedio", f"{res.get('promedio', 0):.4f}", BLUE), unsafe_allow_html=True)
                            with c2: st.markdown(card_metric("Maximo", f"{res.get('maximo', 0):.4f}", GREEN), unsafe_allow_html=True)
                            with c3: st.markdown(card_metric("Minimo", f"{res.get('minimo', 0):.4f}", RED), unsafe_allow_html=True)
                            with c4: st.markdown(card_metric("Volatilidad (s)", f"{res.get('volatilidad', 0):.4f}", GRAY), unsafe_allow_html=True)
                        st.plotly_chart(fig_fwl_ponderado(df_pond), use_container_width=True)
                    else:
                        st.info("No se pudo calcular el FWL ponderado.")
                elif suma > 1.0:
                    st.warning("Ajuste los pesos para que la suma no exceda 1.0")
            # =====================================================================
            # TAB 2: PREDICCIONES
            # =====================================================================
            with tab2:
                st.markdown(section_title("Datos de prediccion"), unsafe_allow_html=True)
                filtros = st.columns(4)
                with filtros[0]:
                    if st.button("Ver Base", use_container_width=True): st.session_state.pred_filtro = "Base"
                with filtros[1]:
                    if st.button("Ver Adverso", use_container_width=True): st.session_state.pred_filtro = "Adverso"
                with filtros[2]:
                    if st.button("Ver Optimista", use_container_width=True): st.session_state.pred_filtro = "Optimista"
                with filtros[3]:
                    if st.button("Ver todas", use_container_width=True): st.session_state.pred_filtro = "Todas"
                st.markdown(f"<p style='font-size:11px;color:{MUTED};margin:8px 0;'>Filtro activo: <b>{st.session_state.pred_filtro}</b></p>", unsafe_allow_html=True)
                df_end = datos.get('fecha_endogena')
                endogena_cols = datos.get('endogenas_cols', [])
                if df_end is not None and not df_end.empty and endogena_cols:
                    fecha_col = 'fecha' if 'fecha' in df_end.columns else df_end.columns[0]
                    base_col = adv_col = opt_col = None
                    for col in endogena_cols:
                        col_str = str(col).upper()
                        if col_str == 'BASE': base_col = col
                        elif col_str in ['ADVERSO', 'ADVERSA']: adv_col = col
                        elif col_str == 'OPTIMISTA': opt_col = col
                    df_pred = pd.DataFrame()
                    df_pred['Fecha'] = pd.to_datetime(df_end[fecha_col]).dt.strftime('%Y-%m-%d')
                    cols_export = ['Fecha']
                    if base_col and base_col in df_end.columns and st.session_state.pred_filtro in ["Base", "Todas"]:
                        df_pred['Base'] = df_end[base_col].astype(float).round(4)
                        cols_export.append('Base')
                    if adv_col and adv_col in df_end.columns and st.session_state.pred_filtro in ["Adverso", "Todas"]:
                        df_pred['Adverso'] = df_end[adv_col].astype(float).round(4)
                        cols_export.append('Adverso')
                    if opt_col and opt_col in df_end.columns and st.session_state.pred_filtro in ["Optimista", "Todas"]:
                        df_pred['Optimista'] = df_end[opt_col].astype(float).round(4)
                        cols_export.append('Optimista')
                    st.dataframe(df_pred[cols_export], use_container_width=True, hide_index=True, height=400)
                    csv = df_pred[cols_export].to_csv(index=False).encode('utf-8')
                    st.download_button("Descargar CSV", csv, f"predicciones_{st.session_state.modelo_seleccionado}.csv", "text/csv")
                else:
                    st.info("No hay datos de predicciones.")
            # =====================================================================
            # TAB 3: DIAGNOSTICOS
            # =====================================================================
            with tab3:
                st.markdown(render_leyenda_scores(), unsafe_allow_html=True)
                pruebas = datos.get('pruebas')
                render_metricas_diagnostico(pruebas)
                st.markdown(divider(), unsafe_allow_html=True)
                render_diagnosticos_corporativo(pruebas)
                st.markdown(divider(), unsafe_allow_html=True)
                st.markdown(section_title("Distribucion de residuos"), unsafe_allow_html=True)
                residuos = datos.get('residuos_ind')
                if residuos is not None and not residuos.empty:
                    res_col = None
                    for c in residuos.columns:
                        if 'residuo' in str(c).lower():
                            res_col = c
                            break
                    if res_col:
                        vals = residuos[res_col].dropna().astype(float)
                        media, std = vals.mean(), vals.std()
                        st.plotly_chart(fig_histograma_residuos(vals, media, std), use_container_width=True)
                        st.markdown(section_title("Estadisticas descriptivas"), unsafe_allow_html=True)
                        c1, c2, c3, c4, c5 = st.columns(5)
                        with c1: st.markdown(card_metric("Media", f"{media:.4f}"), unsafe_allow_html=True)
                        with c2: st.markdown(card_metric("Desv. Std.", f"{std:.4f}"), unsafe_allow_html=True)
                        with c3: st.markdown(card_metric("Asimetria", f"{stats.skew(vals):.4f}"), unsafe_allow_html=True)
                        with c4: st.markdown(card_metric("Curtosis", f"{stats.kurtosis(vals):.4f}"), unsafe_allow_html=True)
                        with c5: st.markdown(card_metric("Observaciones", f"{len(vals)}"), unsafe_allow_html=True)
                else:
                    st.info("No hay datos de residuos.")
                st.markdown(divider(), unsafe_allow_html=True)
                render_seccion_coeficientes(datos, key_prefix="diag")
            # =====================================================================
            # TAB 4: COMPARAR
            # =====================================================================
            with tab4:
                st.markdown(section_title("Comparador de modelos"), unsafe_allow_html=True)
                st.markdown(f"<p style='font-size:11px;color:{MUTED};margin:0 0 10px;'>Seleccione hasta 3 modelos para comparar lado a lado.</p>", unsafe_allow_html=True)
                todos_modelos = list(st.session_state.modelos_data.keys())
                default_sel = [m for m in st.session_state.get("comparar_sel", []) if m in todos_modelos][:3]
                seleccion_comp = st.multiselect("Modelos a comparar", todos_modelos, default=default_sel, key="comparar_multiselect")
                if len(seleccion_comp) > 3:
                    st.warning("Se seleccionaron mas de 3 modelos. Solo se compararan los primeros 3.")
                    seleccion_comp = seleccion_comp[:3]
                st.session_state.comparar_sel = seleccion_comp
                if seleccion_comp:
                    cols_comp = st.columns(len(seleccion_comp))
                    for i_comp, nombre_comp in enumerate(seleccion_comp):
                        with cols_comp[i_comp]:
                            render_columna_comparacion(nombre_comp)
                else:
                    st.info("Seleccione al menos un modelo para iniciar la comparacion.")
            # --- Bottom nav bar ---
            nav_sticky = st.session_state.get("nav_sticky", True)
            if nav_sticky:
                st.markdown(f"""
                <style>
                div[data-testid="stVerticalBlockBorderWrapper"]:has(> div > div.st-key-nav_flechas),
                .st-key-nav_flechas {{
                    position: fixed !important;
                    bottom: 22px;
                    left: 50%;
                    transform: translateX(-46%);
                    z-index: 9999;
                    background: {WHITE};
                    border: 1px solid {BORDER};
                    border-radius: 14px;
                    box-shadow: 0 8px 28px rgba(11,37,69,0.16);
                    padding: 6px 10px !important;
                    width: auto !important;
                    max-width: 560px;
                }}
                .block-container {{ padding-bottom: 120px !important; }}
                </style>
                """, unsafe_allow_html=True)
            else:
                st.markdown("<div style='height:40px;'></div>", unsafe_allow_html=True)
            with st.container(key="nav_flechas"):
                nav_cols = st.columns([1, 2, 1])
                with nav_cols[0]:
                    if st.button("Anterior", disabled=current_idx == 0, key="btn_prev_real", use_container_width=True):
                        st.session_state.pending_modelo = modelos_list[current_idx - 1]
                        st.rerun()
                with nav_cols[1]:
                    st.markdown(
                        f"""
                        <div style="text-align:center;padding:6px 4px;">
                            <p style="font-size:10px;color:{MUTED};margin:0;text-transform:uppercase;letter-spacing:0.5px;font-weight:600;">Modelo {current_idx + 1} de {len(modelos_list)}</p>
                            <p style="font-size:13px;color:{NAVY};font-weight:700;margin:2px 0 0;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">{st.session_state.modelo_seleccionado}</p>
                        </div>
                        """,
                        unsafe_allow_html=True
                    )
                with nav_cols[2]:
                    if st.button("Siguiente", disabled=current_idx == len(modelos_list) - 1, key="btn_next_real", use_container_width=True):
                        st.session_state.pending_modelo = modelos_list[current_idx + 1]
                        st.rerun()
