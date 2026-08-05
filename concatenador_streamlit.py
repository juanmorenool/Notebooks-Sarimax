"""
Modulo de concatenador de archivos Impacto para IFRS 9.
Extrae la logica del notebook y la expone como funciones reutilizables.
"""

import pandas as pd
from pathlib import Path
from functools import reduce
import re
from io import BytesIO

# =============================================================================
# CONFIGURACION
# =============================================================================

ALIAS_PRODUCTO = {
    "cons": "consumo",
    "consumo": "consumo",
    "corporativo": "corporativo",
    "corp": "corporativo",
    "pyme": "pymes",
    "pymes": "pymes",
    "tarjeta": "tarjeta",
    "vehiculo": "vehiculo",
    "vivi": "vivienda",
    "vivienda": "vivienda",
}

KEYS_MERGE = ["PAIS", "date", "Scenario_name"]

# =============================================================================
# FUNCIONES
# =============================================================================

def detectar_producto_desde_nombre(nombre_archivo: str) -> str:
    """
    Detecta el nombre del producto a partir del nombre del archivo.
    Ejemplo: "Impacto_cons_CO.csv" -> "consumo"

    Args:
        nombre_archivo: nombre del archivo CSV (ej: "Impacto_tarjeta_CO.csv")

    Returns:
        Nombre del producto normalizado

    Raises:
        ValueError si no se puede detectar el producto
    """
    match = re.search(r"Impacto_([a-zA-Z]+)", nombre_archivo)
    if not match:
        raise ValueError(f"No se pudo detectar el producto en: {nombre_archivo}")

    producto_raw = match.group(1).lower()
    return ALIAS_PRODUCTO.get(producto_raw, producto_raw)


def procesar_dataframe_impacto(df: pd.DataFrame, nombre_archivo: str) -> tuple:
    """
    Procesa un DataFrame de impacto:
    1. Detecta el producto desde el nombre del archivo
    2. Encuentra la columna Impacto_*
    3. La renombra al estandar Impacto_{producto}

    Args:
        df: DataFrame leido del CSV
        nombre_archivo: nombre del archivo (para detectar el producto)

    Returns:
        Tupla (df_procesado, aviso_si_renombrado, producto)

    Raises:
        ValueError si hay problemas con las columnas
    """
    producto = detectar_producto_desde_nombre(nombre_archivo)

    # Buscar columna de impacto
    col_impacto = [c for c in df.columns if c.startswith("Impacto_")]
    if len(col_impacto) != 1:
        raise ValueError(
            f"{nombre_archivo}: se esperaba 1 columna Impacto_*, "
            f"se encontraron {col_impacto}"
        )

    nuevo_nombre = f"Impacto_{producto}"
    col_actual = col_impacto[0]
    aviso = ""

    if col_actual != nuevo_nombre:
        aviso = f"Columna renombrada: '{col_actual}' -> '{nuevo_nombre}'"
        df = df.rename(columns={col_actual: nuevo_nombre})

    return df, aviso, producto


def concatenar_impactos(dataframes_dict: dict, 
                       pais: str = None,
                       fecha_format: str = "%d%b%Y") -> pd.DataFrame:
    """
    Concatena multiples DataFrames de impacto en uno solo.

    Args:
        dataframes_dict: diccionario {nombre_archivo: df_procesado, ...}
        pais: codigo de pais (para validacion, opcional)
        fecha_format: formato de la columna 'date'

    Returns:
        DataFrame consolidado, ordenado por KEYS_MERGE

    Raises:
        ValueError si los DataFrames no tienen las columnas requeridas
    """
    if not dataframes_dict:
        raise ValueError("No hay DataFrames para concatenar")

    # Validar que todas tengan las claves
    for nombre, df in dataframes_dict.items():
        faltantes = [k for k in KEYS_MERGE if k not in df.columns]
        if faltantes:
            raise ValueError(
                f"{nombre} no tiene las columnas requeridas: {faltantes}"
            )

    # Merge
    frames = list(dataframes_dict.values())
    df_result = reduce(
        lambda left, right: pd.merge(left, right, on=KEYS_MERGE, how="outer"),
        frames
    )

    # Convertir fecha
    df_result["date"] = pd.to_datetime(df_result["date"], format=fecha_format)

    # Ordenar
    df_result = df_result.sort_values(KEYS_MERGE).reset_index(drop=True)

    return df_result


def generar_csv_bytes(df: pd.DataFrame) -> bytes:
    """Convierte un DataFrame a bytes CSV (listo para download)."""
    buffer = BytesIO()
    df.to_csv(buffer, index=False, encoding="utf-8")
    buffer.seek(0)
    return buffer.getvalue()


# =============================================================================
# FLUJO COMPLETO
# =============================================================================

def procesar_archivos_impacto(archivos_cargados: list, 
                               pais: str = "CO",
                               fecha_format: str = "%d%b%Y") -> dict:
    """
    Flujo completo: recibe archivos, los procesa, los concatena.

    Args:
        archivos_cargados: lista de UploadedFile de Streamlit
        pais: codigo de pais para validacion
        fecha_format: formato de la columna 'date'

    Returns:
        Diccionario con:
        - 'df': DataFrame consolidado
        - 'avisos': lista de strings con warnings
        - 'errores': lista de strings con errores (si los hay)
        - 'archivos_procesados': lista de nombres
    """
    avisos = []
    errores = []
    dataframes_dict = {}
    archivos_procesados = []

    for uploaded_file in archivos_cargados:
        try:
            # Leer CSV
            df = pd.read_csv(uploaded_file)

            # Procesar
            df_proc, aviso, producto = procesar_dataframe_impacto(
                df, 
                uploaded_file.name
            )

            if aviso:
                avisos.append(aviso)

            dataframes_dict[uploaded_file.name] = df_proc
            archivos_procesados.append(uploaded_file.name)

        except Exception as e:
            errores.append(f"{uploaded_file.name}: {str(e)}")

    # Si no hay errores criticos, concatenar
    df_final = None
    if dataframes_dict:
        try:
            df_final = concatenar_impactos(dataframes_dict, pais=pais, fecha_format=fecha_format)
        except Exception as e:
            errores.append(f"Error en concatenacion: {str(e)}")
    else:
        errores.append("No se pudo procesar ningun archivo")

    return {
        'df': df_final,
        'avisos': avisos,
        'errores': errores,
        'archivos_procesados': archivos_procesados,
    }
