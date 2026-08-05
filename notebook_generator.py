import json
import copy
from typing import Dict, Tuple
import unicodedata


# ==================== MAPEOS ====================

PAIS_MAP = {
    "Colombia": "CO",
    "Panamá": "PA",
    "Costa Rica": "CR",
}

CARTERA_MAP = {
    "vivienda": "vivi",
    "consumo": "cons",
    "tarjeta": "tc",
    "vehículo": "vehic",
    "corporativo": "corp",      # ← código propio, independiente de comercial
    "pymes": "pyme",
    "comercial": "comercial",  # ← mantiene su propio código
    "hipotecas": "hipo",
}

CARTERA_LABEL = {
    "vivi": "vivienda",
    "cons": "consumo",
    "tc": "tarjeta",
    "vehic": "vehículo",
    "corp": "corporativo",      # ← desmapea corp → corporativo
    "comercial": "comercial",   # ← desmapea comercial → comercial
    "pyme": "pymes",
    "hipo": "hipotecas",
}

IMPACTO_LABEL_MAP = {
    "vivi": "vivienda",
    "cons": "consumo",
    "tc": "tarjeta",
    "vehic": "vehiculo",
    "corp": "corporativo",      # ← impacto para corporativo
    "comercial": "comercial",   # ← impacto para comercial
    "pyme": "pyme",
    "hipo": "hipotecas",
}


# ==================== GENERADOR ====================

def _normalizar_texto(texto: str) -> str:
    if texto is None:
        return ""
    texto = texto.strip().lower()
    texto = "".join(
        c for c in unicodedata.normalize("NFD", texto)
        if unicodedata.category(c) != "Mn"
    )
    return texto.replace(" ", "_")


def _resolver_pais_codigo(pais: str) -> str:
    clave = _normalizar_texto(pais)
    if len(clave) == 2:
        return clave.upper()
    for nombre, codigo in PAIS_MAP.items():
        if _normalizar_texto(nombre) == clave:
            return codigo
    return pais.upper()


def _resolver_cartera_codigo(cartera: str) -> str:
    clave = _normalizar_texto(cartera)

    if clave in CARTERA_LABEL:
        return clave

    for nombre, codigo in CARTERA_MAP.items():
        if _normalizar_texto(nombre) == clave:
            return codigo

    return clave


def _resolver_cartera_label(cartera: str) -> str:
    codigo = _resolver_cartera_codigo(cartera)
    return CARTERA_LABEL.get(codigo, _normalizar_texto(cartera))

def _resolver_cartera_impacto(cartera: str) -> str:
    codigo = _resolver_cartera_codigo(cartera)
    return IMPACTO_LABEL_MAP.get(codigo, _normalizar_texto(cartera))


def generar_nombres_archivos(pais: str, cartera: str) -> Dict[str, str]:
    """Genera los 4 nombres de archivos para el generador."""
    pais_lower = _resolver_pais_codigo(pais).lower()
    cartera_codigo = _resolver_cartera_codigo(cartera)
    return {
        "hist": f"hist_{cartera_codigo}_{pais_lower}_fwl_pd12.xlsx",
        "opt": f"opt_macro_davi__{cartera_codigo}_{pais_lower}_fwl_pd12.xlsx",
        "base": f"bas_macro_davi_{cartera_codigo}_{pais_lower}_fwl_pd12.xlsx",
        "adv": f"adv_macro_davi__{cartera_codigo}_{pais_lower}_fwl_pd12.xlsx",
    }


def reemplazar_en_celda(source: str, reemplazos: Dict[str, str]) -> str:
    """
    Reemplaza valores en una celda de código.
    Los reemplazos son: {"VARIABLE": "nuevo_valor"}
    """
    import re
    
    lines = source.split('\n')
    new_lines = []
    
    for line in lines:
        # Saltar comentarios puros
        if line.strip().startswith('#'):
            new_lines.append(line)
            continue
        
        # Para cada variable a reemplazar
        line_modified = False
        for var, valor in reemplazos.items():
            # Buscar: VAR (con espacios flexibles) = 
            pattern = rf'^(\s*){var}\s*='
            if re.search(pattern, line) and not line_modified:
                # Conservar indentación
                indent = len(line) - len(line.lstrip())
                # Extraer comentario si existe
                if '#' in line:
                    comment_part = line[line.index('#'):]
                    line = ' ' * indent + f'{var} = {valor}  {comment_part}'
                else:
                    line = ' ' * indent + f'{var} = {valor}'
                line_modified = True
        
        new_lines.append(line)
    
    return '\n'.join(new_lines)


def _ensure_source_list(cell_source):
    """
    Asegura que el source de una celda sea una lista de strings.
    Los notebooks Jupyter requieren source como lista de strings.
    """
    if isinstance(cell_source, list):
        # Si ya es lista, asegurar que cada elemento termine con \n excepto el último
        result = []
        for i, line in enumerate(cell_source):
            if i < len(cell_source) - 1:
                if not line.endswith('\n'):
                    line = line + '\n'
            result.append(line)
        return result
    elif isinstance(cell_source, str):
        # Si es string, convertir a lista preservando saltos de línea
        lines = cell_source.split('\n')
        result = []
        for i, line in enumerate(lines):
            if i < len(lines) - 1:
                result.append(line + '\n')
            else:
                # Última línea: solo agregar \n si el original terminaba con \n
                result.append(line)
        return result
    else:
        return [str(cell_source)]


def generar_notebook_generador(
    template_path: str,
    pais: str,
    cartera: str,
    fecha_inicio: str,
    fecha_fin: str,
    modo_endo: str,
    ventana_mm: int,
    nombres_custom: Dict[str, str] = None,
) -> Dict:
    """
    Carga el notebook template del generador y reemplaza los parámetros.
    """
    with open(template_path, 'r', encoding='utf-8') as f:
        nb = json.load(f)

    cartera_codigo = _resolver_cartera_codigo(cartera)
    
    # Genera nombres de archivos
    if nombres_custom:
        archivos = nombres_custom
    else:
        archivos = generar_nombres_archivos(pais, cartera)
    
    # Parámetros a reemplazar en la CELDA 4 (configuración)
    reemplazos_celda4 = {
        'PAIS': f'"{pais}"',
        'CARTERA': f'"{cartera_codigo}"',
        'FECHA_INICIO': f'"{fecha_inicio}"',
        'FECHA_FIN_HIST': f'"{fecha_fin}"',
        'MODO_ENDO': f'"{modo_endo}"',
        'VENTANA_MEDIA_MOVIL': str(ventana_mm),
        'NOMBRE_ARCHIVO_HIST': f'"{archivos["hist"]}"',
        'NOMBRE_ARCHIVO_OPT': f'"{archivos["opt"]}"',
        'NOMBRE_ARCHIVO_BASE': f'"{archivos["base"]}"',
        'NOMBRE_ARCHIVO_ADV': f'"{archivos["adv"]}"',
    }
    
    # Reemplaza en celda 4
    for i, cell in enumerate(nb['cells']):
        if i == 4 and cell['cell_type'] == 'code':
            source = ''.join(cell['source']) if isinstance(cell['source'], list) else cell['source']
            source = reemplazar_en_celda(source, reemplazos_celda4)
            nb['cells'][i]['source'] = _ensure_source_list(source)
    
    # Reemplaza en celda de carga de exógenas (sin depender de índice fijo)
    exo_reemplazado = False
    endo_reemplazado = False
    for i, cell in enumerate(nb['cells']):
        if cell['cell_type'] != 'code':
            continue
        source = ''.join(cell['source']) if isinstance(cell['source'], list) else cell['source']
        lines = source.split('\n')
        new_lines = []
        line_changed = False
        for line in lines:
            if 'Variables Macro FWL' in line and 'read_excel' in line:
                new_lines.append(f'exo = pd.read_excel(\'Variables Macro FWL {pais}.xlsx\')')
                line_changed = True
                exo_reemplazado = True
            elif 'test_pd_lra_' in line and 'read_csv' in line:
                new_lines.append(f"endo = pd.read_csv('test_pd_lra_{cartera_codigo}_01mar26.csv')")
                line_changed = True
                endo_reemplazado = True
            else:
                new_lines.append(line)
        if line_changed:
            source = '\n'.join(new_lines)
            nb['cells'][i]['source'] = _ensure_source_list(source)
        if exo_reemplazado and endo_reemplazado:
            break
    
    return nb


def generar_notebook_motor(
    template_path: str,
    pais: str,
    cartera: str,
    tipo_modelo: str,
    max_lags: int,
    nombres_archivos: Dict[str, str],
) -> Dict:
    """
    Carga el notebook template del motor y reemplaza los parámetros.
    """
    with open(template_path, 'r', encoding='utf-8') as f:
        nb = json.load(f)

    pais_codigo = _resolver_pais_codigo(pais)
    cartera_codigo = _resolver_cartera_codigo(cartera)
    cartera_impacto = _resolver_cartera_impacto(cartera)
    
    # Parámetros a reemplazar en la CELDA 4 (identificación endógena)
    reemplazos_celda4 = {
        'ARCHIVO_ENDOGENA': f'"{nombres_archivos["hist"]}"',
        'ARCHIVO_EXO_BAS': f'"{nombres_archivos["base"]}"',
        'ARCHIVO_EXO_ADV': f'"{nombres_archivos["adv"]}"',
        'ARCHIVO_EXO_OPT': f'"{nombres_archivos["opt"]}"',
    }
    
    for i, cell in enumerate(nb['cells']):
        if i == 4 and cell['cell_type'] == 'code':
            source = ''.join(cell['source']) if isinstance(cell['source'], list) else cell['source']
            source = reemplazar_en_celda(source, reemplazos_celda4)
            nb['cells'][i]['source'] = _ensure_source_list(source)
    
    # Parámetros en la CELDA 9 (motor de modelos SARIMAX)
    reemplazos_celda9 = {
        'TIPO_ENDOGENA': f'"{tipo_modelo}"',
        'MAX_LAGS': str(max_lags),
        'ARCHIVO_SALIDA': f'"Modelos_{pais_codigo}_{cartera_codigo}.xlsx"',
        'PAIS_REPORTE': f'"{pais_codigo}"',
        'ARCHIVO_REPORTE': f'"Impacto_{cartera_codigo}_{pais_codigo}.csv"',
        'COLUMNA_IMPACTO': f'"Impacto_{cartera_impacto}"',
    }
    
    for i, cell in enumerate(nb['cells']):
        if i == 9 and cell['cell_type'] == 'code':
            source = ''.join(cell['source']) if isinstance(cell['source'], list) else cell['source']
            source = reemplazar_en_celda(source, reemplazos_celda9)
            nb['cells'][i]['source'] = _ensure_source_list(source)
    
    return nb


def guardar_notebook(nb: Dict, output_path: str) -> None:
    """Guarda un notebook a disco."""
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(nb, f, ensure_ascii=False, indent=1)


# ==================== METADATA EMBEBIDA ====================

def escribir_meta_embebida(file_path: str, meta: Dict, prefix: str = "sarimax_meta") -> None:
    """
    Escribe metadata embebida en un archivo Excel usando custom properties.
    
    Args:
        file_path: ruta al archivo Excel
        meta: diccionario con metadata (pais, cartera, modo_endogena, ventana_mm, tipo_modelo, etc.)
        prefix: prefijo para las propiedades custom
    """
    import openpyxl
    
    wb = openpyxl.load_workbook(file_path)
    
    # Convertir metadata a JSON string
    meta_json = json.dumps(meta, ensure_ascii=False)
    
    # Dividir en chunks de max 255 chars (límite de custom props en Excel)
    chunk_size = 255
    chunks = [meta_json[i:i+chunk_size] for i in range(0, len(meta_json), chunk_size)]
    
    # Guardar número de chunks
    wb.custom_doc_props.add(f"{prefix}_n", len(chunks), "int")
    
    # Guardar cada chunk
    for idx, chunk in enumerate(chunks, 1):
        prop_name = f"{prefix}_{idx:02d}"
        # Eliminar si ya existe
        if prop_name in wb.custom_doc_props.names:
            del wb.custom_doc_props[prop_name]
        wb.custom_doc_props.add(prop_name, chunk, "str")
    
    wb.save(file_path)
    wb.close()


def generar_meta_dict(
    pais: str,
    cartera: str,
    modo_endogena: str = None,
    ventana_mm: int = None,
    tipo_modelo: str = None,
    vif_max: float = None,
    fwl_factor_min: float = None,
    fwl_factor_max: float = None,
    max_exog_por_modelo: int = None,
) -> Dict:
    """
    Genera un diccionario de metadata estandarizado para los notebooks.
    """
    return {
        "pais": pais,
        "cartera": cartera,
        "generador_modo_endogena": modo_endogena,
        "generador_ventana_mm": ventana_mm,
        "motor_tipo_endogena": tipo_modelo,
        "motor_vif_max": vif_max,
        "motor_fwl_factor_min": fwl_factor_min,
        "motor_fwl_factor_max": fwl_factor_max,
        "motor_max_exog_por_modelo": max_exog_por_modelo,
    }
