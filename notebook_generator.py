import json
import copy
from typing import Dict, Tuple


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
    "corporativo": "comercial",
    "pymes": "pyme",
    "comercial": "comercial",
    "hipotecas": "hipo",
}

CARTERA_LABEL = {
    "vivi": "vivienda",
    "cons": "consumo",
    "tc": "tarjeta",
    "vehic": "vehículo",
    "comercial": "corporativo",
    "pyme": "pymes",
    "hipo": "hipotecas",
}


# ==================== GENERADOR ====================

def generar_nombres_archivos(pais: str, cartera: str) -> Dict[str, str]:
    """Genera los 4 nombres de archivos para el generador."""
    pais_lower = pais.lower()
    return {
        "hist": f"hist_{cartera}_{pais_lower}_fwl_pd12.xlsx",
        "opt": f"opt_macro_davi__{cartera}_{pais_lower}_fwl_pd12.xlsx",
        "base": f"bas_macro_davi_{cartera}_{pais_lower}_fwl_pd12.xlsx",
        "adv": f"adv_macro_davi__{cartera}_{pais_lower}_fwl_pd12.xlsx",
    }


def reemplazar_en_celda(source: str, reemplazos: Dict[str, str]) -> str:
    """
    Reemplaza valores en una celda de código.
    Los reemplazos son: {"VARIABLE": "nuevo_valor"}
    
    Ejemplo:
      reemplazos = {
        "PAIS": '"PA"',
        "MAX_LAGS": "5",
      }
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
            # Usar regex para ser flexible con espacios
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
    
    Args:
        template_path: ruta al Generacion_Variacion__2_.ipynb
        pais: código país (CO, PA, CR)
        cartera: código cartera (vivi, cons, tc, etc.)
        fecha_inicio: YYYY-MM-DD
        fecha_fin: YYYY-MM-DD
        modo_endo: "actual" o "media_movil"
        ventana_mm: número (1-12)
        nombres_custom: dict con keys "hist", "opt", "base", "adv" si quiere overridear
    
    Returns:
        Notebook modificado (dict)
    """
    with open(template_path, 'r', encoding='utf-8') as f:
        nb = json.load(f)
    
    # Genera nombres de archivos
    if nombres_custom:
        archivos = nombres_custom
    else:
        archivos = generar_nombres_archivos(pais, cartera)
    
    # Parámetros a reemplazar en la CELDA 4 (configuración)
    reemplazos_celda4 = {
        'PAIS': f'"{pais}"',
        'CARTERA': f'"{cartera}"',
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
            # Convertir source a string si es lista
            if isinstance(cell['source'], list):
                source = ''.join(cell['source'])
            else:
                source = cell['source']
            
            # Hacer el reemplazo
            source = reemplazar_en_celda(source, reemplazos_celda4)
            
            # Guardar de vuelta como lista (formato de notebook)
            nb['cells'][i]['source'] = source.split('\n')
    
    # Reemplaza en celda 7 (carga de exógenas)
    for i, cell in enumerate(nb['cells']):
        if i == 7 and cell['cell_type'] == 'code':
            if isinstance(cell['source'], list):
                source = ''.join(cell['source'])
            else:
                source = cell['source']
            
            # Buscar y reemplazar la línea de lectura de archivo
            lines = source.split('\n')
            new_lines = []
            for line in lines:
                if 'Variables Macro FWL' in line and 'read_excel' in line:
                    new_lines.append(f'exo = pd.read_excel(\'Variables Macro FWL {pais}.xlsx\')')
                else:
                    new_lines.append(line)
            source = '\n'.join(new_lines)
            nb['cells'][i]['source'] = source.split('\n')
    
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
    
    Args:
        template_path: ruta al Motor_Sarimax_Vivi_CO__1_.ipynb
        pais: código país (CO, PA, CR)
        cartera: código cartera (vivi, cons, tc, etc.)
        tipo_modelo: "total" o "logit"
        max_lags: número (1-12)
        nombres_archivos: dict con keys "hist", "opt", "base", "adv"
    
    Returns:
        Notebook modificado (dict)
    """
    with open(template_path, 'r', encoding='utf-8') as f:
        nb = json.load(f)
    
    pais_lower = pais.lower()
    
    # Parámetros a reemplazar en la CELDA 4 (identificación endógena)
    reemplazos_celda4 = {
        'ARCHIVO_ENDOGENA': f'"{nombres_archivos["hist"]}"',
        'ARCHIVO_EXO_BAS': f'"{nombres_archivos["base"]}"',
        'ARCHIVO_EXO_ADV': f'"{nombres_archivos["adv"]}"',
        'ARCHIVO_EXO_OPT': f'"{nombres_archivos["opt"]}"',
    }
    
    for i, cell in enumerate(nb['cells']):
        if i == 4 and cell['cell_type'] == 'code':
            if isinstance(cell['source'], list):
                source = ''.join(cell['source'])
            else:
                source = cell['source']
            
            source = reemplazar_en_celda(source, reemplazos_celda4)
            nb['cells'][i]['source'] = source.split('\n')
    
    # Parámetros en la CELDA 9 (motor de modelos SARIMAX)
    cartera_label = CARTERA_LABEL.get(cartera, cartera)
    
    reemplazos_celda9 = {
        'TIPO_ENDOGENA': f'"{tipo_modelo}"',
        'MAX_LAGS': str(max_lags),
        'ARCHIVO_SALIDA': f'"Modelos_{pais}_{cartera}.xlsx"',
        'PAIS_REPORTE': f'"{pais}"',
        'ARCHIVO_REPORTE': f'"Impacto_{cartera}_{pais}.csv"',
    }
    
    for i, cell in enumerate(nb['cells']):
        if i == 9 and cell['cell_type'] == 'code':
            if isinstance(cell['source'], list):
                source = ''.join(cell['source'])
            else:
                source = cell['source']
            
            source = reemplazar_en_celda(source, reemplazos_celda9)
            nb['cells'][i]['source'] = source.split('\n')
    
    return nb


def guardar_notebook(nb: Dict, output_path: str) -> None:
    """Guarda un notebook a disco."""
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(nb, f, ensure_ascii=False, indent=1)
