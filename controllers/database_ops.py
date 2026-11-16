import os
import json
import pandas as pd

DATA_JSON = 'data/data.json'
CARPETA_DATA = 'data'

# Variables globales para caché en memoria y evitar lecturas repetitivas al disco
_cache_secciones = None
_cache_mapeo = None
_cache_mtime_dict = {}

def _obtener_mtime_archivos():
    """Obtiene la fecha de modificación de los archivos Excel para controlar cambios."""
    mtimes = {}
    if os.path.exists(CARPETA_DATA):
        for archivo in os.listdir(CARPETA_DATA):
            if archivo.endswith('.xlsx'):
                ruta = os.path.join(CARPETA_DATA, archivo)
                try:
                    mtimes[archivo] = os.path.getmtime(ruta)
                except:
                    pass
    return mtimes

def leer_excel_inteligente(ruta_archivo, hoja):
    """Busca automáticamente en las primeras 15 filas la fila de encabezados 
    que contenga palabras clave como CÉDULA o CEDULA."""
    try:
        df_raw = pd.read_excel(ruta_archivo, sheet_name=hoja, header=None, nrows=15)
        header_row = 0
        for idx, row in df_raw.iterrows():
            row_str = " ".join([str(val).upper() for val in row.values if pd.notna(val)])
            if any(k in row_str for k in ['CEDULA', 'CÉDULA', 'Nº DE CEDULA', 'Nº CEDULA']):
                header_row = idx
                break
        
        df = pd.read_excel(ruta_archivo, sheet_name=hoja, header=header_row, dtype=str)
        return df
    except Exception as e:
        print(f"Error leyendo hoja {hoja} en {ruta_archivo}: {e}")
        return pd.DataFrame()

def cargar_estructura_dinamica():
    """Lee automáticamente todos los archivos y hojas de la carpeta data 
    utilizando caché en memoria para evitar lecturas lentas y repetidas al disco."""
    global _cache_secciones, _cache_mapeo, _cache_mtime_dict
    
    mtimes_actuales = _obtener_mtime_archivos()
    
    # Si la caché es válida y los archivos del disco no han cambiado, se devuelve de inmediato (Rendimiento instantáneo)
    if _cache_secciones is not None and _cache_mapeo is not None and mtimes_actuales == _cache_mtime_dict:
        return _cache_secciones, _cache_mapeo

    secciones_dict = {}
    mapeo_seccion_archivo = {}
    
    if not os.path.exists(CARPETA_DATA):
        os.makedirs(CARPETA_DATA)
        
    for archivo in os.listdir(CARPETA_DATA):
        if archivo.endswith('.xlsx'):
            ruta_archivo = os.path.join(CARPETA_DATA, archivo)
            try:
                excel_file = pd.ExcelFile(ruta_archivo)
                for hoja in excel_file.sheet_names:
                    df_temp = leer_excel_inteligente(ruta_archivo, hoja)
                    if df_temp.empty:
                        continue
                    columnas = [str(c).strip() if not str(c).startswith('Unnamed') else 'CAMPO SIN NOMBRE' for c in df_temp.columns]
                    
                    secciones_dict[hoja] = columnas
                    mapeo_seccion_archivo[hoja] = archivo
            except Exception as e:
                print(f"Error procesando estructura de {archivo}: {e}")
                
    # Actualizar la caché global
    _cache_secciones = secciones_dict
    _cache_mapeo = mapeo_seccion_archivo
    _cache_mtime_dict = mtimes_actuales
    
    return _cache_secciones, _cache_mapeo

def load_json():
    ruta_json = DATA_JSON
    usuarios = {"admin": "12345678"}
    
    if os.path.exists(ruta_json):
        try:
            with open(ruta_json, 'r', encoding='utf-8') as f:
                data = json.load(f)
                if "users" in data:
                    usuarios = data["users"]
        except:
            pass
            
    secciones_dict, _ = cargar_estructura_dinamica()
    
    return {
        "users": usuarios,
        "modules": {
            "REGISTRO GENERAL": {
                "archivo": "MULTIPLE",
                "secciones": secciones_dict if secciones_dict else {"Hoja1": ["CÉDULA"]}
            }
        }
    }

def obtener_ruta_excel_por_seccion(section_name):
    """Ubica exactamente a qué archivo Excel pertenece cada sección (hoja)."""
    _, mapeo = cargar_estructura_dinamica()
    if section_name in mapeo:
        return os.path.join(CARPETA_DATA, mapeo[section_name])
    return None

def obtener_ruta_excel(module_name):
    _, mapeo = cargar_estructura_dinamica()
    if mapeo:
        primer_archivo = list(mapeo.values())[0]
        return os.path.join(CARPETA_DATA, primer_archivo)
    return os.path.join(CARPETA_DATA, 'DATA_GENERAL.xlsx')

def init_excel():
    if not os.path.exists(CARPETA_DATA):
        os.makedirs(CARPETA_DATA)
    
    archivos_existentes = [f for f in os.listdir(CARPETA_DATA) if f.endswith('.xlsx')]
    if not archivos_existentes:
        ruta_base = os.path.join(CARPETA_DATA, 'DATA_GENERAL.xlsx')
        with pd.ExcelWriter(ruta_base, engine='openpyxl') as writer:
            df = pd.DataFrame(columns=["CÉDULA", "APELLIDOS Y NOMBRES", "CARGO"])
            df.to_excel(writer, sheet_name="EMPLEADOS ACTIVOS", index=False)

def format_number(val):
    try:
        return f"{int(val):,}".replace(',', '.')
    except:
        return str(val)