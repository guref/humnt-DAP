import hashlib
# Parche de compatibilidad para Python 3.8 y ReportLab
_orig_md5 = hashlib.md5
def _patched_md5(*args, **kwargs):
    kwargs.pop('usedforsecurity', None)
    return _orig_md5(*args, **kwargs)
hashlib.md5 = _patched_md5

from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file
import os
import json
import shutil
import unicodedata
import pandas as pd
from datetime import datetime
from functools import wraps
from controllers.database_ops import load_json, init_excel, format_number, obtener_ruta_excel, obtener_ruta_excel_por_seccion, CARPETA_DATA
from controllers.pdf_logic import generate_section_pdf, generate_full_module_pdf, generate_global_search_pdf_personalizado
from controllers.database_ops import load_json, init_excel, format_number, obtener_ruta_excel, obtener_ruta_excel_por_seccion, CARPETA_DATA
from controllers.pdf_logic import generate_section_pdf, generate_full_module_pdf, generate_global_search_pdf_personalizado
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

app = Flask(__name__)
app.secret_key = 'clave_secreta_segura'

JSON_PATH = os.path.join(CARPETA_DATA, 'data.json')
JSON_DB_PATH = os.path.join(CARPETA_DATA, 'base_empleados.json')
SOLICITUDES_PATH = os.path.join(CARPETA_DATA, 'solicitudes.json')
# Archivo exclusivo y aislado para la recuperación de contraseña
SOLICITUDES_RECUPERACION_PATH = os.path.join(CARPETA_DATA, 'solicitudes_recuperacion.json')

def cargar_solicitudes_recuperacion():
    if os.path.exists(SOLICITUDES_RECUPERACION_PATH):
        try:
            with open(SOLICITUDES_RECUPERACION_PATH, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            return []
    return []

def guardar_solicitudes_recuperacion(sols):
    os.makedirs(CARPETA_DATA, exist_ok=True)
    with open(SOLICITUDES_RECUPERACION_PATH, 'w', encoding='utf-8') as f:
        json.dump(sols, f, ensure_ascii=False, indent=4)

init_excel()

UBICACIONES_ESPECIFICAS = [
    'AMBULATORIO URBANO TIPO III "DR. JOSÉ ANTONIO SERRES" DEL ESTADO MONAGAS',
    'AMBULATORIO URBANO TIPO III "DR. JOSE MARIA VARGAS" DEL ESTADO MONAGAS',
    'CENTRO CARDIOVASCULAR ORIENTAL "DR. MIGUEL HERNANDEZ" DEL ESTADO MONAGAS',
    'DIRECCION DE SALUD AMBIENTAL DEL ESTADO MONAGAS',
    'DIRECCION REGIONAL DE SALUD DEL ESTADO MONAGAS',
    'DIRECCION REGIONAL SALUD (CENTRALIZADA)',
    'HOSPITAL GENERAL TIPO I "DR. ERNESTO GUZMAN SAAVEDRA", CAICARA DEL ESTADO MONAGAS',
    'HOSPITAL GENERAL TIPO I "DR. LUIS GONZALEZ ESPINOZA", PUNTA DE MATA DEL ESTADO MONAGAS',
    'HOSPITAL GENERAL TIPO I "DR. JOSE ANTONIO URRESTARAZU", CARIPE DEL ESTADO MONAGAS',
    'HOSPITAL GENERAL TIPO I "DR. PABLO VILLARROEL", SAN ANTONIO DE MATURIN DEL ESTADO MONAGAS',
    'HOSPITAL GENERAL TIPO I "DR. TULIO LOPEZ RAMIREZ", BARRANCAS DEL ESTADO MONAGAS',
    'HOSPITAL GENERAL TIPO I "DRA. ANA TORREALBA", TEMBLADOR DEL ESTADO MONAGAS',
    'HOSPITAL GENERAL TIPO I "DRA. ELVIRA BUENO MEZA", ARAGUA DE MATURIN DEL ESTADO MONAGAS',
    'HOSPITAL GENERAL TIPO I "SIMON BOLIVAR" DEL ESTADO MONAGAS',
    'HOSPITAL GENERAL TIPO II "DARIO MARQUEZ", CARIPITO DEL ESTADO MONAGAS',
    'HOSPITAL PSIQUIATRICO "DR. LUIS DANIEL BEAUPERTHUY" DEL ESTADO MONAGAS',
    'HOSPITAL UNIVERSITARIO TIPO IV "DR. MANUEL NUÑEZ TOVAR" DEL ESTADO MONAGAS',
    'UNIDAD SANITARIA DE MATURIN DEL ESTADO MONAGAS',
    'INSAMONAGAS (CENTRALIZADA)'
]

def cargar_solicitudes():
    if os.path.exists(SOLICITUDES_PATH):
        try:
            with open(SOLICITUDES_PATH, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            return []
    return []

def guardar_solicitudes(solicitudes):
    os.makedirs(CARPETA_DATA, exist_ok=True)
    with open(SOLICITUDES_PATH, 'w', encoding='utf-8') as f:
        json.dump(solicitudes, f, ensure_ascii=False, indent=4)

def verificar_migracion_usuarios():
    try:
        if os.path.exists(JSON_PATH):
            with open(JSON_PATH, 'r', encoding='utf-8') as f:
                data = json.load(f)
        else:
            data = {"users": {}, "modules": {}}
        modificado = False
        if 'users' not in data:
            data['users'] = {}
            modificado = True
        data['users']['admin'] = {"password": "12345678", "role": "admin"}
        modificado = True
        for usr, info in data.get('users', {}).items():
            if usr == 'admin': continue
            if isinstance(info, str):
                data['users'][usr] = {"password": info, "role": "operator"}
                modificado = True
            elif isinstance(info, dict) and 'password' not in info:
                info['password'] = "12345678"
                modificado = True
        if modificado:
            with open(JSON_PATH, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=4)
    except Exception as e:
        print(f"Nota en migración de usuarios: {e}")

verificar_migracion_usuarios()

def admin_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user' not in session or session.get('role') != 'admin':
            flash('Acceso denegado. Esta sección es exclusiva para el Administrador del sistema.', 'danger')
            return redirect(url_for('dashboard'))
        return f(*args, **kwargs)
    return decorated_function

def registrar_auditoria(usuario, accion):
    log_path = os.path.join(CARPETA_DATA, 'audit.log')
    tiempo_actual = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    try:
        with open(log_path, 'a', encoding='utf-8') as f:
            f.write(f"[{tiempo_actual}] - Usuario: {usuario} - Acción: {accion}\n")
    except Exception as e:
        print(f"Error escribiendo log: {e}")

def leer_excel_inteligente(ruta_archivo, hoja=0):
    try:
        df_raw = pd.read_excel(ruta_archivo, sheet_name=hoja, header=None, nrows=12)
        if df_raw.empty:
            return pd.DataFrame()
        
        header_row_idx = 0
        for r_idx in range(min(10, len(df_raw))):
            row_vals = [str(val).strip().upper() for val in df_raw.iloc[r_idx].values]
            if any('CEDULA' in v or 'CÉDULA' in v or 'Nº DE CEDULA' in v for v in row_vals):
                header_row_idx = r_idx
                break
                
        # Leemos el Excel asegurándonos de que todo se interprete inicialmente sin restricciones de tipo estricto
        df = pd.read_excel(ruta_archivo, sheet_name=hoja, header=header_row_idx)
        df = df.dropna(how='all')
        
        # =====================================================================
        # SANITIZACIÓN UNIVERSAL A TEXTO PLANO (Str)
        # =====================================================================
        for col in df.columns:
            # 1. Remplazar nulos reales de pandas por texto limpio de forma segura
            df[col] = df[col].fillna('Sin registro')
            
            # 2. Convertir a string y limpiar espacios en blanco extremos
            df[col] = df[col].astype(str).str.strip()
            
            # 3. Limpiar valores textuales nulos comunes que vengan como strings
            df[col] = df[col].replace(['nan', 'NaT', 'NAT', 'NaN', 'NONE', 'None', ''], 'Sin registro')
            
            # 4. Limpiar artefactos de Excel (solo si la columna no requiere decimales estrictos)
            # Elimina el .0 típico de enteros leídos como float en pandas
            df[col] = df[col].str.replace(r'\.0$', '', regex=True)
            
            # Limpia la hora por defecto de los Timestamps en fechas
            df[col] = df[col].str.replace(' 00:00:00', '', regex=True)

        return df
    except Exception as e:
        print(f"Error en leer_excel_inteligente para {ruta_archivo} ({hoja}): {e}")
        return pd.DataFrame()
    
def normalizar_nombre_col_global(c, nombre_seccion=""):
    c_str = str(c).replace('\r', ' ').replace('\n', ' ').strip().replace('°', 'º')
    c_str = " ".join(c_str.split())
        
    if c_str.startswith('Unnamed'):
        return 'CÉDULA' if 'CEDULA' in c_str.upper() or c_str in ['Unnamed: 1', 'Unnamed: 2'] else None
        
    if c_str.replace('.', '').isdigit() or not c_str:
        return None  
        
    c_up = ''.join(ch for ch in unicodedata.normalize('NFD', c_str.upper()) if unicodedata.category(ch) != 'Mn')
    c_up = c_up.replace('°', 'º').strip()
    
    palabras_excluidas = [
        'GAVETERO', 'HOJA', 'SHEET', 'INFORME', 'CONSOLIDADO', 
        'HORARIO MATERNAL', 'HORARIO LABORAL', 'INFRAMONAGAS',
        'COTIZACION', 'COTIZACIÓN', 'SERVICIO DONDE LABORA', 'SERVICIO DONE LABORA',
        'ESPECIALIDAD', 'DIAGNOSTICO DEL MEDICO', 'DIAGNÓSTICO DEL MÉDICO'
    ]
    if any(p in c_up for p in palabras_excluidas):
        return None

    if c_up in ['Nº', 'N°', 'N', 'ITEM', '#'] or c_up.startswith('Nº '):
        return None

    if any(k in c_up for k in ['CARGO', 'DENOMINACION', 'DENOMINACIÓN CARGO', 'PUESTO', 'CARGO NOMINAL', 'OCUPACION', 'FUNCION']):
        return 'CARGO'

    if any(k in c_up for k in ['NOMBRE', 'APELLIDO', 'TRABAJADOR', 'EMPLEADO', 'FUNCIONARIO', 'APELLIDOS Y NOMBRES', 'NOMBRE Y APELLIDO', 'NOMBRES Y APELLIDO', 'NOMBRE Y APELLIDOS', 'NOMBRES Y APELLIDOS']):
        return 'NOMBRES Y APELLIDOS'

    tokens_col = c_up.split()
    es_cedula = (
        any(k in c_up for k in ['CEDULA', 'CÉDULA', 'Nº DE CEDULA', 'NUMERO DE CEDULA', 'NRO DE CEDULA']) or
        any(token in ['C.I', 'C.I.', 'CI', 'C.I:', 'CI:'] for token in tokens_col)
    ) and not 'UBICACION' in c_up

    if es_cedula:
        return 'CÉDULA'

    # 2. Validar INGRESO con términos específicos (sin palabras sueltas como 'AÑO' o 'ANO')
    if any(k in c_up for k in ['ING NOMINAL', 'FECHA DE ING', 'ANO ING', 'AÑO ING', 'FECHA ING', 'FECHA DE ING NOMINAL', 'AÑO DE INGRESO', 'FECHA DE INGRESO', 'INGRESO', 'F. ING']):
        return 'AÑO DE INGRESO'

    # 1. Validar primero EGRESO para evitar colisiones con la palabra "AÑO"
    if any(k in c_up for k in ['EGRESO', 'AÑO EGR', 'AÑO DE EGRESO', 'FECHA EGRESO']):
        return 'AÑO DE EGRESO'

    if c_up in ['PERSONAL', 'TIPO PERSONAL', 'TIPO DE PERSONAL', 'TIPO TRABAJADOR'] or ('TIPO' in c_up and 'PERSONAL' in c_up):
        return 'TIPO PERSONAL'
    if 'GRADO' in c_up:
        return 'GRADO'
    if any(k in c_up for k in ['DEPENDENCIA', 'DEPARTAMENTO', 'UNIDAD', 'DIRECCION', 'COORD']):
        return 'DEPENDENCIA'
    if 'UBICACION' in c_up:
        return 'UBICACIÓN NOMINAL'
    if any(k in c_up for k in ['TELEFONO', 'TELF', 'TLF', 'MOVIL', 'CELULAR']):
        return 'TELÉFONO'
    if any(k in c_up for k in ['OBSERVACION', 'DIAGNOSTICO', 'NOTA']):
        return 'OBSERVACIONES'
    if any(k in c_up for k in ['PERIODO', 'LAPSO', 'CICLO', 'PERIODO DE VACACION']):
        return 'PERÍODO'
    if any(k in c_up for k in ['DISFRUTE', 'ESTATUS', 'VACACION', 'VACACIONES', 'SITUACION']):
        return 'ESTATUS / DISFRUTE'
    if any(k in c_up for k in ['DESDE', 'FECHA DESDE', 'F. DESDE']):
        return 'DESDE'
    if any(k in c_up for k in ['HASTA', 'FECHA HASTA', 'F. HASTA']):
        return 'HASTA'
    if any(k in c_up for k in ['DIAS', 'DÍAS']):
        return 'DIAS'
    if any(k in c_up for k in ['FECHA DE NAC', 'FECHA NAC', 'F. NAC', 'FECHA DE NACIMIENTO']):
        return 'FECHA DE NACIMIENTO'
    if any(k in c_up for k in ['OBRERO', 'JUBILADO', 'PENSIONADO', 'SOBREVIVIENTE']):
        if 'JUBILADO' in c_up:
            return 'OBREROS JUBILADOS' if 'OBRERO' in c_up else 'JUBILADOS'
        if 'OBRERO' in c_up:
            return 'OBREROS'
        return c_up
    if 'MEDICO RESIDENTE' in c_up:
        return 'MÉDICOS RESIDENTES GAV 8'
    if 'EMPLEADOS OTRAS INST' in c_up:
        return 'EMPLEADOS OTRAS INST.'
        
    return c_up

def limpiar_texto_comparacion(texto):
    if not texto:
        return ""
    texto_norm = ''.join(ch for ch in unicodedata.normalize('NFD', str(texto).upper()) if unicodedata.category(ch) != 'Mn')
    for char in ['"', '“', '”', "'", '.', ',', '-']:
        texto_norm = texto_norm.replace(char, ' ')
    return " ".join(texto_norm.split())

# ==========================================
# SINCRONIZACIÓN EXCEL -> JSON (GARANTIZA LA CREACIÓN DE base_empleados.json)
# ==========================================
def sincronizar_excel_a_json():
    resultados = []
    if not os.path.exists(CARPETA_DATA):
        return

    seen_records = set()
    archivos = [f for f in os.listdir(CARPETA_DATA) if f.endswith('.xlsx') and not f.startswith('~$')]

    for archivo in archivos:
        ruta_archivo = os.path.join(CARPETA_DATA, archivo)
        try:
            excel_file = pd.ExcelFile(ruta_archivo)
            for hoja in excel_file.sheet_names:
                df = leer_excel_inteligente(ruta_archivo, hoja)
                if df.empty:
                    continue

                mapa_columnas = {}
                for c in df.columns:
                    norm_c = normalizar_nombre_col_global(c, hoja)
                    if norm_c:
                        mapa_columnas[c] = norm_c

                for idx, row in df.iterrows():
                    cedula_encontrada = ""
                    for c in df.columns:
                        val = row[c]
                        if pd.isna(val):
                            continue
                        val_str = str(val).strip()
                        if val_str.endswith('.0'):
                            val_str = val_str[:-2]
                        val_str = val_str.replace('.', '')
                        
                        if val_str.isdigit() and 5 <= len(val_str) <= 9:
                            cedula_encontrada = val_str
                            break

                    if not cedula_encontrada:
                        continue

                    record_key = f"{archivo}_{hoja}_{cedula_encontrada}_{idx}"
                    if record_key in seen_records:
                        continue
                    seen_records.add(record_key)

                    fila_datos = {'CÉDULA': cedula_encontrada}
                    for c_orig, c_norm in mapa_columnas.items():
                        val = row[c_orig]
                        if pd.isna(val) or str(val).lower() == 'nan':
                            val_str = "Sin registro"
                        else:
                            val_str = str(val).strip()
                            if val_str.endswith('.0'):
                                val_str = val_str[:-2]
                            if '00:00:00' in val_str:
                                val_str = val_str.split(' ')[0]
                            if not val_str:
                                val_str = "Sin registro"
                        fila_dict_val = fila_datos.get(c_norm, "Sin registro")
                        if fila_dict_val == "Sin registro":
                            fila_datos[c_norm] = val_str
                        elif val_str != "Sin registro":
                            fila_datos[c_norm] = val_str

                    resultados.append({
                        'archivo': archivo,
                        'seccion': hoja,
                        'fila': idx + 1,
                        'datos_fila': fila_datos
                    })
        except Exception as e:
            print(f"Error procesando {archivo} ({hoja}): {e}")

    empleados_consolidados = {}
    for res in resultados:
        fila = res.get('datos_fila', {})
        cedula_val = str(fila.get('CÉDULA', '')).strip()
        if not cedula_val:
            continue

        if cedula_val not in empleados_consolidados:
            empleados_consolidados[cedula_val] = {
                'CÉDULA': cedula_val,
                'HISTORIAL_ARCHIVOS': []
            }

        empleados_consolidados[cedula_val]['HISTORIAL_ARCHIVOS'].append({
            'archivo': res['archivo'],
            'seccion': res['seccion']
        })

        for k, v in fila.items():
            if k == 'CÉDULA': continue
            val_str = str(v).strip()
            if val_str and val_str != 'Sin registro':
                empleados_consolidados[cedula_val][k] = val_str

    try:
        os.makedirs(CARPETA_DATA, exist_ok=True)
        with open(JSON_DB_PATH, 'w', encoding='utf-8') as f:
            json.dump(empleados_consolidados, f, ensure_ascii=False, indent=4)
        print("-> Archivo base_empleados.json creado y sincronizado exitosamente.")
    except Exception as e:
        print(f"Error guardando JSON de empleados: {e}")

# ==========================================
# BÚSQUEDA GLOBAL INSTITUCIONAL (BLOQUE POR ARCHIVO Y HOJA)
# ==========================================
def buscar_en_todos_los_excels(cedula_buscada="", ubicaciones_seleccionadas=None):
    if not os.path.exists(CARPETA_DATA):
        return []

    bloques_resultados = []
    cedula_str = str(cedula_buscada).strip()
    ubs_sel_clean = [limpiar_texto_comparacion(u) for u in ubicaciones_seleccionadas] if ubicaciones_seleccionadas else []

    archivos = [f for f in os.listdir(CARPETA_DATA) if f.endswith('.xlsx') and not f.startswith('~$')]
    
    columnas_excluidas_global = [
        'USUARIO', 'SOLICITANTE', 'ESTADO', 'NOTA_ADMIN', 'FECHA REGISTRO', 'FECHA_REGISTRO'
    ]

    for archivo in archivos:
        ruta_archivo = os.path.join(CARPETA_DATA, archivo)
        try:
            excel_file = pd.ExcelFile(ruta_archivo)
            for hoja in excel_file.sheet_names:
                df = leer_excel_inteligente(ruta_archivo, hoja)
                if df.empty:
                    continue

                mapa_columnas = {}
                columnas_hoja = []
                for c in df.columns:
                    norm_c = normalizar_nombre_col_global(c, hoja)
                    if norm_c and norm_c not in columnas_excluidas_global:
                        mapa_columnas[c] = norm_c
                        if norm_c not in columnas_hoja:
                            columnas_hoja.append(norm_c)

                if not columnas_hoja:
                    continue

                filas_coincidentes = []

                for _, row in df.iterrows():
                    cedula_encontrada = ""
                    for c in df.columns:
                        val = row[c]
                        if pd.isna(val):
                            continue
                        val_str = str(val).strip()
                        if val_str.endswith('.0'):
                            val_str = val_str[:-2]
                        val_str = val_str.replace('.', '')
                        
                        if val_str.isdigit() and 5 <= len(val_str) <= 9:
                            cedula_encontrada = val_str
                            break

                    if not cedula_encontrada:
                        continue

                    if cedula_str and cedula_str != cedula_encontrada:
                        continue

                    fila_dict = {}
                    for c_orig, c_norm in mapa_columnas.items():
                        val = row[c_orig]
                        if pd.isna(val) or str(val).lower() == 'nan':
                            val_str = "Sin registro"
                        else:
                            val_str = str(val).strip()
                            if val_str.endswith('.0'):
                                val_str = val_str[:-2]
                            if '00:00:00' in val_str:
                                val_str = val_str.split(' ')[0]
                            if not val_str:
                                val_str = "Sin registro"
                        fila_dict[c_norm] = val_str

                    if ubs_sel_clean:
                        ub_actual = limpiar_texto_comparacion(fila_dict.get('UBICACIÓN NOMINAL', ''))
                        if not any(u_sel in ub_actual for u_sel in ubs_sel_clean):
                            continue

                    if fila_dict not in filas_coincidentes:
                        filas_coincidentes.append(fila_dict)

                if filas_coincidentes:
                    bloques_resultados.append({
                        'origen': f"Archivo: {archivo} — Sección/Hoja: {hoja}",
                        'columnas': columnas_hoja,
                        'filas': filas_coincidentes
                    })

        except Exception as e:
            print(f"Error procesando {archivo} ({hoja}): {e}")

    return bloques_resultados

def obtener_cargos_unicos():
    cargos = set()
    if os.path.exists(JSON_DB_PATH):
        try:
            with open(JSON_DB_PATH, 'r', encoding='utf-8') as f:
                db = json.load(f)
                for emp in db.values():
                    c = emp.get('CARGO', '')
                    if c and c != 'Sin registro':
                        cargos.add(c)
            return sorted(list(cargos))
        except:
            pass
    return []

def obtener_todas_las_columnas_unicas():
    columnas_unicas = set()
    campos_base = [
        'CÉDULA', 'NOMBRES Y APELLIDOS', 'CARGO', 'UBICACIÓN NOMINAL', 'AÑO DE INGRESO', 'AÑO DE EGRESO',
        'DEPENDENCIA', 'GRADO', 'ESTATUS / DISFRUTE', 'PERÍODO', 'PERIODO', 'DESDE', 'HASTA', 'DIAS',
        'OBSERVACIONES', 'TELÉFONO', 'TIPO PERSONAL', 'FECHA DE NACIMIENTO', 'SERVICIO'
    ]
    for cb in campos_base:
        columnas_unicas.add(cb)

    # Columnas administrativas prohibidas que no deben mostrarse como checkboxes
    columnas_excluidas = [
        'USUARIO', 'SOLICITANTE', 'ESTADO', 'NOTA_ADMIN', 
        'FECHA REGISTRO', 'FECHA_REGISTRO', 'FECHA DE SU SOLICITUD', 'FECHA DE RECIBIDO'
    ]

    # ESCANEO DINÁMICO EN TIEMPO REAL DE TODOS LOS EXCEL EN LA CARPETA DATA
    if os.path.exists(CARPETA_DATA):
        archivos = [f for f in os.listdir(CARPETA_DATA) if f.endswith('.xlsx') and not f.startswith('~$')]
        for archivo in archivos:
            ruta_archivo = os.path.join(CARPETA_DATA, archivo)
            try:
                excel_file = pd.ExcelFile(ruta_archivo)
                for hoja in excel_file.sheet_names:
                    df = leer_excel_inteligente(ruta_archivo, hoja)
                    if df.empty:
                        continue
                    for c in df.columns:
                        norm_c = normalizar_nombre_col_global(c, hoja)
                        if norm_c and norm_c not in columnas_excluidas:
                            columnas_unicas.add(norm_c)
            except Exception as e:
                print(f"Error leyendo columnas dinámicas de {archivo}: {e}")

    # Respaldamos también con el JSON base si existe
    if os.path.exists(JSON_DB_PATH):
        try:
            with open(JSON_DB_PATH, 'r', encoding='utf-8') as f:
                db = json.load(f)
                for emp in db.values():
                    for k in emp.keys():
                        k_upper = k.strip().upper()
                        if k_upper in ['CÉDULA', 'CEDULA', 'HISTORIAL_ARCHIVOS']:
                            continue
                        if k_upper not in columnas_excluidas:
                            columnas_unicas.add(k_upper)
        except Exception as e:
            print(f"Error leyendo columnas del JSON: {e}")
            
    columnas_filtradas = [c for c in columnas_unicas if c not in columnas_excluidas]
    return sorted(list(columnas_filtradas))

def obtener_historial_modulo(nombre_excel_modulo, cedula_buscada):
    try:
        ced_str = str(cedula_buscada).strip().replace('.0', '').replace('.', '')
        historial = []
        vistos = set()
        
        if not os.path.exists(CARPETA_DATA):
            return []

        archivos = [f for f in os.listdir(CARPETA_DATA) if f.endswith('.xlsx') and not f.startswith('~$')]
        
        if nombre_excel_modulo in archivos:
            archivos.remove(nombre_excel_modulo)
            archivos.insert(0, nombre_excel_modulo)

        modulo_lower = nombre_excel_modulo.lower()
        palabras_clave = []
        if 'vacaciones' in modulo_lower:
            palabras_clave = ['VACACION', 'VACACIONES', 'DISFRUTE']
        elif 'reposos' in modulo_lower:
            palabras_clave = ['REPOSO', 'PERMISO', 'REPOSOS', 'PERMISOS']
        elif 'jubilados' in modulo_lower:
            palabras_clave = ['JUBILADO', 'FALLECIDO', 'JUBILADOS', 'FALLECIDOS', 'EGRESO', 'PENSIONADO']

        for archivo in archivos:
            ruta_archivo = os.path.join(CARPETA_DATA, archivo)
            try:
                excel_file = pd.ExcelFile(ruta_archivo)
                for hoja in excel_file.sheet_names:
                    df = leer_excel_inteligente(ruta_archivo, hoja)
                    if df.empty:
                        continue

                    col_ced_match = None
                    for c in df.columns:
                        norm_c = normalizar_nombre_col_global(c, hoja)
                        if norm_c == 'CÉDULA':
                            col_ced_match = c
                            break

                    if not col_ced_match:
                        continue

                    hoja_str = f"{archivo} {hoja}".upper()
                    
                    for _, row in df.iterrows():
                        val_ced = row[col_ced_match]
                        if pd.isna(val_ced):
                            continue
                        
                        val_str = str(val_ced).strip()
                        if val_str.endswith('.0'):
                            val_str = val_str[:-2]
                        val_str = val_str.replace('.', '')

                        if val_str == ced_str:
                            row_dict = {}
                            fila_texto_completo = ""
                            
                            for c in df.columns:
                                norm_col = normalizar_nombre_col_global(c, hoja)
                                if not norm_col:
                                    continue
                                
                                v = row[c]
                                if pd.isna(v) or str(v).lower() in ['nan', 'nat', '']:
                                    v_str = 'Sin registro'
                                else:
                                    v_str = str(v).strip()
                                    if v_str.endswith('.0'):
                                        v_str = v_str[:-2]
                                    if '00:00:00' in v_str:
                                        v_str = v_str.split(' ')[0]
                                    if not v_str:
                                        v_str = 'Sin registro'
                                
                                row_dict[norm_col] = v_str
                                fila_texto_completo += f" {v_str.upper()}"

                            row_dict['CÉDULA'] = ced_str

                            coincide_modulo = True
                            if palabras_clave and not ('vacaciones.xlsx' in archivo or 'reposos_permisos.xlsx' in archivo or 'jubilados_fallecidos.xlsx' in archivo):
                                coincide_modulo = any(pc in fila_texto_completo or pc in hoja_str for pc in palabras_clave)

                            if not coincide_modulo:
                                continue

                            tiene_datos_modulo = any(v != 'Sin registro' for k, v in row_dict.items() if k not in [
                                'CÉDULA', 'NOMBRES Y APELLIDOS', 'CARGO', 'UBICACIÓN NOMINAL', 'AÑO DE INGRESO', 'DEPENDENCIA', 'TELÉFONO', 'GRADO'
                            ])

                            if tiene_datos_modulo:
                                key_evento = f"{row_dict.get('PERÍODO', '')}_{row_dict.get('ESTATUS / DISFRUTE', '')}_{row_dict.get('DESDE', '')}_{row_dict.get('HASTA', '')}_{row_dict.get('DIAS', '')}"
                                if key_evento not in vistos:
                                    vistos.add(key_evento)
                                    historial.append(row_dict)

            except Exception as e:
                print(f"Error escaneando archivo {archivo} ({hoja}) para historial: {e}")

        return historial[::-1]
    except Exception as e:
        print(f"Error general en obtener_historial_modulo: {e}")
        return []

# ==========================================
# RUTAS DE AUTENTICACIÓN Y PANEL PRINCIPAL
# ==========================================
@app.route('/', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '').strip()
        
        data = {}
        if os.path.exists(JSON_PATH):
            try:
                with open(JSON_PATH, 'r', encoding='utf-8') as f:
                    data = json.load(f)
            except Exception as e:
                print(f"Error leyendo el JSON en login: {e}")
        
        users_db = data.get('users', {})
        if username in users_db:
            user_info = users_db[username]
            if isinstance(user_info, dict):
                stored_pwd = user_info.get('password', '')
                user_role = user_info.get('role', 'operator')
            else:
                stored_pwd = user_info
                user_role = 'admin' if username == 'admin' else 'operator'
            
            if stored_pwd == password:
                session['user'] = username
                session['role'] = user_role
                registrar_auditoria(username, "Inicio de sesión exitoso")
                return redirect(url_for('dashboard'))
                
        flash('Credenciales incorrectas', 'danger')
    return render_template('login.html')

@app.route('/logout')
def logout():
    if 'user' in session:
        registrar_auditoria(session['user'], "Cierre de sesión")
    session.clear()
    return redirect(url_for('login'))

@app.route('/dashboard')
def dashboard():
    if 'user' not in session:
        return redirect(url_for('login'))
    config = load_json()
    modules_dict = config.get('modules', {})
    role = session.get('role')
    user = session.get('user')
    solicitudes = cargar_solicitudes()
    mis_solicitudes = [s for s in solicitudes if role == 'admin' or s['usuario'] == user]
    return render_template('index.html', modules_dict=modules_dict, role=role, mis_solicitudes=mis_solicitudes)

def calcular_dias_vacaciones_oficial(tipo_personal, grado, anio_ingreso_str, lista_periodos=None):
    try:
        anio_ingreso_limpio = str(anio_ingreso_str).strip().split('-')[0]
        anio_ing = int(anio_ingreso_limpio)
    except:
        anio_ing = datetime.now().year

    tipo_up = str(tipo_personal).strip().upper()
    grado_up = str(grado).strip().upper().replace('GRADO', '').strip()

    if not lista_periodos:
        anos_actuales = max(1, datetime.now().year - anio_ing)
        lista_periodos = [datetime.now().year]

    total_dias_acumulados = 0

    for p in lista_periodos:
        try:
            anio_periodo = int(str(p).strip())
            anos_servicio = max(1, (anio_periodo - anio_ing) + 1)
        except:
            anos_servicio = 1

        dias_este_periodo = 18

        if 'MEDICO' in tipo_up or 'MÉDICO' in tipo_up:
            if 1 <= anos_servicio <= 3: dias_este_periodo = 18
            elif 4 <= anos_servicio <= 5: dias_este_periodo = 20
            elif 6 <= anos_servicio <= 7: dias_este_periodo = 22
            elif 8 <= anos_servicio <= 9: dias_este_periodo = 24
            elif 10 <= anos_servicio <= 11: dias_este_periodo = 26
            elif 12 <= anos_servicio <= 13: dias_este_periodo = 28
            elif 14 <= anos_servicio <= 15: dias_este_periodo = 31
            elif 16 <= anos_servicio <= 17: dias_este_periodo = 34
            elif 18 <= anos_servicio <= 19: dias_este_periodo = 37
            elif 20 <= anos_servicio <= 21: dias_este_periodo = 40
            elif 22 <= anos_servicio <= 23: dias_este_periodo = 43
            else: dias_este_periodo = 46
        else:
            if anos_servicio >= 16:
                dias_este_periodo = 30
            else:
                tabla_general = {
                    1: {'EMPLEADO': 18, 'OBRERO': 20, 'CONTRATADO': 15},
                    2: {'EMPLEADO': 18, 'OBRERO': 20, 'CONTRATADO': 16},
                    3: {'EMPLEADO': 18, 'OBRERO': 20, 'CONTRATADO': 17},
                    4: {'EMPLEADO': 18, 'OBRERO': 20, 'CONTRATADO': 18},
                    5: {'EMPLEADO': 19, 'OBRERO': 20, 'CONTRATADO': 19},
                    6: {'EMPLEADO': 21, 'OBRERO': 22, 'CONTRATADO': 20},
                    7: {'EMPLEADO': 21, 'OBRERO': 22, 'CONTRATADO': 21},
                    8: {'EMPLEADO': 22, 'OBRERO': 22, 'CONTRATADO': 22},
                    9: {'EMPLEADO': 23, 'OBRERO': 23, 'CONTRATADO': 23},
                    10: {'EMPLEADO': 24, 'OBRERO': 24, 'CONTRATADO': 24},
                    11: {'EMPLEADO': 25, 'OBRERO': 25, 'CONTRATADO': 25},
                    12: {'EMPLEADO': 26, 'OBRERO': 26, 'CONTRATADO': 26},
                    13: {'EMPLEADO': 27, 'OBRERO': 27, 'CONTRATADO': 27},
                    14: {'EMPLEADO': 28, 'OBRERO': 28, 'CONTRATADO': 28},
                    15: {'EMPLEADO': 29, 'OBRERO': 29, 'CONTRATADO': 29},
                }
                tipo_key = 'EMPLEADO'
                if 'OBRERO' in tipo_up: tipo_key = 'OBRERO'
                elif 'CONTRATADO' in tipo_up: tipo_key = 'CONTRATADO'

                dias_este_periodo = tabla_general.get(anos_servicio, {}).get(tipo_key, 18)

        total_dias_acumulados += dias_este_periodo

    return total_dias_acumulados

# ==========================================
# MÓDULO 1: VACACIONES
# ==========================================
@app.route('/vacaciones', methods=['GET', 'POST'])
def vacaciones():
    if 'user' not in session: return redirect(url_for('login'))
    cedula = request.args.get('cedula', '').strip()
    empleado_data = None
    cargos = obtener_cargos_unicos()
    historial = []

    if cedula:
        res = buscar_en_todos_los_excels(cedula)
        # Búsqueda robusta en todos los bloques institucionales
        for bloque in res:
            for fila in bloque.get('filas', []):
                if str(fila.get('CÉDULA', '')).strip() == str(cedula).strip():
                    empleado_data = fila
                    break
            if empleado_data:
                break
        historial = obtener_historial_modulo('vacaciones.xlsx', cedula)

    if request.method == 'POST':
        tipo_p = request.form.get('tipo_personal', '').strip().upper()
        grado_val = request.form.get('grado', '').strip()
        anio_ing = request.form.get('anio_ingreso', '').strip()
        lista_per = request.form.getlist('periodos')

        cedula_post = request.form.get('cedula', '').strip()
        if not cedula_post.isdigit():
            flash('La cédula debe contener únicamente números.', 'danger')
            return redirect(url_for('vacaciones', cedula=cedula_post))

        if anio_ing:
            val_ing_check = anio_ing.replace('/', '').replace('-', '')
            if not val_ing_check.isdigit():
                flash('El año de ingreso tiene un formato no válido.', 'danger')
                return redirect(url_for('vacaciones', cedula=cedula_post))

        dias_calculados_backend = calcular_dias_vacaciones_oficial(tipo_p, grado_val, anio_ing, lista_per)

        nueva_sol = {
            'id': datetime.now().strftime('%Y%m%d%H%M%S'),
            'modulo': 'VACACIONES',
            'usuario': session.get('user'),
            'cedula': request.form.get('cedula', '').strip(),
            'nombres': request.form.get('nombres', '').strip().upper(),
            'tipo_personal': tipo_p,
            'cargo': request.form.get('cargo', '').strip().upper(),
            'grado': grado_val,
            'dependencia': request.form.get('dependencia', '').strip().upper(),
            'ubicacion_nominal': request.form.get('ubicacion_nominal', '').strip().upper(),
            'anio_ingreso': anio_ing,
            'periodos': ", ".join(lista_per),
            'estatus_disfrute': 'VACACION',
            'dias': str(dias_calculados_backend),
            'observaciones': request.form.get('observaciones', '').strip().upper(),
            'desde': request.form.get('desde', '').strip(),
            'hasta': request.form.get('hasta', '').strip(),
            'telefono': request.form.get('telefono', '').strip(),
            'fecha_registro': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'estado': 'PENDIENTE'
        }
        sols = cargar_solicitudes()
        sols.append(nueva_sol)
        guardar_solicitudes(sols)
        registrar_auditoria(session['user'], f"Envió solicitud de vacaciones para cédula {nueva_sol['cedula']}")
        flash('Solicitud de registro de vacaciones enviada al Administrador para su validación.', 'success')
        return redirect(url_for('vacaciones', cedula=nueva_sol['cedula']))

    solicitudes = [s for s in cargar_solicitudes() if s['modulo'] == 'VACACIONES' and (session.get('role') == 'admin' or s['usuario'] == session.get('user'))]
    return render_template('vacaciones.html', cedula_buscada=cedula, empleado_data=empleado_data, cargos_disponibles=cargos, historial=historial, solicitudes=solicitudes, role=session.get('role'))

# ==========================================
# MÓDULO 2: REPOSOS / PERMISOS
# ==========================================
@app.route('/reposos_permisos', methods=['GET', 'POST'])
def reposos_permisos():
    if 'user' not in session: return redirect(url_for('login'))
    cedula = request.args.get('cedula', '').strip()
    empleado_data = None
    cargos = obtener_cargos_unicos()
    historial = []

    if cedula:
        res = buscar_en_todos_los_excels(cedula)
        for bloque in res:
            for fila in bloque.get('filas', []):
                if str(fila.get('CÉDULA', '')).strip() == str(cedula).strip():
                    empleado_data = fila
                    break
            if empleado_data:
                break
        historial = obtener_historial_modulo('reposos_permisos.xlsx', cedula)

    if request.method == 'POST':
        cedula_post = request.form.get('cedula', '').strip()
        anio_ing = request.form.get('anio_ingreso', '').strip()

        if not cedula_post.isdigit():
            flash('La cédula debe contener únicamente números.', 'danger')
            return redirect(url_for('reposos_permisos', cedula=cedula_post))

        if anio_ing:
            val_ing_check = anio_ing.replace('/', '').replace('-', '')
            if not val_ing_check.isdigit():
                flash('El año de ingreso tiene un formato no válido.', 'danger')
                return redirect(url_for('reposos_permisos', cedula=cedula_post))

        nueva_sol = {
            'id': datetime.now().strftime('%Y%m%d%H%M%S'),
            'modulo': 'REPOSOS_PERMISOS',
            'usuario': session.get('user'),
            'cedula': request.form.get('cedula', '').strip(),
            'nombres': request.form.get('nombres', '').strip().upper(),
            'cargo': request.form.get('cargo', '').strip().upper(),
            'tipo_personal': request.form.get('tipo_personal', '').strip().upper(),
            'anio_ingreso': anio_ing,
            'ubicacion_nominal': request.form.get('ubicacion_nominal', '').strip().upper(),
            'periodos': ", ".join(request.form.getlist('periodos')),
            'estatus_disfrute': request.form.get('estatus_disfrute', '').strip().upper(),
            'dias': request.form.get('dias', '').strip(),
            'desde': request.form.get('desde', '').strip(),
            'hasta': request.form.get('hasta', '').strip(),
            'fecha_registro': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'estado': 'PENDIENTE'
        }
        sols = cargar_solicitudes()
        sols.append(nueva_sol)
        guardar_solicitudes(sols)
        registrar_auditoria(session['user'], f"Envió solicitud de reposo/permiso para cédula {nueva_sol['cedula']}")
        flash('Solicitud de reposo/permiso enviada al Administrador para su validación.', 'success')
        return redirect(url_for('reposos_permisos', cedula=nueva_sol['cedula']))

    solicitudes = [s for s in cargar_solicitudes() if s['modulo'] == 'REPOSOS_PERMISOS' and (session.get('role') == 'admin' or s['usuario'] == session.get('user'))]
    return render_template('reposos_permisos.html', cedula_buscada=cedula, empleado_data=empleado_data, cargos_disponibles=cargos, historial=historial, solicitudes=solicitudes, role=session.get('role'))

# ==========================================
# MÓDULO 3: JUBILADOS / FALLECIDOS
# ==========================================
@app.route('/jubilados_fallecidos', methods=['GET', 'POST'])
def jubilados_fallecidos():
    if 'user' not in session: return redirect(url_for('login'))
    cedula = request.args.get('cedula', '').strip()
    empleado_data = None
    historial = []

    if cedula:
        res = buscar_en_todos_los_excels(cedula)
        for bloque in res:
            for fila in bloque.get('filas', []):
                if str(fila.get('CÉDULA', '')).strip() == str(cedula).strip():
                    empleado_data = fila
                    break
            if empleado_data:
                break
        historial = obtener_historial_modulo('jubilados_fallecidos.xlsx', cedula)

    if request.method == 'POST':
        cedula_post = request.form.get('cedula', '').strip()
        anio_ing = request.form.get('anio_ingreso', '').strip()

        if not cedula_post.isdigit():
            flash('La cédula debe contener únicamente números.', 'danger')
            return redirect(url_for('jubilados_fallecidos', cedula=cedula_post))

        if anio_ing:
            val_ing_check = anio_ing.replace('/', '').replace('-', '')
            if not val_ing_check.isdigit():
                flash('El año de ingreso tiene un formato no válido.', 'danger')
                return redirect(url_for('jubilados_fallecidos', cedula=cedula_post))

        nueva_sol = {
            'id': datetime.now().strftime('%Y%m%d%H%M%S'),
            'modulo': 'JUBILADOS_FALLECIDOS',
            'usuario': session.get('user'),
            'cedula': request.form.get('cedula', '').strip(),
            'nombres': request.form.get('nombres', '').strip().upper(),
            'anio_ingreso': anio_ing,
            'anio_egreso': request.form.get('anio_egreso', '').strip(),
            'ubicacion_nominal': request.form.get('ubicacion_nominal', '').strip().upper(),
            'cargo': request.form.get('cargo', '').strip().upper(),
            'fecha_registro': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'estado': 'PENDIENTE'
        }
        sols = cargar_solicitudes()
        sols.append(nueva_sol)
        guardar_solicitudes(sols)
        registrar_auditoria(session['user'], f"Envió solicitud de jubilado/fallecido para cédula {nueva_sol['cedula']}")
        flash('Solicitud de jubilado/fallecido enviada al Administrador para su validación.', 'success')
        return redirect(url_for('jubilados_fallecidos', cedula=nueva_sol['cedula']))

    solicitudes = [s for s in cargar_solicitudes() if s['modulo'] == 'JUBILADOS_FALLECIDOS' and (session.get('role') == 'admin' or s['usuario'] == session.get('user'))]
    return render_template('jubilados_fallecidos.html', cedula_buscada=cedula, empleado_data=empleado_data, historial=historial, solicitudes=solicitudes, role=session.get('role'))

# ==========================================
# PROCESAR SOLICITUDES (ADMIN) CON NOTAS
# ==========================================
@app.route('/admin/procesar_solicitud/<sol_id>/<accion>', methods=['POST'])
@admin_required
def procesar_solicitud(sol_id, accion):
    sols = cargar_solicitudes()
    sol_encontrada = None
    for s in sols:
        if s['id'] == sol_id:
            sol_encontrada = s
            break
    
    if not sol_encontrada:
        flash('Solicitud no encontrada.', 'danger')
        return redirect(url_for('dashboard'))

    nota_admin = request.form.get('nota_admin', '').strip().upper()
    if not nota_admin:
        nota_admin = "SIN OBSERVACIONES"

    sol_encontrada['nota_admin'] = nota_admin

    if accion == 'aprobar':
        sol_encontrada['estado'] = 'APROBADO'
        modulo = sol_encontrada['modulo']
        nombre_excel = 'vacaciones.xlsx' if modulo == 'VACACIONES' else ('reposos_permisos.xlsx' if modulo == 'REPOSOS_PERMISOS' else 'jubilados_fallecidos.xlsx')
        ruta_excel = os.path.join(CARPETA_DATA, nombre_excel)
        
        os.makedirs(CARPETA_DATA, exist_ok=True)
        datos_guardar = {k: v for k, v in sol_encontrada.items() if k not in ['id', 'modulo', 'usuario', 'estado']}
        datos_excel = {k.upper().replace('_', ' '): v for k, v in datos_guardar.items()}
        
        if os.path.exists(ruta_excel):
            df_ex = pd.read_excel(ruta_excel)
            df_ex = pd.concat([df_ex, pd.DataFrame([datos_excel])], ignore_index=True)
        else:
            df_ex = pd.DataFrame([datos_excel])
        df_ex.to_excel(ruta_excel, index=False)
        
        registrar_auditoria(session['user'], f"Aprobó registro de {modulo} para cédula {sol_encontrada['cedula']} [Nota: {nota_admin}]")
        flash(f'Aprobado el registro de la cédula {sol_encontrada["cedula"]}. Observación guardada.', 'success')
    elif accion == 'rechazar':
        sol_encontrada['estado'] = 'RECHAZADO'
        registrar_auditoria(session['user'], f"Rechazó solicitud de {sol_encontrada['modulo']} para cédula {sol_encontrada['cedula']} [Motivo: {nota_admin}]")
        flash(f'Solicitud de la cédula {sol_encontrada["cedula"]} rechazada.', 'warning')

    guardar_solicitudes(sols)
    return redirect(request.referrer or url_for('dashboard'))
@app.route('/solicitar_recuperacion', methods=['POST'])
def solicitar_recuperacion():
    usuario = request.form.get('usuario_recuperacion', '').strip()
    if not usuario:
        flash('Debe ingresar un nombre de usuario válido.', 'danger')
        return redirect(url_for('login'))
    
    sols = cargar_solicitudes_recuperacion()
    nueva_solicitud = {
        'id': len(sols) + 1,
        'usuario': usuario,
        'fecha': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'estado': 'Pendiente'
    }
    sols.append(nueva_solicitud)
    guardar_solicitudes_recuperacion(sols)
    
    flash('Solicitud de recuperación de contraseña enviada al Administrador.', 'success')
    return redirect(url_for('login'))

@app.route('/atender_recuperacion/<int:sol_id>', methods=['POST'])
@admin_required
def atender_recuperacion(sol_id):
    sols = cargar_solicitudes_recuperacion()
    for s in sols:
        if int(s.get('id', 0)) == int(sol_id):
            s['estado'] = 'Atendido'
            
    guardar_solicitudes_recuperacion(sols)
    registrar_auditoria(session['user'], f"Marcó como atendida la solicitud de recuperación de contraseña ID: {sol_id}")
    flash('Solicitud de recuperación marcada como atendida.', 'success')
    return redirect(url_for('admin_panel'))
# ==========================================
# BUSCADOR GLOBAL Y EXPORTACIÓN (CON PAGINACIÓN POR BLOQUES)
# ==========================================
@app.route('/buscador_global', methods=['GET', 'POST'])
@app.route('/buscar', methods=['GET', 'POST'])
def buscador_global():
    if 'user' not in session:
        return redirect(url_for('login'))
        
    usuario_actual = session.get('user', 'Desconocido')
    todas_las_columnas = obtener_todas_las_columnas_unicas()
    cedula_consultada = request.args.get('search_cedula', '').strip() or request.args.get('cedula', '').strip()
    columnas_seleccionadas = request.args.getlist('cols')
    ubicaciones_seleccionadas = request.args.getlist('ubicaciones_esp')

    if cedula_consultada and not cedula_consultada.isdigit():
        flash('La cédula debe contener únicamente números.', 'danger')
        return redirect(url_for('buscador_global', cols=columnas_seleccionadas, ubicaciones_esp=ubicaciones_seleccionadas))
    
    try:
        page = int(request.args.get('page', 1))
        if page < 1:
            page = 1
    except ValueError:
        page = 1
    per_page = 20

    if request.method == 'POST':
        cedula_consultada = request.form.get('search_cedula', '').strip() or request.form.get('cedula', '').strip()
        columnas_seleccionadas = request.form.getlist('cols')
        ubicaciones_seleccionadas = request.form.getlist('ubicaciones_esp')
        if cedula_consultada and not cedula_consultada.isdigit():
            flash('La cédula debe contener únicamente números.', 'danger')
            return redirect(url_for('buscador_global', cols=columnas_seleccionadas, ubicaciones_esp=ubicaciones_seleccionadas))
        registrar_auditoria(usuario_actual, f"Realizó búsqueda en Buscador Global (Cédula: {cedula_consultada if cedula_consultada else 'Masiva general'})")
        return redirect(url_for('buscador_global', search_cedula=cedula_consultada, cols=columnas_seleccionadas, ubicaciones_esp=ubicaciones_seleccionadas, page=1))

    if cedula_consultada:
        registrar_auditoria(usuario_actual, f"Consultó cédula en Buscador Global: {cedula_consultada} (Pág. {page})")
    else:
        registrar_auditoria(usuario_actual, f"Accedió a consulta general en Buscador Global (Pág. {page})")

    # Obtiene la lista de bloques completos
    bloques_completos = buscar_en_todos_los_excels(cedula_consultada, ubicaciones_seleccionadas)

    # Filtrar columnas si el usuario marcó casillas específicas
    if columnas_seleccionadas:
        for b in bloques_completos:
            cols_filtradas = [c for c in b['columnas'] if c in columnas_seleccionadas]
            if cols_filtradas:
                b['columnas'] = cols_filtradas

    total_resultados = sum(len(b['filas']) for b in bloques_completos)
    total_paginas = max(1, (total_resultados + per_page - 1) // per_page)
    if page > total_paginas:
        page = total_paginas

    start_idx = (page - 1) * per_page
    end_idx = start_idx + per_page

    resultados_bloques_paginados = []
    contador_global = 0

    for bloque in bloques_completos:
        filas_bloque = bloque['filas']
        n_filas = len(filas_bloque)
        bloque_inicio = contador_global
        bloque_fin = contador_global + n_filas

        inter_inicio = max(start_idx, bloque_inicio)
        inter_fin = min(end_idx, bloque_fin)

        if inter_inicio < inter_fin:
            slice_inicio = inter_inicio - bloque_inicio
            slice_fin = inter_fin - bloque_inicio
            resultados_bloques_paginados.append({
                'origen': bloque['origen'],
                'columnas': bloque['columnas'],
                'filas': filas_bloque[slice_inicio:slice_fin],
                'total_bloque': n_filas
            })

        contador_global += n_filas

    delta = 2
    rango_paginas = []
    for p in range(1, total_paginas + 1):
        if p == 1 or p == total_paginas or (page - delta <= p <= page + delta):
            rango_paginas.append(p)
        elif rango_paginas[-1] != '...':
            rango_paginas.append('...')

    return render_template('buscador_global.html', 
                           todas_las_columnas=todas_las_columnas,
                           ubicaciones_especificas=UBICACIONES_ESPECIFICAS,
                           columnas_seleccionadas=columnas_seleccionadas,
                           ubicaciones_seleccionadas=ubicaciones_seleccionadas,
                           resultados_bloques=resultados_bloques_paginados,
                           cedula_consultada=cedula_consultada,
                           page=page,
                           total_paginas=total_paginas,
                           total_resultados=total_resultados,
                           rango_paginas=rango_paginas,
                           role=session.get('role'))

@app.route('/exportar_excel_personalizado', methods=['GET', 'POST'])
def exportar_excel_personalizado():
    if 'user' not in session:
        return redirect(url_for('login'))
        
    try:
        cedula = request.args.get('cedula', '').strip() or request.form.get('cedula', '').strip() or request.args.get('search_cedula', '').strip()
        cols = request.args.getlist('cols') or request.form.getlist('cols')
        ubicaciones_seleccionadas = request.args.getlist('ubicaciones_esp') or request.form.getlist('ubicaciones_esp')
        
        bloques_resultados = buscar_en_todos_los_excels(cedula, ubicaciones_seleccionadas)
        if cedula:
            gaveteros_raw = buscar_en_gaveteros_fisicos(cedula)
            for gav in gaveteros_raw:
                bloques_resultados.append(gav)
        
        if cols:
            for b in bloques_resultados:
                cols_filtradas = [c for c in b['columnas'] if c in cols]
                if cols_filtradas:
                    b['columnas'] = cols_filtradas

        os.makedirs('reportes_generados', exist_ok=True)
        nombre_archivo = f'Reporte_Global_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
        ruta_guardado = os.path.join('reportes_generados', nombre_archivo)
        
        with pd.ExcelWriter(ruta_guardado, engine='openpyxl') as writer:
            if not bloques_resultados:
                df_vacio = pd.DataFrame([{'Mensaje': 'No se encontraron registros'}])
                df_vacio.to_excel(writer, sheet_name='Reporte Global', index=False)
            else:
                fila_inicio = 0
                for bloque in bloques_resultados:
                    df_origen = pd.DataFrame([[f"Origen: {bloque['origen']}"]])
                    df_origen.to_excel(writer, sheet_name='Reporte Global', startrow=fila_inicio, index=False, header=False)
                    fila_inicio += 1
                    
                    columnas_bloque = bloque['columnas']
                    filas_bloque = []
                    for f in bloque['filas']:
                        filas_bloque.append({c: f.get(c, 'Sin registro') for c in columnas_bloque})
                    
                    df_bloque = pd.DataFrame(filas_bloque)
                    df_bloque.to_excel(writer, sheet_name='Reporte Global', startrow=fila_inicio, index=False)
                    fila_inicio += len(df_bloque) + 3

        registrar_auditoria(session['user'], f"Generó reporte Excel estructurado por bloques (Cédula: {cedula if cedula else 'General'})")
        return send_file(ruta_guardado, as_attachment=True, download_name=nombre_archivo)
        
    except Exception as e:
        print(f"ERROR AL EXPORTAR EXCEL: {e}")
        flash(f"Error al generar el archivo Excel: {e}", "danger")
        return redirect(url_for('buscador_global'))
    
@app.route('/exportar_pdf_personalizado', methods=['GET', 'POST'])
def exportar_pdf_personalizado():
    if 'user' not in session:
        return redirect(url_for('login'))
        
    cedula = request.values.get('cedula', '').strip()
    ubicaciones_seleccionadas = request.values.getlist('ubicaciones_esp')
    
    bloques_resultados = buscar_en_todos_los_excels(cedula, ubicaciones_seleccionadas)
    
    os.makedirs('reportes_generados', exist_ok=True)
    nombre_archivo = f'Reporte_Global_Bloques_{datetime.now().strftime("%Y%m%d_%H%M%S")}.pdf'
    ruta_guardado = os.path.join('reportes_generados', nombre_archivo)
    
    doc = SimpleDocTemplate(
        ruta_guardado,
        pagesize=landscape(letter),
        rightMargin=30, leftMargin=30, topMargin=20, bottomMargin=20
    )
    
    styles = getSampleStyleSheet()
    style_title = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontSize=10, textColor=colors.HexColor('#1b365d'), alignment=1, spaceAfter=2, leading=12)
    style_subtitle = ParagraphStyle('SubTitleStyle', parent=styles['Heading2'], fontSize=9, textColor=colors.HexColor('#2c3e50'), spaceBefore=2, spaceAfter=4, leading=11)
    style_cell = ParagraphStyle('CellStyle', parent=styles['Normal'], fontSize=7.5, textColor=colors.HexColor('#333333'), leading=9)
    style_header = ParagraphStyle('HeaderStyle', parent=styles['Normal'], fontSize=7.5, textColor=colors.white, fontName="Helvetica-Bold", alignment=1, leading=9)
    
    elements = []
    elements.append(Paragraph("HOSPITAL UNIVERSITARIO DR. MANUEL NÚÑEZ TOVAR (HUMNT)", style_title))
    elements.append(Paragraph(f"Reporte Global de Expedientes — Cédula Consultada: {cedula if cedula else 'Búsqueda Masiva General'}", style_subtitle))
    elements.append(Spacer(1, 4))
    
    if not bloques_resultados:
        elements.append(Paragraph("No se encontraron registros en los archivos del sistema para los parámetros especificados.", style_cell))
    else:
        for bloque in bloques_resultados:
            elements.append(Paragraph(f"📁 {bloque['origen']}", style_subtitle))
            elements.append(Spacer(1, 2))
            
            columnas = bloque['columnas']
            
            # Estrategia de partición: Máximo 6 columnas por bloque, manteniendo Cédula y Nombres como anclas fijas
            LIMITE_COLUMNAS_POR_PAGINA = 6
            columnas_base = ['CÉDULA', 'NOMBRES Y APELLIDOS']
            cols_dinamicas = [c for c in columnas if c not in columnas_base]
            
            sub_listas_cols = [cols_dinamicas[i:i + LIMITE_COLUMNAS_POR_PAGINA] for i in range(0, len(cols_dinamicas), LIMITE_COLUMNAS_POR_PAGINA)]
            if not sub_listas_cols:
                sub_listas_cols = [[]]
                
            for idx_sub, sub_cols in enumerate(sub_listas_cols):
                columnas_actuales = []
                for cb in columnas_base:
                    if cb in columnas and cb not in columnas_actuales:
                        columnas_actuales.append(cb)
                for sc in sub_cols:
                    if sc not in columnas_actuales:
                        columnas_actuales.append(sc)
                
                if len(sub_listas_cols) > 1:
                    elements.append(Paragraph(f"<i>(Parte {idx_sub + 1} de {len(sub_listas_cols)} de la tabla)</i>", style_cell))
                    elements.append(Spacer(1, 2))
                
                data_tabla = []
                header_row = [Paragraph(str(col), style_header) for col in columnas_actuales]
                data_tabla.append(header_row)
                
                for fila in bloque['filas']:
                    row_cells = [Paragraph(str(fila.get(col, 'Sin registro')), style_cell) for col in columnas_actuales]
                    data_tabla.append(row_cells)
                    
                if len(data_tabla) > 1:
                    ancho_disponible = 732
                    
                    # Asignación de anchos inteligentes proporcionales para evitar desbordes y textos cortados
                    col_widths = []
                    for col in columnas_actuales:
                        col_up = str(col).upper()
                        if 'UBICACIÓN' in col_up or 'UBICACION' in col_up or 'NOMINAL' in col_up:
                            col_widths.append(150)
                        elif 'NOMBRE' in col_up or 'APELLIDO' in col_up or 'CARGO' in col_up:
                            col_widths.append(110)
                        elif 'CÉDULA' in col_up or 'CEDULA' in col_up or 'TELÉFONO' in col_up or 'TELEFONO' in col_up:
                            col_widths.append(65)
                        elif 'FECHA' in col_up or 'DESDE' in col_up or 'HASTA' in col_up:
                            col_widths.append(65)
                        else:
                            col_widths.append(70)
                            
                    suma_anchos = sum(col_widths)
                    if suma_anchos > ancho_disponible:
                        factor = ancho_disponible / suma_anchos
                        col_widths = [w * factor for w in col_widths]
                    else:
                        sobrante = (ancho_disponible - suma_anchos) / len(col_widths)
                        col_widths = [w + sobrante for w in col_widths]
                    
                    t = Table(data_tabla, colWidths=col_widths, repeatRows=1)
                    t.setStyle(TableStyle([
                        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1b365d')),
                        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
                        ('TOPPADDING', (0, 0), (-1, -1), 3),
                        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#b0c4de')),
                        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#f8f9fa')])
                    ]))
                    elements.append(t)
                    elements.append(Spacer(1, 4))
            
            elements.append(Spacer(1, 6))
                
    doc.build(elements)
    registrar_auditoria(session['user'], f"Generó reporte PDF estructurado por bloques (Cédula: {cedula if cedula else 'General'})")
    return send_file(ruta_guardado, as_attachment=True, download_name=nombre_archivo)
# ==========================================
# PANEL DE ADMINISTRADOR COMPLETO
# ==========================================
@app.route('/admin/panel', methods=['GET', 'POST'])
@admin_required
def admin_panel():
    config = load_json()
    users_db = config.get('users', {})
    
    if request.method == 'POST':
        form_type = request.form.get('form_type')
        
        if form_type == 'new_user':
            nuevo_usuario = request.form.get('new_username', '').strip()
            nueva_clave = request.form.get('new_password', '').strip()
            nuevo_rol = request.form.get('new_role', 'operator')
            
            if not nuevo_usuario or not nueva_clave:
                flash('Debe ingresar un usuario y una contraseña válidos.', 'danger')
            else:
                try:
                    with open(JSON_PATH, 'r', encoding='utf-8') as f:
                        full_data = json.load(f)
                    if 'users' not in full_data: full_data['users'] = {}
                    full_data['users'][nuevo_usuario] = {"password": nueva_clave, "role": nuevo_rol}
                    with open(JSON_PATH, 'w', encoding='utf-8') as f:
                        json.dump(full_data, f, indent=4)
                    registrar_auditoria(session['user'], f"Creación de usuario: {nuevo_usuario} ({nuevo_rol})")
                    flash(f'Usuario "{nuevo_usuario}" creado correctamente.', 'success')
                except Exception as e:
                    flash(f'Error al guardar el nuevo usuario: {e}', 'danger')
                    
        elif form_type == 'edit_user':
            usuario_a_editar = request.form.get('edit_username', '').strip()
            nueva_clave = request.form.get('edit_password', '').strip()
            nuevo_rol = request.form.get('edit_role', 'operator')
            
            if not usuario_a_editar or not nueva_clave:
                flash('Debe ingresar una contraseña válida.', 'danger')
            else:
                try:
                    with open(JSON_PATH, 'r', encoding='utf-8') as f:
                        full_data = json.load(f)
                    if 'users' in full_data and usuario_a_editar in full_data['users']:
                        full_data['users'][usuario_a_editar] = {"password": nueva_clave, "role": nuevo_rol}
                        with open(JSON_PATH, 'w', encoding='utf-8') as f:
                            json.dump(full_data, f, indent=4)
                        registrar_auditoria(session['user'], f"Actualizó credenciales del usuario: {usuario_a_editar} ({nuevo_rol})")
                        flash(f'Usuario "{usuario_a_editar}" actualizado correctamente.', 'success')
                except Exception as e:
                    flash(f'Error al actualizar el usuario: {e}', 'danger')

        elif form_type == 'delete_user':
            usuario_a_eliminar = request.form.get('delete_username', '').strip()
            if usuario_a_eliminar == 'admin':
                flash('No se puede eliminar al usuario administrador principal.', 'danger')
            else:
                try:
                    with open(JSON_PATH, 'r', encoding='utf-8') as f:
                        full_data = json.load(f)
                    if 'users' in full_data and usuario_a_eliminar in full_data['users']:
                        del full_data['users'][usuario_a_eliminar]
                        with open(JSON_PATH, 'w', encoding='utf-8') as f:
                            json.dump(full_data, f, indent=4)
                        registrar_auditoria(session['user'], f"Eliminó al usuario: {usuario_a_eliminar}")
                        flash(f'Usuario "{usuario_a_eliminar}" eliminado del sistema.', 'success')
                except Exception as e:
                    flash(f'Error al eliminar el usuario: {e}', 'danger')
                    
        elif form_type == 'edit_master':
            try:
                row_idx = int(request.form.get('row_index'))
                archivo_maestro = os.path.join(CARPETA_DATA, 'master_registro_general.xlsx')
                if os.path.exists(archivo_maestro):
                    df_master = pd.read_excel(archivo_maestro)
                    if row_idx in df_master.index:
                        for col in df_master.columns:
                            val_form = request.form.get(f'col_{col}')
                            if val_form is not None:
                                df_master.at[row_idx, col] = val_form.strip().upper()
                        df_master.to_excel(archivo_maestro, index=False)
                        registrar_auditoria(session['user'], f"Modificó registro maestro en fila #{row_idx}")
                        flash('Registro maestro actualizado exitosamente.', 'success')
            except Exception as e:
                flash(f'Error al modificar el registro maestro: {e}', 'danger')
                
        elif form_type == 'delete_master':
            try:
                row_idx = int(request.form.get('row_index'))
                archivo_maestro = os.path.join(CARPETA_DATA, 'master_registro_general.xlsx')
                if os.path.exists(archivo_maestro):
                    df_master = pd.read_excel(archivo_maestro)
                    if row_idx in df_master.index:
                        cedula_afectada = df_master.loc[row_idx].get('CÉDULA', 'Desconocida')
                        df_master = df_master.drop(row_idx).reset_index(drop=True)
                        df_master.to_excel(archivo_maestro, index=False)
                        registrar_auditoria(session['user'], f"Eliminó registro maestro de cédula: {cedula_afectada}")
                        flash('Registro eliminado del archivo maestro exitosamente.', 'success')
            except Exception as e:
                flash(f'Error al eliminar el registro maestro: {e}', 'danger')
                
        return redirect(url_for('admin_panel'))
        
    log_path = os.path.join(CARPETA_DATA, 'audit.log')
    logs = []
    if os.path.exists(log_path):
        with open(log_path, 'r', encoding='utf-8') as f:
            logs = f.readlines()[-40:]
            
    search_cedula = request.args.get('search_cedula_master', '').strip()
    master_data_row = None
    row_index_encontrado = None
    master_columns = []
    
    archivo_maestro = os.path.join(CARPETA_DATA, 'master_registro_general.xlsx')
    if os.path.exists(archivo_maestro):
        try:
            df_master = pd.read_excel(archivo_maestro)
            master_columns = list(df_master.columns)
            if search_cedula:
                for idx, row in df_master.iterrows():
                    val_ced = str(row.get('CÉDULA', ''))
                    if search_cedula in val_ced:
                        row_index_encontrado = int(idx)
                        master_data_row = row.to_dict()
                        break
        except Exception as e:
            print(f"Error leyendo master excel: {e}")

        # =========================================================================
        # NUEVO: Cargamos las solicitudes de recuperación desde su archivo aislado
        # =========================================================================
        
    return render_template('admin_panel.html', 
                           users=users_db, 
                           logs=logs, 
                           solicitudes=cargar_solicitudes()[::-1],
                           master_columns=master_columns,
                           search_cedula=search_cedula,
                           master_data_row=master_data_row,
                           row_index_encontrado=row_index_encontrado,
                           role=session.get('role'))
    
@app.route('/admin/descargar_respaldo_jornada', methods=['POST'])
@admin_required
def descargar_respaldo_jornada():
    try:
        origen = CARPETA_DATA
        destino_dir = 'backup_jornada'
        os.makedirs(destino_dir, exist_ok=True)
        
        fecha_actual = datetime.now().strftime('%Y-%m-%d_%H-%M')
        nombre_zip = f'respaldo_jornada_{fecha_actual}'
        ruta_zip_sin_extension = os.path.join(destino_dir, nombre_zip)
        
        if os.path.exists(origen):
            archivo_generado = shutil.make_archive(ruta_zip_sin_extension, 'zip', origen)
            registrar_auditoria(session['user'], f"Generó respaldo de la jornada: {os.path.basename(archivo_generado)}")
            
            return send_file(
                archivo_generado, 
                as_attachment=True, 
                download_name=f'Respaldo_Datos_{fecha_actual}.zip'
            )
    except Exception as e:
        flash(f'Error al generar el respaldo de la jornada: {e}', 'danger')
    
    return redirect(url_for('admin_panel'))

if __name__ == '__main__':
    if not os.path.exists(JSON_DB_PATH):
        print("Generando base de datos inicial de empleados...")
        sincronizar_excel_a_json()
    app.run(host='0.0.0.0', port=5000, debug=True)