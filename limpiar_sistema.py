import os
import pandas as pd

CARPETA_DATA = 'data'

archivos_modulos = ['vacaciones.xlsx', 'reposos_permisos.xlsx', 'jubilados_fallecidos.xlsx']

for archivo in archivos_modulos:
    ruta = os.path.join(CARPETA_DATA, archivo)
    if os.path.exists(ruta):
        try:
            # Leer el excel para mantener solo las cabeceras originales
            df = pd.read_excel(ruta)
            df_vacio = pd.DataFrame(columns=df.columns)
            df_vacio.to_excel(ruta, index=False)
            print(f"Limpiado con éxito: {archivo}")
        except Exception as e:
            print(f"No se pudo limpiar {archivo}: {e}")

# Limpiar solicitudes pendientes
ruta_solicitudes = os.path.join(CARPETA_DATA, 'solicitudes.json')
if os.path.exists(ruta_solicitudes):
    with open(ruta_solicitudes, 'w', encoding='utf-8') as f:
        f.write('[]')
    print("Bandeja de solicitudes reiniciada.")

print("¡Sistema listo para producción!")