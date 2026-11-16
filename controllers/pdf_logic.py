import os
import pandas as pd
from flask import url_for
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from controllers.database_ops import obtener_ruta_excel, load_json, format_number

CARPETA_PDFS = 'static/pdf'

def asegurar_carpeta_pdf():
    if not os.path.exists(CARPETA_PDFS):
        os.makedirs(CARPETA_PDFS)

def generate_section_pdf(section_name):
    asegurar_carpeta_pdf()
    pdf_path = os.path.join(CARPETA_PDFS, f'reporte_{section_name.replace(" ", "_")}.pdf')
    
    config = load_json()
    ruta_excel = None
    for mod_name, mod_data in config.get('modules', {}).items():
        if section_name in mod_data.get('secciones', {}):
            ruta_excel = obtener_ruta_excel(mod_name)
            break
            
    if not ruta_excel or not os.path.exists(ruta_excel):
        ruta_excel = os.path.join('data', 'DATA1.xlsx')

    df = pd.DataFrame()
    if os.path.exists(ruta_excel):
        try:
            df = pd.read_excel(ruta_excel, sheet_name=section_name, dtype=str)
        except:
            pass

    doc = SimpleDocTemplate(pdf_path, pagesize=letter, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30, pageCompression=0)
    elementos = []
    styles = getSampleStyleSheet()
    
    titulo_estilo = ParagraphStyle(
        'TituloEstilo', parent=styles['Heading1'],
        fontSize=14, textColor=colors.HexColor('#1b365d'), spaceAfter=15, alignment=1
    )
    
    elementos.append(Paragraph(f"REPORTE OFICIAL - SECCIÓN: {section_name}", titulo_estilo))
    elementos.append(Spacer(1, 10))

    data_tabla = [["N°", "CÉDULA REGISTRADA"]]
    if not df.empty:
        col_objetivo = df.columns[0]
        registros = [str(x) for x in df[col_objetivo].dropna().tolist() if str(x).replace('.0','').isdigit()]
        for idx, reg in enumerate(registros, 1):
            data_tabla.append([str(idx), format_number(reg)])

    tabla = Table(data_tabla, colWidths=[50, 400], repeatRows=1)
    tabla.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1b365d')),
        ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,0), 10),
        ('BOTTOMPADDING', (0,0), (-1,0), 6),
        ('BACKGROUND', (0,1), (-1,-1), colors.HexColor('#f8f9fa')),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#dee2e6')),
        ('FONTNAME', (0,1), (-1,-1), 'Helvetica'),
        ('FONTSIZE', (0,1), (-1,-1), 9),
    ]))
    
    elementos.append(tabla)
    doc.build(elementos)
    return pdf_path

def generate_full_module_pdf(module_name):
    asegurar_carpeta_pdf()
    pdf_path = os.path.join(CARPETA_PDFS, f'reporte_modulo_{module_name.replace(" ", "_")}.pdf')
    
    config = load_json()
    modules_dict = config.get('modules', {})
    
    doc = SimpleDocTemplate(pdf_path, pagesize=letter, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30, pageCompression=0)
    elementos = []
    styles = getSampleStyleSheet()
    
    titulo_estilo = ParagraphStyle(
        'TituloEstilo', parent=styles['Heading1'],
        fontSize=14, textColor=colors.HexColor('#1b365d'), spaceAfter=15, alignment=1
    )
    
    elementos.append(Paragraph(f"REPORTE GENERAL - MÓDULO: {module_name}", titulo_estilo))
    elementos.append(Spacer(1, 10))

    if module_name in modules_dict:
        ruta_excel = obtener_ruta_excel(module_name)
        secciones = modules_dict[module_name].get('secciones', {})
        
        if ruta_excel and os.path.exists(ruta_excel):
            for sec_name in secciones.keys():
                elementos.append(Paragraph(f"<b>Sección: {sec_name}</b>", styles['Heading2']))
                elementos.append(Spacer(1, 5))
                
                try:
                    df = pd.read_excel(ruta_excel, sheet_name=sec_name, dtype=str)
                    data_tabla = [["N°", "CÉDULA"]]
                    if not df.empty:
                        col_obj = df.columns[0]
                        regs = [str(x) for x in df[col_obj].dropna().tolist() if str(x).replace('.0','').isdigit()]
                        for i, r in enumerate(regs, 1):
                            data_tabla.append([str(i), format_number(r)])
                    
                    t = Table(data_tabla, colWidths=[40, 350], repeatRows=1)
                    t.setStyle(TableStyle([
                        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#2b4c7e')),
                        ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
                        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                        ('FONTSIZE', (0,0), (-1,0), 9),
                        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#dee2e6')),
                        ('FONTSIZE', (0,1), (-1,-1), 8),
                    ]))
                    elementos.append(t)
                    elementos.append(Spacer(1, 15))
                except Exception as e:
                    elementos.append(Paragraph(f"Error cargando sección: {e}", styles['Normal']))

    doc.build(elementos)
    return pdf_path

def generate_global_search_pdf_personalizado(filas_datos, columnas_seleccionadas, pdf_path):
    styles = getSampleStyleSheet()
    
    style_header = ParagraphStyle(
        'HeaderStyle', parent=styles['Normal'],
        fontName='Helvetica-Bold', fontSize=8, textColor=colors.whitesmoke, alignment=1
    )
    
    style_body = ParagraphStyle(
        'BodyStyle', parent=styles['Normal'],
        fontName='Helvetica', fontSize=8, textColor=colors.HexColor('#212529'), alignment=0
    )

    if not columnas_seleccionadas and filas_datos:
        columnas_seleccionadas = list(filas_datos[0].keys())

    elements = []
    
    ANCHO_UTIL = 732  # Ancho útil en hoja horizontal Letter
    
    # Si es una búsqueda de UNA SOLA CÉDULA, agrupamos en 3 bloques de hasta 10 columnas para que entren en 1 sola hoja.
    # Si son varios registros (búsqueda masiva), mantenemos bloques de 8 columnas por página.
    es_cedula_individual = (len(filas_datos) == 1)
    MAX_COLS_POR_BLOQUE = 10 if es_cedula_individual else 8

    bloques_columnas = [
        columnas_seleccionadas[i:i + MAX_COLS_POR_BLOQUE] 
        for i in range(0, len(columnas_seleccionadas), MAX_COLS_POR_BLOQUE)
    ]

    for idx_bloque, cols_bloque in enumerate(bloques_columnas):
        # Si NO es cédula individual y ya pasó el primer bloque, saltamos de página
        if idx_bloque > 0:
            if not es_cedula_individual:
                elements.append(PageBreak())
            else:
                elements.append(Spacer(1, 12)) # Espacio elegante entre bloques en la misma página

        # Construir encabezados del bloque
        headers = [Paragraph(str(col), style_header) for col in cols_bloque]
        data_rows = [headers]

        # Construir filas de datos para este bloque
        for fila in filas_datos:
            row = []
            for col in cols_bloque:
                valor = fila.get(col, 'Sin registro')
                if valor is None or str(valor).strip() == '' or str(valor).lower() == 'nan':
                    valor = 'Sin registro'
                row.append(Paragraph(str(valor), style_body))
            data_rows.append(row)

        # Asignar anchos proporcionales
        ancho_col = max(65, ANCHO_UTIL // len(cols_bloque))
        col_widths = [ancho_col] * len(cols_bloque)

        t = Table(data_rows, colWidths=col_widths, repeatRows=1)
        t.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1b365d')),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
            ('TOPPADDING', (0, 0), (-1, -1), 4),
            ('LEFTPADDING', (0, 0), (-1, -1), 4),
            ('RIGHTPADDING', (0, 0), (-1, -1), 4),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#cbd5e1')),
        ]))

        elements.append(t)

    # Generar documento final en orientación horizontal (landscape)
    doc = SimpleDocTemplate(
        pdf_path, 
        pagesize=landscape(letter),
        rightMargin=30, leftMargin=30, 
        topMargin=30, bottomMargin=30
    )
    doc.build(elements)
    
    return pdf_path