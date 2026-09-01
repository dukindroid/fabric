# -*- coding: utf-8 -*-
from pathlib import Path

from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Table, TableStyle, HRFlowable
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY

BASE_DIR = Path(__file__).parent
PDF_FILE = BASE_DIR / 'CV-DLP-Auxiliar-Almacen.pdf'

DARK_COLOR = colors.HexColor('#1F3864')
TEXT_COLOR = colors.HexColor('#222222')
MUTED_COLOR = colors.HexColor('#555555')
LINE_COLOR = colors.HexColor('#1F3864')

def build_pdf():
    margin = 32
    doc = SimpleDocTemplate(
        str(PDF_FILE),
        pagesize=letter,
        leftMargin=margin,
        rightMargin=margin,
        topMargin=28,
        bottomMargin=28
    )
    
    styles = getSampleStyleSheet()
    
    name_style = ParagraphStyle(
        'CVName', parent=styles['Normal'],
        fontName='Helvetica-Bold', fontSize=17, leading=19,
        alignment=TA_CENTER, textColor=DARK_COLOR, spaceAfter=3
    )
    title_style = ParagraphStyle(
        'CVTitle', parent=styles['Normal'],
        fontName='Helvetica-Bold', fontSize=9.5, leading=11.5,
        alignment=TA_CENTER, textColor=colors.HexColor('#333333'), spaceAfter=3
    )
    contact_style = ParagraphStyle(
        'CVContact', parent=styles['Normal'],
        fontName='Helvetica', fontSize=8.5, leading=11,
        alignment=TA_CENTER, textColor=MUTED_COLOR, spaceAfter=6
    )
    section_style = ParagraphStyle(
        'CVSection', parent=styles['Normal'],
        fontName='Helvetica-Bold', fontSize=10.5, leading=12.5,
        textColor=DARK_COLOR, spaceBefore=5, spaceAfter=2, keepWithNext=True
    )
    body_style = ParagraphStyle(
        'CVBody', parent=styles['Normal'],
        fontName='Helvetica', fontSize=8.5, leading=11,
        textColor=TEXT_COLOR, alignment=TA_JUSTIFY, spaceAfter=3
    )
    job_header_style = ParagraphStyle(
        'CVJobHeader', parent=styles['Normal'],
        fontName='Helvetica-Bold', fontSize=9, leading=11,
        textColor=TEXT_COLOR, spaceBefore=3, spaceAfter=1, keepWithNext=True
    )
    job_date_style = ParagraphStyle(
        'CVJobDate', parent=styles['Normal'],
        fontName='Helvetica-Bold', fontSize=8, leading=10,
        textColor=colors.HexColor('#4A607A'), spaceAfter=1.5, keepWithNext=True
    )
    bullet_style = ParagraphStyle(
        'CVBullet', parent=styles['Normal'],
        fontName='Helvetica', fontSize=8.2, leading=10.4,
        textColor=TEXT_COLOR, leftIndent=10, firstLineIndent=-10, spaceAfter=1.5
    )

    story = []

    # Header
    story.append(Paragraph('JAVIER GUTIÉRREZ HERRERA', name_style))
    story.append(Paragraph('AUXILIAR DE ALMACÉN Y CONTROL DE INVENTARIOS', title_style))
    story.append(Paragraph('Villa de Álvarez, Colima &nbsp;|&nbsp; +52 (333) 963 5400 &nbsp;|&nbsp; dukintosh@gmail.com &nbsp;|&nbsp; LinkedIn: Dukindroid', contact_style))
    story.append(HRFlowable(width='100%', thickness=1.2, color=LINE_COLOR, spaceBefore=1, spaceAfter=4))

    # Perfil
    story.append(Paragraph('PERFIL PROFESIONAL', section_style))
    story.append(Paragraph(
        'Profesional con sólida experiencia operativa en la gestión integral de almacenes, control de inventarios y logística. Experiencia comprobada en recepción, carga, descarga y acomodo de mercancía, garantizando la correcta rotación de productos (PEPS) y la oportuna detección de mermas. Orientado a la eficiencia en el registro de entradas y salidas, así como en la atención cordial y oportuna a clientes internos y proveedores. Acostumbrado al trabajo físico y de alto dinamismo, con actitud de servicio, proactividad y licencia de chofer vigente, ideal para apoyar en tareas de distribución.',
        body_style
    ))
    story.append(HRFlowable(width='100%', thickness=0.5, color=colors.HexColor('#D0D7DE'), spaceBefore=2, spaceAfter=3))

    # Habilidades
    story.append(Paragraph('HABILIDADES Y COMPETENCIAS', section_style))
    skills = [
        ('<b>Gestión de Almacén e Inventarios:</b> Recepción, acomodo, carga y descarga eficiente de mercancía. Control riguroso de entradas y salidas para asegurar la exactitud del inventario físico y teórico.'),
        ('<b>Control de Calidad y Rotación:</b> Manejo de mercancía aplicando principios de rotación de producto (PEPS) y cuidado de empaques. Inspección para la detección y reporte oportuno de mermas o daños.'),
        ('<b>Logística y Operación:</b> Operación de equipos de carga y montacargas en pasillos reducidos. Manejo de vehículos de carga (licencia de chofer vigente) y surtido ágil de pedidos.'),
        ('<b>Atención a Proveedores y Clientes:</b> Trato directo con proveedores para recepción de mercancía y cotejo de facturas. Atención cordial a clientes internos y externos garantizando el mejor servicio.'),
        ('<b>Sistemas y Control Administrativo:</b> Manejo de sistemas ERP (Odoo/SAP) y hojas de cálculo para el registro y control preciso de movimientos de almacén.')
    ]
    for s in skills:
        story.append(Paragraph(f'&bull;&nbsp; {s}', bullet_style))
    story.append(HRFlowable(width='100%', thickness=0.5, color=colors.HexColor('#D0D7DE'), spaceBefore=2, spaceAfter=3))

    # Experiencia
    story.append(Paragraph('EXPERIENCIA LABORAL', section_style))
    
    # Job 1
    story.append(Paragraph('Grupo Minero La Fundición S.A. de C.V. | Asistente de Operaciones y Almacén', job_header_style))
    story.append(Paragraph('Abril 2024 – Febrero 2026', job_date_style))
    job1_bullets = [
        'Gestión y control integral del almacén de insumos, materiales y refacciones, asegurando recepción, acomodo y despacho eficiente.',
        'Ejecución de entradas y salidas en sistema ERP (Odoo), garantizando la exactitud del inventario diario.',
        'Control de rotación de insumos y consumibles, y detección de mermas para optimizar el uso de los recursos.',
        'Atención directa a operadores (cliente interno) y proveedores, validando especificaciones y surtiendo requerimientos operativos sin demoras.'
    ]
    for b in job1_bullets:
        story.append(Paragraph(f'&bull;&nbsp; {b}', bullet_style))
    
    # Job 2
    story.append(Paragraph('Surtidora Ferretera de Colima | Asistente de Almacén y Montacarguista', job_header_style))
    story.append(Paragraph('Febrero 2026 – Marzo 2026', job_date_style))
    job2_bullets = [
        'Carga, descarga y acomodo de mercancía pesada y diversa, optimizando el aprovechamiento del espacio en el almacén.',
        'Recepción de proveedores, cotejo de facturas con mercancía física y revisión de calidad para evitar mermas.',
        'Operación de montacargas en pasillos reducidos y surtido eficiente de pedidos para clientes de mostrador y rutas de entrega.'
    ]
    for b in job2_bullets:
        story.append(Paragraph(f'&bull;&nbsp; {b}', bullet_style))

    # Job 3
    story.append(Paragraph('Comercialización y Asesoría Comercial | Atención al Cliente y Logística', job_header_style))
    story.append(Paragraph('Experiencia comercial previa seleccionada', job_date_style))
    job3_bullets = [
        'Recepción y acomodo de productos, asegurando el orden y limpieza de las áreas de exhibición y resguardo.',
        'Atención directa al cliente, surtido de pedidos, manejo responsable de valores y cobranza.'
    ]
    for b in job3_bullets:
        story.append(Paragraph(f'&bull;&nbsp; {b}', bullet_style))

    # Job 4
    story.append(Paragraph('Corporativo ICEP / Servicios Multioc | Soporte Técnico y Control Administrativo', job_header_style))
    story.append(Paragraph('Julio 2016 – Diciembre 2017', job_date_style))
    job4_bullets = [
        'Elaboración de reportes administrativos y consultas de datos para análisis de existencias y control de operaciones.',
        'Soporte técnico en sitio y gestión de rutas logísticas de mantenimiento preventivo.'
    ]
    for b in job4_bullets:
        story.append(Paragraph(f'&bull;&nbsp; {b}', bullet_style))

    story.append(HRFlowable(width='100%', thickness=0.5, color=colors.HexColor('#D0D7DE'), spaceBefore=2, spaceAfter=3))

    # Formación e Información adicional
    story.append(Paragraph('FORMACIÓN ACADÉMICA E INFORMACIÓN ADICIONAL', section_style))
    
    info_p1 = Paragraph(
        '<b>Licenciatura en Administración</b> | ICEP Colima | En curso<br/>'
        '<b>Ingeniería en Sistemas de Computación</b> | ICEP Colima | Titulado (2016)<br/>'
        '<b>Formación complementaria:</b> ERP (Odoo), control de inventarios, gestión de almacén.',
        body_style
    )
    info_p2 = Paragraph(
        '<b>Licencia de conducir:</b> Sí (Vigente, tipo chofer — disponibilidad rutas).<br/>'
        '<b>Disponibilidad:</b> Horario de turno completo.<br/>'
        '<b>Idiomas:</b> Inglés C2 Proficient (Certificación EFSET).',
        body_style
    )

    col_table = Table([[info_p1, info_p2]], colWidths=['54%', '46%'])
    col_table.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('LEFTPADDING', (0,0), (-1,-1), 0),
        ('RIGHTPADDING', (0,0), (-1,-1), 0),
        ('TOPPADDING', (0,0), (-1,-1), 0),
        ('BOTTOMPADDING', (0,0), (-1,-1), 0),
    ]))
    story.append(col_table)

    doc.build(story)
    print('Generated PDF:', PDF_FILE)

if __name__ == '__main__':
    build_pdf()

