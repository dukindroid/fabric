# -*- coding: utf-8 -*-
import re
from pathlib import Path
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY

BASE_DIR = Path(__file__).parent
MD_FILE = BASE_DIR / 'CV-Zapata-Vendedor-Refacciones.md'
DOCX_FILE = BASE_DIR / 'CV-Zapata-Vendedor-Refacciones.docx'
PDF_FILE = BASE_DIR / 'CV-Zapata-Vendedor-Refacciones.pdf'

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
    story.append(Paragraph('VENDEDOR DE REFACCIONES | CARGA PESADA, FLOTILLA DIÉSEL Y CAMPO', title_style))
    story.append(Paragraph('Villa de Álvarez, Colima &nbsp;|&nbsp; +52 (333) 963 5400 &nbsp;|&nbsp; dukintosh@gmail.com &nbsp;|&nbsp; LinkedIn: Dukindroid', contact_style))
    story.append(HRFlowable(width='100%', thickness=1.2, color=LINE_COLOR, spaceBefore=1, spaceAfter=4))

    # Perfil
    story.append(Paragraph('PERFIL PROFESIONAL', section_style))
    story.append(Paragraph(
        'Profesional con sólida experiencia operativa en campo en mantenimiento, supervisión y seguimiento de servicios para flotilla diésel y maquinaria pesada, combinada con habilidades comerciales en atención a clientes, prospección y ventas técnicas. Amplio conocimiento práctico de motores diésel, refacciones mecánicas, sistemas de filtración, lubricantes y componentes de desgaste en entornos de trabajo pesado y cantera. Experto en control de inventarios de partes, cotización ágil y gestión de pedidos mediante catálogos técnicos y sistemas ERP. Orientado a la resolución de necesidades mecánicas directas de clientes y talleres en campo y mostrador, con visión analítica, actitud de servicio y licencia de conducir vigente.',
        body_style
    ))
    story.append(HRFlowable(width='100%', thickness=0.5, color=colors.HexColor('#D0D7DE'), spaceBefore=2, spaceAfter=3))

    # Habilidades
    story.append(Paragraph('HABILIDADES TÉCNICAS Y COMERCIALES', section_style))
    skills = [
        ('<b>Refacciones y Mecánica Diésel:</b> Identificación, cotización y selección de refacciones para tractocamiones y flotilla pesada diésel (motores diésel, sistemas de lubricación, filtración, frenos, suspensión y bandas); diagnóstico básico y asesoría técnica directa con mecánicos, transportistas y jefes de taller en campo.'),
        ('<b>Ventas y Atención al Cliente:</b> Prospección de clientes y flotillas en campo, cotización ágil de partes de alta rotación, seguimiento integral de pedidos (surtido, entrega y cobranza), atención en mostrador y negociación comercial.'),
        ('<b>Logística y Almacén:</b> Control de inventarios físicos/teóricos de refacciones, recepción, acomodo, surtido ágil de piezas y operación de montacargas en patio y pasillos reducidos.'),
        ('<b>Sistemas y Catálogos:</b> Manejo de catálogos electrónicos de partes, sistemas ERP (Odoo/SAP) y hojas de cálculo para control de ventas y pedidos.'),
        ('<b>Idiomas:</b> Inglés C2 (Certificación EFSET / Bilingüe) — lectura fluida e interpretación de manuales técnicos, diagramas y números de parte en inglés.')
    ]
    for s in skills:
        story.append(Paragraph(f'&bull;&nbsp; {s}', bullet_style))
    story.append(HRFlowable(width='100%', thickness=0.5, color=colors.HexColor('#D0D7DE'), spaceBefore=2, spaceAfter=3))

    # Experiencia
    story.append(Paragraph('EXPERIENCIA LABORAL', section_style))
    
    # Job 1
    story.append(Paragraph('Grupo Minero La Fundición S.A. de C.V. | Asistente de Operaciones, Flotilla Diésel y Almacén', job_header_style))
    story.append(Paragraph('Abril 2024 – Febrero 2026', job_date_style))
    job1_bullets = [
        'Supervisión y planeación de mantenimiento preventivo y correctivo para flotilla pesada diésel (tractocamiones, excavadoras, payloaders) y equipos en cantera.',
        'Gestión y control integral del almacén de refacciones y materiales, identificando piezas requeridas y asegurando disponibilidad para minimizar tiempos muertos.',
        'Apoyo directo en campo en servicios mecánicos: cambio de filtros, fluidos, engrase, bandas, soldadura y apoyo en motores diésel.',
        'Enlace directo con talleres mecánicos, operadores y proveedores de refacciones, cotizando componentes y validando especificaciones técnicas.',
        'Registro y control de inventarios y activos mediante sistema ERP (Odoo).'
    ]
    for b in job1_bullets:
        story.append(Paragraph(f'&bull;&nbsp; {b}', bullet_style))
    
    # Job 2
    story.append(Paragraph('Surtidora Ferretera de Colima | Asistente de Almacén y Montacarguista', job_header_style))
    story.append(Paragraph('Febrero 2026 – Marzo 2026', job_date_style))
    job2_bullets = [
        'Operación de montacargas en pasillos reducidos para carga, descarga y acomodo de mercancía pesada y ferretera.',
        'Control de inventarios físicos, recepción de mercancía y surtido de pedidos de mostrador y entrega.'
    ]
    for b in job2_bullets:
        story.append(Paragraph(f'&bull;&nbsp; {b}', bullet_style))

    # Job 3
    story.append(Paragraph('Comercialización y Asesoría Comercial | Vendedor y Atención al Cliente', job_header_style))
    story.append(Paragraph('Experiencia comercial previa seleccionada', job_date_style))
    job3_bullets = [
        'Ventas consultivas y atención directa a clientes, prospección comercial, cotización de productos de alto valor, negociación y cierre de ventas.',
        'Manejo de caja, control de cobranza y atención al cliente en estación de servicio (despacho, fluidos y manejo responsable de valores).'
    ]
    for b in job3_bullets:
        story.append(Paragraph(f'&bull;&nbsp; {b}', bullet_style))

    # Job 4
    story.append(Paragraph('Corporativo ICEP / Servicios Multioc | Soporte Técnico y Control Administrativo', job_header_style))
    story.append(Paragraph('Julio 2016 – Diciembre 2017', job_date_style))
    job4_bullets = [
        'Soporte técnico en sitio y gestión de rutas de mantenimiento preventivo.',
        'Elaboración de reportes administrativos y consultas de datos para análisis de existencias y control de operaciones.'
    ]
    for b in job4_bullets:
        story.append(Paragraph(f'&bull;&nbsp; {b}', bullet_style))

    story.append(HRFlowable(width='100%', thickness=0.5, color=colors.HexColor('#D0D7DE'), spaceBefore=2, spaceAfter=3))

    # Formación e Información adicional
    story.append(Paragraph('FORMACIÓN ACADÉMICA E INFORMACIÓN ADICIONAL', section_style))
    
    info_p1 = Paragraph(
        '<b>Licenciatura en Administración</b> | ICEP Colima | En curso<br/>'
        '<b>Ingeniería en Sistemas de Computación</b> | ICEP Colima | Titulado (2016)<br/>'
        '<b>Formación complementaria:</b> ERP (Odoo), gestión de inventarios y catálogos de partes.',
        body_style
    )
    info_p2 = Paragraph(
        '<b>Licencia de conducir:</b> Sí (vigente — trabajo de campo y visitas).<br/>'
        '<b>Idiomas:</b> Inglés C2 Proficient (EFSET / Bilingüe).<br/>'
        '<b>Disponibilidad:</b> Inmediata y de tiempo completo (campo y mostrador).',
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
