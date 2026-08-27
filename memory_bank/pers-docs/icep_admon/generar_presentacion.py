"""
Generador de presentación: Crujisanas — Expansión Internacional
Requiere: pip install python-pptx
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from itertools import count

# ----------------------------------------------------------------------
# PALETA DE COLORES (estética premium/artesanal, tal como la marca)
# ----------------------------------------------------------------------
NEGRO   = RGBColor(0x18, 0x18, 0x18)
GRIS    = RGBColor(0x2E, 0x2E, 0x2E)
CREMA   = RGBColor(0xF3, 0xEE, 0xE3)
NARANJA = RGBColor(0xD9, 0x7A, 0x3F)   # camote
VERDE   = RGBColor(0x7A, 0x9A, 0x5B)   # jícama
MORADO  = RGBColor(0x7A, 0x3B, 0x5D)   # betabel

prs = Presentation()
prs.slide_width  = Inches(13.333)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]
SW, SH = prs.slide_width, prs.slide_height
pager = count(1)

# ----------------------------------------------------------------------
# FUNCIONES AUXILIARES
# ----------------------------------------------------------------------
def add_slide():
    return prs.slides.add_slide(BLANK)

def set_bg(slide, color):
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = color

def rect(slide, x, y, w, h, color):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    shp.fill.solid()
    shp.fill.fore_color.rgb = color
    shp.line.fill.background()
    shp.shadow.inherit = False
    return shp

def triangle(slide, x, y, w, h, color, rotation=0):
    shp = slide.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE, x, y, w, h)
    shp.fill.solid()
    shp.fill.fore_color.rgb = color
    shp.line.fill.background()
    shp.rotation = rotation
    shp.shadow.inherit = False
    return shp

def txt(slide, x, y, w, h, text, size=16, color=CREMA, bold=False,
        align=PP_ALIGN.LEFT, italic=False, font="Calibri"):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.alignment = align
    r = p.runs[0]
    r.font.size = Pt(size)
    r.font.color.rgb = color
    r.font.bold = bold
    r.font.italic = italic
    r.font.name = font
    return tb

def bullets(slide, x, y, w, h, items, size=15, color=CREMA):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(10)
        p.line_spacing = 1.15
        r = p.add_run()
        r.text = f"●  {item}"
        r.font.size = Pt(size)
        r.font.color.rgb = color
        r.font.name = "Calibri"
    return tb

def bullets_lead(slide, x, y, w, h, items, size=15, lead_color=NARANJA, color=CREMA):
    """items: lista de tuplas (encabezado_en_negritas, resto_del_texto)"""
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    for i, (lead, rest) in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(12)
        p.line_spacing = 1.15
        r1 = p.add_run()
        r1.text = f"●  {lead} — "
        r1.font.bold = True
        r1.font.size = Pt(size)
        r1.font.color.rgb = lead_color
        r1.font.name = "Calibri"
        r2 = p.add_run()
        r2.text = rest
        r2.font.size = Pt(size)
        r2.font.color.rgb = color
        r2.font.name = "Calibri"
    return tb

def header(slide, title, section_no=None, accent=NARANJA):
    rect(slide, Inches(0), Inches(0), Inches(0.15), SH, accent)
    if section_no:
        txt(slide, Inches(11.5), Inches(0.12), Inches(1.6), Inches(1.0),
            section_no, size=48, color=GRIS, bold=True, align=PP_ALIGN.RIGHT)
    txt(slide, Inches(0.6), Inches(0.35), Inches(10.5), Inches(0.9),
        title, size=30, color=CREMA, bold=True)
    rect(slide, Inches(0.65), Inches(1.05), Inches(1.1), Pt(4), accent)

def footer(slide):
    n = next(pager)
    txt(slide, Inches(0.4), Inches(7.05), Inches(7), Inches(0.35),
        "CRUJISANAS  |  Mercadotecnia Internacional", size=10, color=GRIS)
    txt(slide, Inches(12.5), Inches(7.05), Inches(0.5), Inches(0.35),
        str(n), size=10, color=GRIS, align=PP_ALIGN.RIGHT)

def notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text

def stat_box(slide, x, y, w, h, number, label, accent):
    rect(slide, x, y, w, h, GRIS)
    rect(slide, x, y, w, Pt(4), accent)
    txt(slide, x, y + Inches(0.2), w, Inches(0.8), number, size=26,
        color=accent, bold=True, align=PP_ALIGN.CENTER)
    txt(slide, x + Inches(0.15), y + Inches(1.0), w - Inches(0.3), h - Inches(1.1),
        label, size=12, color=CREMA, align=PP_ALIGN.CENTER)

def quad_box(slide, x, y, w, h, title_, desc, accent):
    rect(slide, x, y, w, h, GRIS)
    rect(slide, x, y, Inches(0.07), h, accent)
    txt(slide, x + Inches(0.25), y + Inches(0.12), w - Inches(0.4), Inches(0.4),
        title_, size=16, color=accent, bold=True)
    txt(slide, x + Inches(0.25), y + Inches(0.55), w - Inches(0.45), h - Inches(0.7),
        desc, size=12.5, color=CREMA)

# ========================================================================
# SLIDE 1 — PORTADA
# ========================================================================
s = add_slide(); set_bg(s, NEGRO)
triangle(s, Inches(10.3), Inches(5.6), Inches(2.2), Inches(2.2), VERDE, rotation=15)
triangle(s, Inches(11.3), Inches(5.0), Inches(1.8), Inches(1.8), NARANJA, rotation=-10)
triangle(s, Inches(10.9), Inches(6.2), Inches(1.6), Inches(1.6), MORADO, rotation=35)
rect(s, Inches(0), Inches(0), Inches(0.15), SH, NARANJA)
txt(s, Inches(0.7), Inches(2.3), Inches(9), Inches(0.5),
    "PROYECTO FINAL · MERCADOTECNIA INTERNACIONAL", size=14, color=VERDE, bold=True)
txt(s, Inches(0.65), Inches(2.7), Inches(10), Inches(1.4),
    "CRUJISANAS", size=64, color=CREMA, bold=True)
txt(s, Inches(0.7), Inches(3.9), Inches(9.5), Inches(0.9),
    "De botana artesanal mexicana a marca con potencial\nde expansión internacional",
    size=20, color=CREMA)
txt(s, Inches(0.7), Inches(6.6), Inches(6), Inches(0.5),
    "Caso de estudio · 2026", size=13, color=GRIS, italic=True)
notes(s, "Buenos días/tardes. Mi presentación analiza el potencial de expansión "
         "internacional de Crujisanas, una marca mexicana de snacks saludables "
         "hechos con vegetales de raíz. A lo largo de los próximos minutos les "
         "mostraré por qué este producto tiene condiciones para exportarse, "
         "hacia qué mercado, y bajo qué estrategia, considerando el complejo "
         "panorama económico de 2026.")

# ========================================================================
# SLIDE 2 — AGENDA
# ========================================================================
s = add_slide(); set_bg(s, NEGRO)
header(s, "Contenido de la presentación", accent=VERDE)
items = [
    "01  Contexto macroeconómico 2026",
    "02  La empresa: Crujisanas",
    "03  Justificación de la elección",
    "04  Mercado seleccionado: Guatemala",
    "05  Análisis del entorno",
    "06  Propuesta de internacionalización",
    "07  Retos y conclusiones",
]
bullets(s, Inches(0.8), Inches(1.8), Inches(9), Inches(4.5), items, size=20)
footer(s)
notes(s, "Así está organizada la presentación: partiremos del contexto económico "
         "de 2026, presentaremos la empresa, justificaremos por qué es un buen "
         "candidato de exportación, elegiremos un mercado, analizaremos su "
         "entorno, propondremos una estrategia de entrada y cerraremos con "
         "retos y conclusiones.")

# ========================================================================
# SLIDE 3 — CONTEXTO 2026
# ========================================================================
s = add_slide(); set_bg(s, NEGRO)
header(s, "Un 2026 con más cautela... y oportunidades", "01", accent=NARANJA)
txt(s, Inches(0.65), Inches(1.5), Inches(11.5), Inches(0.8),
    "El entorno global exige elegir mercados cercanos, estables y con menor "
    "exposición a riesgo cambiario y logístico.", size=15, color=CREMA)
stat_box(s, Inches(0.7),  Inches(2.6), Inches(3.6), Inches(3.0), "📈", "Inflación persistente\n(aunque moderándose\nrespecto a picos post-pandemia)", NARANJA)
stat_box(s, Inches(4.6),  Inches(2.6), Inches(3.6), Inches(3.0), "💱", "Peso mexicano volátil\nfrente al dólar", VERDE)
stat_box(s, Inches(8.5),  Inches(2.6), Inches(3.6), Inches(3.0), "🌐", "T-MEC en renegociación\ny mayor proteccionismo\nen EE. UU.", MORADO)
footer(s)
notes(s, "2026 se caracteriza por inflación persistente aunque en moderación, "
         "un peso mexicano volátil frente al dólar, y tensiones comerciales "
         "por la renegociación del T-MEC y políticas más proteccionistas de "
         "Estados Unidos. Esto no elimina las oportunidades de exportación, "
         "pero obliga a privilegiar mercados cercanos y estables por encima "
         "de mercados grandes pero lejanos y riesgosos.")

# ========================================================================
# SLIDE 4 — LA EMPRESA
# ========================================================================
s = add_slide(); set_bg(s, NEGRO)
header(s, "¿Qué es Crujisanas?", "02", accent=VERDE)
bullets(s, Inches(0.65), Inches(1.6), Inches(7.2), Inches(4.5), [
    "Marca mexicana de snacks saludables",
    "Elaborados con jícama, betabel, camote y zanahoria",
    "Horneados, no fritos — sin conservadores ni azúcares añadidas",
    "Nace dentro de la cafetería de un centro de yoga",
], size=17)
stat_box(s, Inches(8.3), Inches(1.6), Inches(4.2), Inches(2.1), "~5 años", "en el mercado mexicano", NARANJA)
stat_box(s, Inches(8.3), Inches(3.9), Inches(4.2), Inches(2.1), "Elena Salame", "Fundadora, junto a un\npequeño equipo de socias", MORADO)
footer(s)
notes(s, "Crujisanas es una marca mexicana de snacks saludables elaborados con "
         "vegetales de raíz —jícama, betabel, camote y zanahoria— horneados en "
         "lugar de fritos, sin conservadores ni azúcares añadidas. Nació hace "
         "casi cinco años dentro de la cafetería de un centro de yoga, fundada "
         "por Elena Salame junto a un pequeño equipo de socias.")

# ========================================================================
# SLIDE 5 — HISTORIA Y DIFERENCIACIÓN
# ========================================================================
s = add_slide(); set_bg(s, NEGRO)
header(s, "Una respuesta a un problema real", "02", accent=VERDE)
stat_box(s, Inches(0.65), Inches(1.6), Inches(3.8), Inches(4.3),
         "🥇", "México: 1er lugar mundial\nen obesidad infantil", NARANJA)
bullets_lead(s, Inches(4.8), Inches(1.7), Inches(7.7), Inches(4.3), [
    ("Visibilidad nacional", "presentación en Shark Tank México."),
    ("Diferenciación clara", "compite contra Sabritas y Barcel, enfocadas en "
     "papa, maíz y trigo, usando vegetales de raíz con más fibra y menor "
     "densidad calórica."),
    ("Producción artesanal", "de baja escala, que le permite posicionarse "
     "como producto premium dentro de una categoría dominada por PepsiCo "
     "y Grupo Bimbo."),
], size=15)
footer(s)
notes(s, "El origen de Crujisanas responde a un problema estructural: México "
         "ocupa el primer lugar mundial en obesidad infantil. La marca ganó "
         "visibilidad nacional en Shark Tank México y se diferencia claramente "
         "de Sabritas y Barcel, que se enfocan en papa, maíz y trigo, mientras "
         "que Crujisanas usa vegetales de raíz con más fibra y menor densidad "
         "calórica, con una producción artesanal que la posiciona como premium.")

# ========================================================================
# SLIDE 6 — FORTALEZAS I
# ========================================================================
s = add_slide(); set_bg(s, NEGRO)
header(s, "¿Por qué tiene potencial? (I)", "03", accent=MORADO)
bullets_lead(s, Inches(0.65), Inches(1.6), Inches(11.7), Inches(5), [
    ("Calidad", "diseño sobrio y oscuro que comunica un producto premium "
     "desde el anaquel, aunque en esencia sea una bolsa de \"papitas\" "
     "de vegetales."),
    ("Innovación", "casi no hay competidores enfocados en jícama, camote, "
     "betabel y zanahoria; es un espacio de mercado poco explotado (blue ocean)."),
    ("Tradición", "retoma ingredientes de la dieta mexicana y los reinterpreta "
     "en un formato moderno de conveniencia."),
    ("Sustentabilidad", "vegetales de raíz con menor huella hídrica, "
     "horneados en lugar de fritos, alineados con el consumo consciente."),
], size=16)
footer(s)
notes(s, "Cuatro elementos justifican su potencial: calidad, con un diseño "
         "premium; innovación, al enfocarse en vegetales de raíz, algo poco "
         "explotado en snacks saludables; tradición, al retomar ingredientes "
         "típicamente mexicanos en un formato moderno; y sustentabilidad, "
         "por el uso de cultivos de menor huella hídrica y procesos de horneado.")

# ========================================================================
# SLIDE 7 — FORTALEZAS II + GRÁFICO
# ========================================================================
s = add_slide(); set_bg(s, NEGRO)
header(s, "¿Por qué tiene potencial? (II)", "03", accent=MORADO)
bullets_lead(s, Inches(0.65), Inches(1.55), Inches(6.4), Inches(4.8), [
    ("Prestigio", "estética \"gourmet\", alejada de la producción masiva de "
     "anaquel."),
    ("Tendencias de consumo", "boom global de wellness, alimentación "
     "consciente y \"clean label\", también presente en Latinoamérica."),
    ("Identidad mexicana", "jícama y camote como argumento de autenticidad "
     "muy valioso en mercados internacionales."),
], size=15)

chart_data = CategoryChartData()
chart_data.categories = ["Crujisanas", "Promedio\ncompetencia premium"]
chart_data.add_series("Índice de precio", (75, 100))
gframe = s.shapes.add_chart(
    XL_CHART_TYPE.COLUMN_CLUSTERED,
    Inches(7.4), Inches(1.7), Inches(5.3), Inches(4.1), chart_data
)
chart = gframe.chart
chart.has_legend = False
chart.has_title = True
chart.chart_title.text_frame.text = "Ventaja de precio (índice 100 = competencia)"
chart.chart_title.text_frame.paragraphs[0].runs[0].font.size = Pt(13)
chart.chart_title.text_frame.paragraphs[0].runs[0].font.color.rgb = CREMA
plot = chart.plots[0]
plot.has_data_labels = True
plot.data_labels.font.size = Pt(14)
plot.data_labels.font.bold = True
plot.data_labels.font.color.rgb = CREMA
series = plot.series[0]
series.format.fill.solid()
series.format.fill.fore_color.rgb = NARANJA
cat_axis = chart.category_axis
cat_axis.tick_labels.font.size = Pt(12)
cat_axis.tick_labels.font.color.rgb = CREMA
val_axis = chart.value_axis
val_axis.tick_labels.font.color.rgb = CREMA
val_axis.has_major_gridlines = False
footer(s)
notes(s, "Además, Crujisanas proyecta prestigio con una estética tipo gourmet, "
         "responde a tendencias globales de alimentación consciente, y apela "
         "a la identidad mexicana como argumento de autenticidad. Un dato clave: "
         "según su fundadora, el precio de Crujisanas es hasta 25% menor al de "
         "las marcas líderes del segmento saludable, una ventaja competitiva "
         "relevante si se mantiene esa estructura de costos al exportar.")

# ========================================================================
# SLIDE 8 — MERCADO: GUATEMALA
# ========================================================================
s = add_slide(); set_bg(s, NEGRO)
header(s, "Mercado seleccionado: Guatemala", "04", accent=VERDE)
bullets_lead(s, Inches(0.65), Inches(1.6), Inches(11.7), Inches(5), [
    ("Frontera terrestre con México", "reduce costos y riesgos logísticos "
     "frente al flete marítimo/aéreo, volátil en 2026."),
    ("Puerta de entrada al bloque CA-4", "junto con El Salvador, Honduras y "
     "Nicaragua, con libre movilidad de mercancías."),
    ("Demanda creciente", "el mercado guatemalteco de snacks y confitería "
     "crece sostenidamente hacia productos más saludables y premium."),
    ("Idioma y cultura compartidos", "con México, lo que reduce las barreras "
     "de comunicación y adaptación del mensaje publicitario."),
], size=16)
footer(s)
notes(s, "Elegimos Guatemala como mercado de entrada. Es un país fronterizo, "
         "accesible por vía terrestre, lo que reduce el riesgo logístico frente "
         "a mercados transoceánicos. Además forma parte del bloque CA-4, lo que "
         "la convierte en una puerta de entrada a todo Centroamérica. Comparte "
         "idioma y cultura con México, y su mercado de snacks saludables está "
         "en crecimiento en zonas urbanas.")

# ========================================================================
# SLIDE 9 — ANÁLISIS DEL ENTORNO I
# ========================================================================
s = add_slide(); set_bg(s, NEGRO)
header(s, "Análisis del entorno (I)", "05", accent=NARANJA)
quad_box(s, Inches(0.65), Inches(1.6), Inches(5.7), Inches(2.3), "Cultura",
         "Afinidad latinoamericana muy cercana a México; no se anticipan "
         "barreras culturales significativas.", VERDE)
quad_box(s, Inches(6.6), Inches(1.6), Inches(5.9), Inches(2.3), "Idioma",
         "Español en ambos países: no requiere traducción, aunque sí "
         "adaptar el etiquetado a normativas locales.", NARANJA)
quad_box(s, Inches(0.65), Inches(4.15), Inches(5.7), Inches(2.55),
         "Economía", "Crecimiento del PIB ~3.3% anual, pero con alta "
         "desigualdad → enfocar el lanzamiento en clase media urbana "
         "de Ciudad de Guatemala.", MORADO)
quad_box(s, Inches(6.6), Inches(4.15), Inches(5.9), Inches(2.55),
         "Competencia", "Sabritas (PepsiCo) y Barcel (Grupo Bimbo) ya "
         "consolidadas → diferenciarse como opción saludable/premium, "
         "no competir en precio ni volumen.", VERDE)
footer(s)
notes(s, "En cultura e idioma, Guatemala y México son prácticamente idénticos, "
         "lo que facilita mucho la adaptación. En economía, Guatemala crece "
         "cerca de 3.3% anual, pero con una desigualdad muy marcada, por lo "
         "que la estrategia debe enfocarse en un segmento urbano de clase "
         "media, no en una penetración masiva. Y en competencia, Sabritas y "
         "Barcel ya tienen presencia consolidada, así que Crujisanas debe "
         "diferenciarse claramente como opción saludable y premium.")

# ========================================================================
# SLIDE 10 — ANÁLISIS DEL ENTORNO II
# ========================================================================
s = add_slide(); set_bg(s, NEGRO)
header(s, "Análisis del entorno (II)", "05", accent=NARANJA)
quad_box(s, Inches(0.65), Inches(1.6), Inches(5.7), Inches(2.3), "Regulación",
         "RTCA exige etiquetado nutricional en español; posible etiquetado "
         "frontal (Iniciativa 5504) — Crujisanas ya cumple el mexicano.", MORADO)
quad_box(s, Inches(6.6), Inches(1.6), Inches(5.9), Inches(2.3), "Tecnología",
         "E-commerce y redes sociales crecen sostenidamente, sobre todo "
         "entre jóvenes urbanos: viable el marketing digital de bajo costo.", NARANJA)
quad_box(s, Inches(0.65), Inches(4.15), Inches(5.7), Inches(2.55), "Tendencias",
         "Crecimiento de snacks premium y saludables en centros urbanos, "
         "coincidiendo con la propuesta de valor de Crujisanas.", VERDE)
quad_box(s, Inches(6.6), Inches(4.15), Inches(5.9), Inches(2.55), "Factor macro 2026",
         "Aislado de la disputa arancelaria México–EE. UU., pero expuesto "
         "al tipo de cambio y a insumos importados (empaque, aditivos).", MORADO)
footer(s)
notes(s, "En regulación, Guatemala exige etiquetado nutricional bajo el RTCA "
         "y discute una ley de sellos de advertencia frontal; si se aprueba, "
         "Crujisanas ya estaría preparada por cumplir la norma mexicana. "
         "En tecnología, el comercio electrónico y las redes sociales crecen, "
         "permitiendo una entrada digital de bajo costo. Las tendencias de "
         "consumo saludable favorecen al producto, y aunque el país queda "
         "fuera de la disputa arancelaria entre México y Estados Unidos, sí "
         "puede verse afectado por el tipo de cambio y el costo de insumos importados.")

# ========================================================================
# SLIDE 11 — ESTRATEGIA / MARKETING MIX
# ========================================================================
s = add_slide(); set_bg(s, NEGRO)
header(s, "Propuesta de internacionalización", "06", accent=VERDE)
quad_box(s, Inches(0.65), Inches(1.6), Inches(5.7), Inches(2.3), "Producto",
         "Se conserva el diseño; se agregan presentaciones/porciones más "
         "pequeñas y accesibles. Nombre y marca sin cambios.", NARANJA)
quad_box(s, Inches(6.6), Inches(1.6), Inches(5.9), Inches(2.3), "Precio",
         "Ajuste al poder adquisitivo local, sin sacrificar el "
         "posicionamiento premium.", MORADO)
quad_box(s, Inches(0.65), Inches(4.15), Inches(5.7), Inches(2.55), "Plaza",
         "Distribuidor local con relación en supermercados de Ciudad de "
         "Guatemala (ej. La Torre); entrada gradual antes de escalar.", VERDE)
quad_box(s, Inches(6.6), Inches(4.15), Inches(5.9), Inches(2.55), "Promoción",
         "Marketing digital (Instagram/TikTok) + degustaciones en punto "
         "de venta para generar prueba de producto.", NARANJA)
footer(s)
notes(s, "La propuesta de entrada es gradual: en producto, se mantiene el "
         "diseño pero se ofrecen porciones más pequeñas y accesibles; en "
         "precio, se ajusta al poder adquisitivo local sin perder el "
         "posicionamiento premium; en plaza, se recomienda trabajar con un "
         "distribuidor local que ya tenga relación con supermercados, en "
         "lugar de distribución directa; y en promoción, se apuesta por "
         "redes sociales y degustaciones en punto de venta, dado que es una "
         "marca desconocida en ese mercado.")

# ========================================================================
# SLIDE 12 — RETOS
# ========================================================================
s = add_slide(); set_bg(s, NEGRO)
header(s, "Principales retos", "07", accent=MORADO)
bullets_lead(s, Inches(0.65), Inches(1.6), Inches(11.7), Inches(5), [
    ("Competencia consolidada", "Sabritas y Barcel tienen enorme poder de "
     "distribución y economías de escala."),
    ("Desigualdad económica", "el mercado real accesible para un producto "
     "premium es más reducido que la población total del país."),
    ("Incertidumbre macroeconómica 2026", "la volatilidad cambiaria y la "
     "inflación global obligan a revisar precios periódicamente."),
    ("Cumplimiento regulatorio", "posible entrada en vigor de un etiquetado "
     "frontal de advertencia (Iniciativa 5504)."),
    ("Escalamiento de producción", "de un modelo artesanal de baja escala "
     "hacia volúmenes de exportación sostenidos."),
], size=15.5)
footer(s)
notes(s, "Los principales retos son cinco: enfrentar competencia consolidada "
         "con enorme poder de distribución; un mercado meta con fuerte "
         "desigualdad económica, que reduce el segmento realmente accesible; "
         "la incertidumbre macroeconómica de 2026, que exige revisar precios "
         "constantemente; el cumplimiento regulatorio ante un posible "
         "etiquetado frontal; y el reto operativo de escalar una producción "
         "hoy artesanal hacia volúmenes de exportación sostenidos.")

# ========================================================================
# SLIDE 13 — CONCLUSIONES
# ========================================================================
s = add_slide(); set_bg(s, NEGRO)
header(s, "Conclusiones", "07", accent=VERDE)
bullets(s, Inches(0.65), Inches(1.6), Inches(11.7), Inches(4.8), [
    "Un producto no necesita ser fabricado por una multinacional para tener "
    "potencial exportador: basta una propuesta de valor clara y un mercado "
    "que la necesite.",
    "La cercanía geográfica y cultural puede pesar tanto como el tamaño del "
    "mercado, sobre todo en un entorno incierto como el de 2026.",
    "La expansión hacia Guatemala se considera VIABLE, siempre que se "
    "mantenga el posicionamiento premium/artesanal, sin competir en "
    "volumen contra los grandes conglomerados.",
], size=17)
footer(s)
notes(s, "En conclusión, aprendimos que un producto no necesita ser de una "
         "multinacional para tener potencial exportador: basta una propuesta "
         "de valor clara —en este caso, salud, identidad mexicana y precio "
         "competitivo— y un mercado que la necesite. También aprendimos que "
         "la cercanía geográfica y cultural puede ser tan importante como el "
         "tamaño del mercado en un entorno incierto como 2026. Consideramos "
         "que la expansión hacia Guatemala es viable, siempre que Crujisanas "
         "mantenga su posicionamiento premium y no intente competir en volumen "
         "contra Sabritas o Barcel.")

# ========================================================================
# SLIDE 14 — CIERRE
# ========================================================================
s = add_slide(); set_bg(s, NEGRO)
triangle(s, Inches(0.6), Inches(0.6), Inches(1.4), Inches(1.4), VERDE, rotation=20)
triangle(s, Inches(1.7), Inches(0.4), Inches(1.1), Inches(1.1), NARANJA, rotation=-15)
txt(s, Inches(0.7), Inches(2.8), Inches(10), Inches(1.2), "¡Gracias!", size=56, color=CREMA, bold=True)
txt(s, Inches(0.75), Inches(4.0), Inches(9), Inches(0.6), "¿Preguntas o comentarios?", size=20, color=NARANJA)
txt(s, Inches(0.75), Inches(6.3), Inches(11), Inches(1.0),
    "Fuentes: Shark Tank México (YouTube) · Sweets & Snacks International · SESAN Guatemala (Iniciativa 5504)",
    size=10, color=GRIS)
footer(s)
notes(s, "Muchas gracias por su atención. Quedo abierta/o a preguntas sobre "
         "el análisis, el mercado elegido o la estrategia propuesta.")

# ------------------------------------------------------------------------
prs.save("Crujisanas_Expansion_Internacional.pptx")
print("Presentación generada: Crujisanas_Expansion_Internacional.pptx")