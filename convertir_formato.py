from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

doc = Document()
sec = doc.sections[0]
sec.top_margin = Inches(0.6)
sec.bottom_margin = Inches(0.65)
sec.left_margin = Inches(0.85)
sec.right_margin = Inches(0.85)
style = doc.styles['Normal']
style.font.name = 'Arial'
style.font.size = Pt(11)
style.paragraph_format.space_after = Pt(0)
style.paragraph_format.line_spacing = 1.05

def para(text='', bold=False, size=None, align=None, after=0):
    p = doc.add_paragraph()
    if align is not None: p.alignment = align
    p.paragraph_format.space_after = Pt(after)
    r = p.add_run(text)
    r.bold = bold
    if size: r.font.size = Pt(size)
    return p

para('GLAMOUR ESENCIA D`EFRAIN, E.I.R.L.', True, 13, WD_ALIGN_PARAGRAPH.CENTER)
para('Autopista Coronel Fernández Domínguez, Plaza Mónaco, local 205, Urbanización Italia,', size=9, align=WD_ALIGN_PARAGRAPH.CENTER)
para('Santo Domingo Este, Tel: 809-788-0375, correo asesores.fiscales.isr.rd@gmail.com', size=9, align=WD_ALIGN_PARAGRAPH.CENTER)
para('R. N. C. No. 130-74247-2', size=10, align=WD_ALIGN_PARAGRAPH.CENTER, after=34)
para('Santo Domingo Este.')
para('14 de noviembre del año 2026.', after=30)
para('Señores')
para('CAMARA DE COMERCIO Y PRODUCCION DE LA')
para('PROVINCIA DE SANTO DOMINGO, CCPSD')
para('Ciudad. -', after=14)
para('Asunto: Autorización de Representante ante la Cámara de Comercio\n        Para Renovación de Registro Mercantil No.125911SD', after=17)
para('Distinguidos Señores', after=14)
p = para('Quien suscribe, señor EFRAIN ROJAS SORIANO, portador de la cedula de identificación personal y electoral No. 001-1567450-9, por medio de la presente, estamos autorizando al Lic. NERYS SANCHEZ MATOS, portador de la cédula de identidad y electoral No.001-0048723-0, con e-mail asesores.fiscales.isr.rd@gmail.com como representante ante esa Cámara de Comercio y Producción de la Provincia de Santo Domingo, con la finalidad de solicitar y tramitar la Renovación de nuestro Registro Mercantil, de nuestra empresa GLAMOUR ESENCIA D`EFRAIN, E.I.R.L., con RNC No. 130-74247-2.', after=23)
p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
para('En espera de contar con su colaboración, les saluda', after=23)
para('Atentamente,', after=35)
para('Sr. EFRAIN ROJAS SORIANO')
para('Gerente')
filename = 'CARTA_AUTORIZACION_CAMARA_COMERCIO.docx'
doc.save(filename)
print(filename)
