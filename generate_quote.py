"""
Quote Generator - Converts quote data to editable .docx format

This script demonstrates how to generate an editable Word document
that replicates the structure of a PDF quote. Sales reps can then
modify prices, add disclaimers, or make other adjustments easily.
"""

from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from datetime import datetime
from io import BytesIO
from PIL import Image, ImageDraw, ImageFont
import random


def generate_placeholder_image(width, height, seed=None, text=""):
    """
    Generate a placeholder image locally using PIL/Pillow.

    This function creates colorful placeholder images for testing purposes.
    In production, you can replace this with:
    - Actual product images from your database/storage
    - Images from URLs (e.g., product.image_url)
    - Images from local file paths
    - External services like https://picsum.photos/200/300 (if accessible)

    Example of using a real image from a URL:
        response = requests.get(product.image_url)
        return BytesIO(response.content)

    Args:
        width: Image width in pixels
        height: Image height in pixels
        seed: Optional seed for consistent colors
        text: Text to display on the image
    Returns:
        BytesIO object containing the image data
    """
    # Set random seed for consistent colors
    if seed is not None:
        random.seed(seed)

    # Generate a random pastel color
    r = random.randint(100, 255)
    g = random.randint(100, 255)
    b = random.randint(100, 255)
    bg_color = (r, g, b)

    # Create image
    img = Image.new('RGB', (width, height), color=bg_color)
    draw = ImageDraw.Draw(img)

    # Add text if provided
    if text:
        try:
            # Try to use a default font
            font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 40)
        except:
            font = ImageFont.load_default()

        # Calculate text position (centered)
        bbox = draw.textbbox((0, 0), text, font=font)
        text_width = bbox[2] - bbox[0]
        text_height = bbox[3] - bbox[1]
        position = ((width - text_width) // 2, (height - text_height) // 2)

        # Draw text with shadow for better visibility
        shadow_color = (50, 50, 50)
        draw.text((position[0] + 2, position[1] + 2), text, fill=shadow_color, font=font)
        draw.text(position, text, fill=(255, 255, 255), font=font)

    # Save to BytesIO
    img_io = BytesIO()
    img.save(img_io, format='PNG')
    img_io.seek(0)

    return img_io


def set_cell_border(cell, **kwargs):
    """
    Set cell borders
    Usage:
        set_cell_border(cell, top={"sz": 12, "val": "single", "color": "#000000"})
    """
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()

    for edge in ('top', 'left', 'bottom', 'right'):
        if edge in kwargs:
            edge_data = kwargs.get(edge)
            edge_el = OxmlElement(f'w:{edge}')
            for key, value in edge_data.items():
                edge_el.set(qn(f'w:{key}'), str(value))
            tcPr.append(edge_el)


def set_cell_background(cell, fill):
    """Set cell background color"""
    shading_elm = OxmlElement('w:shd')
    shading_elm.set(qn('w:fill'), fill)
    cell._element.get_or_add_tcPr().append(shading_elm)


def add_header_section(doc):
    """Add company header information"""
    # Add logo
    logo_para = doc.add_paragraph()
    logo_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    try:
        logo_image = generate_placeholder_image(400, 200, seed=1, text="LOGO")
        logo_run = logo_para.add_run()
        logo_run.add_picture(logo_image, width=Inches(2.5))
    except Exception as e:
        print(f"Warning: Could not add logo image: {e}")
        logo_para.add_run('[LOGO]')

    # Company name and website
    header = doc.add_paragraph()
    header.alignment = WD_ALIGN_PARAGRAPH.CENTER

    run = header.add_run('www.exteriorconcept.com\n')
    run.font.size = Pt(10)

    run = header.add_run('IMPORTADORA DE MUEBLES DE TERRAZA Y JARDÍN S.A. DE C.V.\n\n')
    run.font.size = Pt(11)
    run.font.bold = True

    run = header.add_run('Polanco\n')
    run.font.size = Pt(12)
    run.font.bold = True

    run = header.add_run('Calle Julio Verne 19, Polanco, Miguel Hidalgo, CDMX\n')
    run.font.size = Pt(9)

    run = header.add_run('Tel: 55 5105 1202\n')
    run.font.size = Pt(9)

    run = header.add_run('info@exteriorconcept.com')
    run.font.size = Pt(9)

    doc.add_paragraph()  # Spacer


def add_customer_info_section(doc, customer_data, quote_data):
    """Add customer and quote information in two columns"""
    table = doc.add_table(rows=5, cols=2)
    table.autofit = False
    table.allow_autofit = False

    # Left column - Customer info
    cells = table.rows[0].cells
    p = cells[0].paragraphs[0]
    run = p.add_run('NOMBRE: ')
    run.font.bold = True
    run = p.add_run(customer_data['nombre'])

    # Right column - Quote number
    p = cells[1].paragraphs[0]
    run = p.add_run('NO. DE COTIZACIÓN: ')
    run.font.bold = True
    run = p.add_run(quote_data['numero'])

    # Row 2
    cells = table.rows[1].cells
    p = cells[0].paragraphs[0]
    run = p.add_run('TELÉFONO: ')
    run.font.bold = True
    run = p.add_run(customer_data['telefono'])

    p = cells[1].paragraphs[0]
    run = p.add_run('FECHA: ')
    run.font.bold = True
    run = p.add_run(quote_data['fecha'])

    # Row 3
    cells = table.rows[2].cells
    p = cells[0].paragraphs[0]
    run = p.add_run('DIRECCIÓN: ')
    run.font.bold = True
    run = p.add_run(customer_data.get('direccion', ''))

    p = cells[1].paragraphs[0]
    run = p.add_run('INFORMACIÓN CONTACTO:')
    run.font.bold = True

    # Row 4
    cells = table.rows[3].cells
    p = cells[0].paragraphs[0]
    run = p.add_run('MAIL: ')
    run.font.bold = True
    run = p.add_run(customer_data.get('mail', ''))

    p = cells[1].paragraphs[0]
    p.add_run(quote_data['contacto'])

    # Row 5
    cells = table.rows[4].cells
    cells[0].text = ''
    p = cells[1].paragraphs[0]
    p.add_run(quote_data.get('email_contacto', ''))

    doc.add_paragraph()  # Spacer


def add_products_table(doc, products):
    """Add products table with all details"""
    # Create table with header row
    table = doc.add_table(rows=1, cols=9)
    table.style = 'Light Grid Accent 1'

    # Header row
    header_cells = table.rows[0].cells
    headers = ['MARCA', 'DESCRIPCIÓN', 'IMAGEN\nREFERENCIA', 'CÓDIGO', 'PZ',
                'PRECIO\nUNITARIO', 'DESCUENTO', 'PRECIO\nDESCUENTO', 'TOTAL']

    for i, header in enumerate(headers):
        cell = header_cells[i]
        set_cell_background(cell, '000000')
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(header)
        run.font.bold = True
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(255, 255, 255)

    # Add product rows
    for idx, product in enumerate(products):
        row_cells = table.add_row().cells

        # Marca
        row_cells[0].text = product['marca']

        # Descripción
        row_cells[1].text = product['descripcion']

        # Imagen Referencia
        p = row_cells[2].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        try:
            # Use different seed for each product to get different images
            product_image = generate_placeholder_image(300, 300, seed=idx + 10, text=f"P{idx+1}")
            img_run = p.add_run()
            img_run.add_picture(product_image, width=Inches(1.2))
        except Exception as e:
            print(f"Warning: Could not add product image {idx}: {e}")
            p.add_run('[IMAGEN]')

        # Código
        row_cells[3].text = product['codigo']

        # Piezas
        p = row_cells[4].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run(str(product['piezas']))

        # Precio Unitario
        p = row_cells[5].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p.add_run(f"$ {product['precio_unitario']:,.2f}")

        # Descuento
        p = row_cells[6].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run(f"{product['descuento_pct']:.1f}%\n")
        p.add_run(f"$ {product['descuento_monto']:,.2f}")

        # Precio con Descuento
        p = row_cells[7].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p.add_run(f"$ {product['precio_descuento']:,.2f}")

        # Total
        p = row_cells[8].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p.add_run(f"$ {product['total']:,.2f}")

    return table


def add_summary_section(doc, summary):
    """Add financial summary table"""
    # Create a right-aligned table for the summary
    table = doc.add_table(rows=4, cols=2)
    table.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    # Set column widths
    table.columns[0].width = Inches(4.5)
    table.columns[1].width = Inches(2.0)

    summary_items = [
        ('SUBTOTAL', summary['subtotal']),
        ('DESCUENTO', summary['descuento']),
        ('TOTAL', summary['total']),
        (f"ANTICIPO ({summary['anticipo_pct']:.1f}%)", summary['anticipo'])
    ]

    for i, (label, value) in enumerate(summary_items):
        row = table.rows[i]

        # Label cell
        set_cell_background(row.cells[0], '000000')
        p = row.cells[0].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run = p.add_run(label)
        run.font.bold = True
        run.font.color.rgb = RGBColor(255, 255, 255)

        # Value cell
        p = row.cells[1].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run = p.add_run(f"$ {value:,.2f}")
        run.font.bold = True


def add_signature_section(doc):
    """Add signature section"""
    doc.add_paragraph()
    doc.add_paragraph('_' * 60)

    p = doc.add_paragraph('NOMBRE Y FIRMA DEL CLIENTE')
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.runs[0]
    run.font.bold = True

    p = doc.add_paragraph('ACEPTO TÉRMINOS Y CONDICIONES')
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.runs[0]
    run.font.bold = True


def add_banking_info(doc):
    """Add banking information table"""
    doc.add_page_break()

    table = doc.add_table(rows=4, cols=4)
    table.style = 'Light Grid Accent 1'

    # Header row
    headers = ['BANCO', 'BBVA', 'SANTANDER', 'MIFEL']
    for i, header in enumerate(headers):
        cell = table.rows[0].cells[i]
        if i == 0:
            set_cell_background(cell, '000000')
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(header)
        run.font.bold = True
        if i == 0:
            run.font.color.rgb = RGBColor(255, 255, 255)

    # Beneficiario
    table.rows[1].cells[0].text = 'BENEFICIARIO'
    set_cell_background(table.rows[1].cells[0], '000000')
    table.rows[1].cells[0].paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
    table.rows[1].cells[0].paragraphs[0].runs[0].font.bold = True

    for i in range(1, 4):
        table.rows[1].cells[i].text = 'IMPORTADORA DE MUEBLES DE TERRAZA Y JARDIN'

    # Número de cuenta
    table.rows[2].cells[0].text = 'NÚMERO DE CUENTA'
    set_cell_background(table.rows[2].cells[0], '000000')
    table.rows[2].cells[0].paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
    table.rows[2].cells[0].paragraphs[0].runs[0].font.bold = True

    table.rows[2].cells[1].text = '0191089269'
    table.rows[2].cells[2].text = '65507881953'
    table.rows[2].cells[3].text = '01600007854'

    # CLABE
    table.rows[3].cells[0].text = 'CLABE'
    set_cell_background(table.rows[3].cells[0], '000000')
    table.rows[3].cells[0].paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
    table.rows[3].cells[0].paragraphs[0].runs[0].font.bold = True

    table.rows[3].cells[1].text = '0121-8000-1910-8926-97'
    table.rows[3].cells[2].text = '0141-8065-5078-8195-39'
    table.rows[3].cells[3].text = '0421-8001-6000-0785-48'


def add_terms_and_conditions(doc):
    """Add commercial policies and terms"""
    doc.add_paragraph()

    heading = doc.add_paragraph('Políticas Comerciales de Exterior Concept')
    heading.runs[0].font.bold = True
    heading.runs[0].font.size = Pt(14)

    # PAGOS section
    section = doc.add_paragraph('PAGOS')
    section.runs[0].font.bold = True
    section.runs[0].font.size = Pt(11)

    terms_pagos = [
        'Al realizar cualquier pago o anticipo, el cliente acepta las políticas comerciales vigentes.',
        'Las cotizaciones tienen una vigencia de 30 días naturales y están sujetas a cambios hasta recibir el anticipo.',
        'El cliente es responsable de revisar su pedido antes de confirmar.',
        'No se aceptan cambios ni cancelaciones una vez confirmado.',
        'En caso de cancelación, se retendrá el 25% del total del pedido más comisiones bancarias aplicables.'
    ]

    for term in terms_pagos:
        p = doc.add_paragraph(term, style='List Bullet')
        p.paragraph_format.left_indent = Inches(0.25)

    # Note: Additional terms sections would be added similarly
    doc.add_paragraph()
    p = doc.add_paragraph('[Términos y condiciones completos pueden ser añadidos aquí...]')
    p.runs[0].font.italic = True


def generate_quote_docx(output_filename='quote_example.docx'):
    """
    Main function to generate the quote document
    """
    doc = Document()

    # Sample data matching the PDF
    customer_data = {
        'nombre': 'Miguel Javier',
        'telefono': '55667788',
        'direccion': '',
        'mail': ''
    }

    quote_data = {
        'numero': 'POL-13-09-25-14-28',
        'fecha': '13 de September 2025',
        'contacto': 'Miguel Cabral',
        'email_contacto': 'miguel@pcuervo.com'
    }

    products = [
        {
            'marca': 'TALENTI\nITALIA',
            'descripcion': 'MOON ALU Sillon Individual\nColor A14 Grafito, Cuerda R8\nDark Grey, Tela C17W Dark Grey',
            'codigo': '15-TLN-MO\nNALUPL',
            'piezas': 1,
            'precio_unitario': 35689.66,
            'descuento_pct': 10.0,
            'descuento_monto': 3568.97,
            'precio_descuento': 32120.69,
            'total': 32120.69
        },
        {
            'marca': 'TALENTI\nITALIA',
            'descripcion': 'MOON ALU Loveseat Color A14\nGrafito, Cuerda R8 Dark Grey,\nTela C17W Dark Grey',
            'codigo': '15-TLN-MO\nNALUDIV2',
            'piezas': 1,
            'precio_unitario': 69913.79,
            'descuento_pct': 10.0,
            'descuento_monto': 6991.38,
            'precio_descuento': 62922.41,
            'total': 62922.41
        }
    ]

    summary = {
        'subtotal': 105603.45,
        'descuento': 10560.34,
        'total': 95043.10,
        'anticipo_pct': 50.0,
        'anticipo': 47521.55
    }

    # Build the document
    add_header_section(doc)
    add_customer_info_section(doc, customer_data, quote_data)
    add_products_table(doc, products)
    doc.add_paragraph()  # Spacer
    add_summary_section(doc, summary)
    add_signature_section(doc)
    add_banking_info(doc)
    add_terms_and_conditions(doc)

    # Add footer with locations
    section = doc.sections[0]
    footer = section.footer
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run('POLANCO - ALTAVISTA - LOMAS - INTERLOMAS - PUEBLA - CANCÚN - CUERNAVACA - GUADALAJARA - LOS CABOS')
    run.font.size = Pt(8)

    # Save the document
    doc.save(output_filename)
    print(f"✓ Quote document generated successfully: {output_filename}")
    print(f"✓ The document is fully editable in Microsoft Word, Google Docs, or LibreOffice")
    return output_filename


if __name__ == '__main__':
    generate_quote_docx()
