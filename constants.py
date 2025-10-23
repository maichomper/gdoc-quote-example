"""
Constants and configuration values for quote generation.

This module contains all text strings, URLs, banking information,
and other configuration values used throughout the quote generator.
"""

# Company Information
COMPANY_WEBSITE = 'www.exteriorconcept.com'
COMPANY_NAME = 'IMPORTADORA DE MUEBLES DE TERRAZA Y JARDÍN S.A. DE C.V.'
COMPANY_LOCATION = 'Polanco'
COMPANY_ADDRESS = (
    'Calle Julio Verne 19, Polanco, Miguel Hidalgo, CDMX'
)
COMPANY_PHONE = 'Tel: 55 5105 1202'
COMPANY_EMAIL = 'info@exteriorconcept.com'

# Table Headers
PRODUCT_TABLE_HEADERS = [
    'MARCA',
    'DESCRIPCIÓN',
    'IMAGEN\nREFERENCIA',
    'CÓDIGO',
    'PZ',
    'PRECIO\nUNITARIO',
    'DESCUENTO',
    'PRECIO\nDESCUENTO',
    'TOTAL'
]

# Banking Information
BANK_BENEFICIARY = 'IMPORTADORA DE MUEBLES DE TERRAZA Y JARDIN'

BANKING_INFO = {
    'banks': ['BBVA', 'SANTANDER', 'MIFEL'],
    'accounts': ['0191089269', '65507881953', '01600007854'],
    'clabe': [
        '0121-8000-1910-8926-97',
        '0141-8065-5078-8195-39',
        '0421-8001-6000-0785-48'
    ]
}

# Labels
LABEL_NOMBRE = 'NOMBRE: '
LABEL_TELEFONO = 'TELÉFONO: '
LABEL_DIRECCION = 'DIRECCIÓN: '
LABEL_MAIL = 'MAIL: '
LABEL_NO_COTIZACION = 'NO. DE COTIZACIÓN: '
LABEL_FECHA = 'FECHA: '
LABEL_INFO_CONTACTO = 'INFORMACIÓN CONTACTO:'

# Summary Labels
SUMMARY_SUBTOTAL = 'SUBTOTAL'
SUMMARY_DESCUENTO = 'DESCUENTO'
SUMMARY_TOTAL = 'TOTAL'
SUMMARY_ANTICIPO_TEMPLATE = 'ANTICIPO ({:.1f}%)'

# Signature Section
SIGNATURE_LINE = '_' * 60
SIGNATURE_TEXT_NAME = 'NOMBRE Y FIRMA DEL CLIENTE'
SIGNATURE_TEXT_ACCEPT = 'ACEPTO TÉRMINOS Y CONDICIONES'

# Banking Table Headers
BANKING_HEADER = 'BANCO'
BANKING_ROW_BENEFICIARIO = 'BENEFICIARIO'
BANKING_ROW_CUENTA = 'NÚMERO DE CUENTA'
BANKING_ROW_CLABE = 'CLABE'

# Terms and Conditions
TERMS_HEADING = 'Políticas Comerciales de Exterior Concept'
TERMS_SECTION_PAGOS = 'PAGOS'

TERMS_PAGOS_LIST = [
    ('Al realizar cualquier pago o anticipo, el cliente acepta '
     'las políticas comerciales vigentes.'),
    ('Las cotizaciones tienen una vigencia de 30 días naturales '
     'y están sujetas a cambios hasta recibir el anticipo.'),
    'El cliente es responsable de revisar su pedido antes de confirmar.',
    'No se aceptan cambios ni cancelaciones una vez confirmado.',
    ('En caso de cancelación, se retendrá el 25% del total del pedido '
     'más comisiones bancarias aplicables.')
]

TERMS_ADDITIONAL_NOTE = (
    '[Términos y condiciones completos pueden ser añadidos aquí...]'
)

# Footer
DEFAULT_FOOTER_TEXT = (
    'POLANCO - ALTAVISTA - LOMAS - INTERLOMAS - PUEBLA - CANCÚN - '
    'CUERNAVACA - GUADALAJARA - LOS CABOS'
)

# Styling Constants
HEADER_BACKGROUND_COLOR = '000000'
CELL_BG_BLACK = '000000'

# Font Sizes (in points)
FONT_SIZE_SMALL = 8
FONT_SIZE_NORMAL = 9
FONT_SIZE_MEDIUM = 10
FONT_SIZE_LARGE = 11
FONT_SIZE_XLARGE = 12
FONT_SIZE_HEADING = 14

# Image Settings
DEFAULT_FONT_PATH = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
DEFAULT_FONT_SIZE = 40
IMAGE_FORMAT = 'PNG'

# Document Settings
DEFAULT_OUTPUT_FILENAME = 'quote_example.docx'
SUCCESS_MESSAGE_TEMPLATE = "✓ Quote document generated successfully: {}"
EDITABLE_MESSAGE = (
    "✓ The document is fully editable in Microsoft Word, "
    "Google Docs, or LibreOffice"
)
