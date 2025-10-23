"""
Quote Generator - Converts quote data to editable .docx format

This script demonstrates how to generate an editable Word document
that replicates the structure of a PDF quote. Sales reps can then
modify prices, add disclaimers, or make other adjustments easily.
"""

from quote_builder import QuoteBuilder
from constants import DEFAULT_OUTPUT_FILENAME


def generate_quote_docx(output_filename=DEFAULT_OUTPUT_FILENAME):
    """
    Main function to generate the quote document using the
    Builder pattern.

    Args:
        output_filename: Name of the output file (default from
                         constants)

    Returns:
        str: The output filename
    """
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
            'descripcion': (
                'MOON ALU Sillon Individual\n'
                'Color A14 Grafito, Cuerda R8\n'
                'Dark Grey, Tela C17W Dark Grey'
            ),
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
            'descripcion': (
                'MOON ALU Loveseat Color A14\n'
                'Grafito, Cuerda R8 Dark Grey,\n'
                'Tela C17W Dark Grey'
            ),
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

    # Build the document using the Builder pattern
    builder = QuoteBuilder()
    (builder.with_header()
            .with_customer_info(customer_data, quote_data)
            .with_products(products)
            .with_summary(summary)
            .with_signature()
            .with_banking_info()
            .with_terms()
            .with_footer()
            .save(output_filename))

    return output_filename


if __name__ == '__main__':
    generate_quote_docx()
