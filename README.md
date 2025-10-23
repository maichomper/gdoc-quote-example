# Quote Generator - PDF to .docx Conversion POC

## Overview

This proof of concept demonstrates how to generate **editable .docx documents** instead of static PDFs for sales quotes. This allows sales reps to easily modify prices, add disclaimers, or adjust content without recreating the entire document.

## Problem Statement

The current PDF-based quote system has limitations:
- PDFs are not editable
- Not all product images are available via API
- Sales reps need to make special price adjustments for customers
- Additional disclaimers must be added manually
- Sales reps have to recreate quotes from scratch in Google Docs/Word

## Solution

Generate quotes as **editable .docx files** that:
- ✅ Maintain the same professional layout as PDFs
- ✅ Can be opened and edited in Word, Google Docs, or LibreOffice
- ✅ Allow sales reps to modify prices on the fly
- ✅ Support adding/removing disclaimers and notes
- ✅ Can include product images (or placeholders)
- ✅ Preserve formatting and tables

## Files in this Project

- `generate_quote.py` - Main script to generate .docx quotes
- `verify_docx.py` - Verification script to inspect generated documents
- `quote_example.docx` - Sample generated quote (editable!)
- `pdfs/POL-13-09-25-14-28.pdf` - Original PDF reference

## Quick Start

### Installation

```bash
pip install python-docx pillow
```

### Generate a Quote

```bash
python generate_quote.py
```

This will create `quote_example.docx` with a fully formatted quote.

### Verify the Output

```bash
python verify_docx.py
```

This shows the structure and content of the generated document.

## Document Structure

The generated .docx includes:

1. **Company Header** - Logo, address, contact info
2. **Customer Information** - Name, phone, email, address
3. **Quote Details** - Quote number, date, contact person
4. **Products Table** - Brand, description, images, codes, pricing
5. **Financial Summary** - Subtotal, discounts, total, advance payment
6. **Signature Section** - Customer name and acceptance
7. **Banking Information** - Multiple bank account details
8. **Terms & Conditions** - Commercial policies

## Key Features

### Fully Editable
- Sales reps can double-click any cell to edit
- Prices can be adjusted in the table
- Rows can be added/removed
- Custom disclaimers can be inserted anywhere

### Professional Formatting
- Tables with borders and shading
- Proper currency formatting ($ 1,234.56)
- Bold headers and labels
- Multi-column layouts

### Image Support
- Product images can be embedded
- Currently uses placeholders `[IMAGEN]`
- Can be replaced with actual product photos

### Cross-Platform
- Works with Microsoft Word
- Compatible with Google Docs (upload to Drive)
- Opens in LibreOffice Writer
- Can be converted back to PDF when final

## Integration with Your Agent

The `generate_quote_docx()` function accepts data structures that can easily come from your API:

```python
customer_data = {
    'nombre': 'Customer Name',
    'telefono': '55667788',
    'direccion': 'Address',
    'mail': 'email@example.com'
}

quote_data = {
    'numero': 'POL-13-09-25-14-28',
    'fecha': '13 de September 2025',
    'contacto': 'Sales Rep Name',
    'email_contacto': 'rep@company.com'
}

products = [
    {
        'marca': 'TALENTI\\nITALIA',
        'descripcion': 'Product description...',
        'codigo': 'PROD-CODE',
        'piezas': 1,
        'precio_unitario': 35689.66,
        'descuento_pct': 10.0,
        'descuento_monto': 3568.97,
        'precio_descuento': 32120.69,
        'total': 32120.69
    }
]

# Generate the document
generate_quote_docx(output_filename='custom_quote.docx')
```

## Next Steps for Production

### 1. Add Product Images
Instead of `[IMAGEN]` placeholders, embed actual product photos:

```python
from docx.shared import Inches

# In the products table
if product.get('image_path'):
    paragraph = row_cells[2].paragraphs[0]
    run = paragraph.add_run()
    run.add_picture(product['image_path'], width=Inches(1.0))
```

### 2. Integrate with FastAPI
Create an endpoint that generates quotes on demand:

```python
from fastapi import FastAPI
from fastapi.responses import FileResponse

app = FastAPI()

@app.post("/generate-quote")
async def create_quote(quote_data: QuoteRequest):
    filename = generate_quote_docx(
        customer_data=quote_data.customer,
        products=quote_data.products,
        output_filename=f"quotes/{quote_data.quote_number}.docx"
    )
    return FileResponse(
        filename,
        media_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        filename=f"{quote_data.quote_number}.docx"
    )
```

### 3. Upload to Google Drive (Optional)
If you want to automatically upload to Google Drive:

```python
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

def upload_to_drive(local_file, folder_id=None):
    credentials = service_account.Credentials.from_service_account_file(
        'service-account.json',
        scopes=['https://www.googleapis.com/auth/drive.file']
    )

    service = build('drive', 'v3', credentials=credentials)

    file_metadata = {
        'name': os.path.basename(local_file),
        'parents': [folder_id] if folder_id else []
    }

    media = MediaFileUpload(local_file, mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')

    file = service.files().create(
        body=file_metadata,
        media_body=media,
        fields='id, webViewLink'
    ).execute()

    return file.get('webViewLink')
```

### 4. Convert to Google Docs Format
After uploading to Drive, convert to native Google Docs:

```python
file_metadata = {
    'name': 'Quote Document',
    'mimeType': 'application/vnd.google-apps.document'
}

media = MediaFileUpload('quote.docx', mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')

file = service.files().create(
    body=file_metadata,
    media_body=media,
    fields='id, webViewLink'
).execute()
```

## Comparison: .docx vs Google Docs API

### .docx (Recommended to start)
**Pros:**
- ✅ No OAuth setup needed
- ✅ Works offline
- ✅ Simple to implement
- ✅ Can be uploaded to Drive later
- ✅ Compatible with multiple platforms

**Cons:**
- ❌ Requires file download/upload for cloud editing
- ❌ Not real-time collaborative

### Google Docs API
**Pros:**
- ✅ Cloud-native
- ✅ Real-time collaboration
- ✅ Automatic sharing via links
- ✅ Version history

**Cons:**
- ❌ Requires service account setup
- ❌ More complex API
- ❌ Requires OAuth flow
- ❌ Needs Google Workspace

## Recommendation

**Start with .docx** for this use case because:
1. Simpler to implement and test
2. Sales reps can download and edit locally
3. Can be uploaded to Google Drive manually if needed
4. More flexible - works with any office suite
5. Your client likely already has Word/Google Docs access

## Testing the Generated Document

1. Download `quote_example.docx`
2. Open in Microsoft Word, Google Docs, or LibreOffice
3. Try editing:
   - Change a price in the product table
   - Modify the customer name
   - Add a new disclaimer paragraph
   - Adjust the discount percentage
4. Save and confirm all changes persist

## Validation Complete ✓

This POC confirms that **generating editable .docx quotes is fully viable** for your use case. Sales reps will be able to:
- Modify prices for special customers
- Add custom disclaimers
- Remove unavailable products
- Add manual product entries with images
- Make quick adjustments without recreating documents

---

**Questions or need modifications?** Let me know!
