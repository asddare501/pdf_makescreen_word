PDF to Word Converter
A Python tool that converts PDF documents to Word format by rendering pages as high-quality images and arranging them in a structured layout.

Features
Page Selection: Automatically skips the first page and processes remaining pages

High-Quality Output: Configurable zoom factor (default 2.5x) for crisp image rendering

Optimized Layout: Two pages per row in landscape orientation for efficient viewing

Page Numbering: Each page includes its original page number for easy reference

Automatic Cleanup: Temporary image files are removed after conversion

Arabic Interface: User-friendly prompts in Arabic

Requirements
bash
pip install PyMuPDF python-docx
Dependencies
PyMuPDF (fitz): PDF rendering and page extraction

python-docx: Word document creation and formatting

Usage
Interactive Mode
Run the script and follow the prompts:

bash
python pdf_to_word_converter.py
You'll be asked for:

Path to the PDF file

Output Word filename (default: output.docx)

Zoom factor for image quality (default: 2.5)

Programmatic Usage
python
from pdf_to_word_converter import pdf_to_word_pymupdf

pdf_to_word_pymupdf(
    pdf_path="document.pdf",
    output_path="converted.docx",
    zoom=2.5
)
How It Works
Opens PDF: Validates and loads the PDF file using PyMuPDF

Skips First Page: Excludes the cover/first page from conversion

Renders Pages: Converts each remaining page to PNG images at specified zoom level

Creates Word Document: Generates a landscape-oriented Word document with minimal margins

Arranges Layout: Places two page images per row in a table structure

Adds Page Numbers: Labels each image with its original page number

Cleans Up: Removes temporary image files after successful conversion

Output Format
Page Orientation: Landscape (11" × 8.5")

Layout: 2 columns per page

Margins: 0.3" on sides/top, 0.1" on bottom

Image Height: 5.2"

Page Numbers: Centered below each image in gray text (9pt)

Error Handling
The script handles common issues:

Missing PDF files

Single-page documents

Corrupted pages (skips and continues)

Image addition failures

File save errors

Limitations
First page is always excluded from conversion

Requires at least 2 pages in the PDF

Text is not extracted; pages are converted to images

Output file size depends on PDF complexity and zoom factor

Use Cases
Creating printable compilations of PDF slides

Archiving documents with preserved visual layout

Preparing materials for review with page-by-page structure

Converting scanned documents to Word format

License
This project is provided as-is for personal and educational use.

Note: For text-editable Word documents, consider using OCR tools if the PDF contains selectable text.
