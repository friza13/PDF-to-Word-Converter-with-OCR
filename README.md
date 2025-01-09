PDF to Word Converter with OCR
This Python application converts PDF files into editable Word documents. It includes two conversion modes:
1. OCR Mode: Extracts text from scanned PDF images using Tesseract OCR.
2. Direct Text Mode: Extracts text directly from PDFs with embedded text.

Features:
- User-friendly GUI built with Tkinter.
- Supports batch conversion of PDF pages into a single Word document.
- Customizable output directory selection.
- OCR support for image-based PDFs using pytesseract.

Prerequisites:
- Python 3.x
- Dependencies: pdf2image, pytesseract, python-docx, PyPDF2.
- External tools: Tesseract-OCR (https://github.com/tesseract-ocr/tesseract) and Poppler (https://github.com/oschwartz10612/poppler-windows/releases).

How to Use:
1. Install the required libraries and tools.
2. Run the script to open the GUI.
3. Select a PDF file, output directory, and conversion mode.
4. Click "Convert" to generate the Word document.

Feel free to contribute and improve this project! 😊
