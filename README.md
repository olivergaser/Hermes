# EML to PDF Converter

This tool converts an EML file into a single PDF, including the email body, inline images, and attachments (Images, PDF, Office documents).
Every page is standardized to DIN A4.

## Prerequisites

### 1. Python Environment
Install the python requirements:
```bash
pip install -r requirements.txt
```

### 2. System Dependencies

**WeasyPrint** (for HTML conversion) requires system libraries:
On macOS (using Homebrew):
```bash
brew install python-tk pango libffi cairo
```

**LibreOffice** (for Office document conversion):
This script looks for `soffice` at `/Applications/LibreOffice.app/Contents/MacOS/soffice` or in your PATH.
- If you have LibreOffice installed in the Applications folder, it should work out of the box.
- Otherwise, install it from https://www.libreoffice.org/ or `brew install --cask libreoffice`.

## Usage

```bash
python converter.py input_email.eml output_file.pdf
```

## Features
- **Body**: Converted to PDF using WeasyPrint (HTML/CSS support).
- **Inline Images**: Embedded correctly in the body.
- **Attachments**:
  - Images: Converted to PDF pages.
  - PDFs: Appended as-is (resized to A4).
  - Office (Docx, Xlsx, etc): Converted via LibreOffice to PDF.
- **Output**: All pages scaled/centered to DIN A4 (210x297mm).
