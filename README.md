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

**macOS** (using Homebrew):
```bash
brew install python-tk pango libffi cairo libxml2
brew install --cask libreoffice
```

**Windows**:
1. **GTK3 Runtime**: Download and install the [GTK3 Runtime for Windows](https://github.com/tschoonj/GTK-for-Windows-Runtime-Environment-Installer/releases).
   - Ensure the `bin` folder (e.g., `C:\Program Files\GTK3-Runtime Win64\bin`) is added to your PATH, or the script will try to find it in standard locations.
2. **LibreOffice**: Install [LibreOffice](https://www.libreoffice.org/download/download/).
   - The script will look for `soffice.exe` in `C:\Program Files\LibreOffice\program\` or `C:\Program Files (x86)\LibreOffice\program\`.

**Linux (Debian/Ubuntu)**:
```bash
sudo apt-get install python3-pip python3-cffi python3-brotli libpango-1.0-0 libpangoft2-1.0-0 libreoffice
```

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


