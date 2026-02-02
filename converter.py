import os
import sys
import email
import tempfile
import mimetypes
import subprocess
import shutil
import base64
import base64
from email import policy
from email.message import EmailMessage
from pathlib import Path
from PIL import Image
from bs4 import BeautifulSoup
from pypdf import PdfWriter, PdfReader, Transformation, PageObject, PaperSize
import filetype
from loguru import logger
import argparse
from pdf2image import convert_from_path
import re
import unicodedata
import xml.etree.ElementTree as ET
from xml.dom import minidom
import zipfile
import extract_msg

# Constants
A4_WIDTH_MM = 210
A4_HEIGHT_MM = 297
DPI = 200 # User requested DPI (affects initial rasterization if needed, mostly for image scaling)
# A4 in points (1/72 inch) - PDF standard
A4_WIDTH_PT = 595.28
A4_HEIGHT_PT = 841.89

# Configure logger to write to file
logger.add("conversion.log", rotation="100 MB", level="INFO")

# macOS / Homebrew fix for WeasyPrint dependencies
if sys.platform == 'darwin':
    import ctypes.util
    # Common paths for Homebrew libraries
    extra_paths = ['/opt/homebrew/lib', '/usr/local/lib']
    current_path = os.environ.get('DYLD_FALLBACK_LIBRARY_PATH', '')
    
    # Update environment variable for subprocesses or late binding
    new_paths = []
    for p in extra_paths:
        if os.path.exists(p):
            new_paths.append(p)
    
    if new_paths:
        if current_path:
            os.environ['DYLD_FALLBACK_LIBRARY_PATH'] = f"{current_path}:{':'.join(new_paths)}"
        else:
            os.environ['DYLD_FALLBACK_LIBRARY_PATH'] = ':'.join(new_paths)

elif sys.platform == 'win32':
    # Windows specific setup for WeasyPrint (GTK3)
    # Check for common GTK3 installation paths if not in PATH
    
    # Detect Python Architecture (32 vs 64 bit)
    is_64bits = sys.maxsize > 2**32
    
    gtk3_paths = []
    if is_64bits:
        gtk3_paths = [
            r'C:\Program Files\GTK3-Runtime Win64\bin',
            r'C:\Program Files (x86)\GTK3-Runtime Win64\bin'            
        ]
    else:
        # 32-bit Python needs 32-bit GTK
        gtk3_paths = [
            r'C:\Program Files (x86)\GTK3-Runtime Win32\bin',
            r'C:\Program Files\GTK3-Runtime Win32\bin'
        ]

    current_path = os.environ.get('PATH', '')
    for p in gtk3_paths:
        if os.path.exists(p) and p not in current_path:
            os.environ['PATH'] = f"{p};{current_path}"
            logger.info(f"Added GTK3 path: {p}")

# Now import WeasyPrint which relies on gobject/pango/cairo
# MUST BE DONE AFTER PATH SETUP
try:
    from weasyprint import HTML, CSS
except OSError as e:
    logger.error(f"Failed to load WeasyPrint dependencies: {e}")
    if sys.platform == 'win32':
        logger.error("Please ensure GTK3 Runtime is installed and matches Python architecture (32/64 bit).")
    sys.exit(1)



def get_soffice_command():
    # Common paths for LibreOffice
    paths = [
        'soffice' # If in PATH
    ]
    
    if sys.platform == 'darwin':
        paths.append('/Applications/LibreOffice.app/Contents/MacOS/soffice')
    elif sys.platform == 'win32':
        paths.extend([
            r'C:\Program Files\LibreOffice\program\soffice.exe',
            r'C:\Program Files (x86)\LibreOffice\program\soffice.exe'
        ])

    for p in paths:
        if shutil.which(p) or os.path.exists(p):
            return p
    return None

def scale_to_a4(pdf_path, output_path):
    """
    Resizes all pages in the PDF to A4 size.
    """
    reader = PdfReader(pdf_path)
    writer = PdfWriter()

    for page in reader.pages:
        # Create a new blank A4 page
        new_page = PageObject.create_blank_page(width=A4_WIDTH_PT, height=A4_HEIGHT_PT)
        
        # Get original dimensions
        orig_width = page.mediabox.width
        orig_height = page.mediabox.height
        
        # Calculate scale factor to fit
        scale_w = A4_WIDTH_PT / orig_width
        scale_h = A4_HEIGHT_PT / orig_height
        scale = min(scale_w, scale_h)
        
        # Center the content
        tx = (A4_WIDTH_PT - orig_width * scale) / 2
        ty = (A4_HEIGHT_PT - orig_height * scale) / 2
        
        op = Transformation().scale(scale).translate(tx, ty)
        page.add_transformation(op)
        
        # Merge onto the blank A4 page (effectively resizing/centering)
        new_page.merge_page(page)
        writer.add_page(new_page)

    with open(output_path, 'wb') as f:
        writer.write(f)

def custom_url_fetcher(url):
    import ssl
    import urllib.request
    from weasyprint import default_url_fetcher
    
    # Handle file:// and data:// schemes with default fetcher
    if url.startswith("file:") or url.startswith("data:"):
        return default_url_fetcher(url)
        
    # Custom fetching for http/https to set User-Agent and ignore SSL errors
    logger.info(f"Fetching URL: {url}")
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    
    req = urllib.request.Request(url)
    req.add_header('User-Agent', 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.114 Safari/537.36')
    
    try:
        response = urllib.request.urlopen(req, context=ctx, timeout=10)
        content_type = response.info().get_content_type()
        data = response.read()
        logger.info(f"Fetched {len(data)} bytes, type: {content_type}")
        return {
            'mime_type': content_type,
            'string': data,
            'redirected_url': response.geturl()
        }
    except Exception as e:
        logger.warning(f"Failed to fetch resource {url}: {e}")
        # Fallback to default which might fail but handles some edge cases
        try:
            return default_url_fetcher(url)
        except:
            return None

def convert_html_to_pdf(html_content, output_path, base_url=None):
    """
    Converts HTML string to PDF using WeasyPrint.
    Ensures A4 size.
    """
    css = CSS(string=f'@page {{ size: A4; margin: 20mm; }} body {{ font-family: sans-serif; }} img {{ max-width: 100%; height: auto; }}')
    HTML(string=html_content, base_url=base_url).write_pdf(output_path, stylesheets=[css], resolution=DPI)

def convert_image_to_pdf(image_path, output_path):
    """
    Converts an image to a single page PDF (A4).
    """
    img = Image.open(image_path)
    
    # Calculate target size in pixels for the PDF @ 72 DPI (PDF Unit)
    # But wait, Pillow save pdf does 1px = 1pt usually? 
    # Let's simple use Pillow to save as PDF, then scale_to_a4 will handle the strict sizing.
    if img.mode == 'RGBA':
        img = img.convert('RGB')
    
    img.save(output_path, "PDF", resolution=DPI)
    
    # We rely on scale_to_a4 later to ensure it is exactly A4
    pass

def convert_office_to_pdf(input_path, output_path):
    """
    Converts Office files to PDF using LibreOffice (soffice).
    """
    soffice = get_soffice_command()
    if not soffice:
        logger.error("LibreOffice (soffice) not found. Cannot convert office document.")
        return False
    
    # LibreOffice conversion places the file in the outdir with the same name .pdf
    # We need a temp dir
    with tempfile.TemporaryDirectory() as temp_dir:
        cmd = [
            soffice, 
            '--headless', 
            '--convert-to', 'pdf', 
            '--outdir', temp_dir,
            input_path
        ]
        try:
            subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            # Find the result
            base_name = os.path.splitext(os.path.basename(input_path))[0]
            resulting_pdf = os.path.join(temp_dir, f"{base_name}.pdf")
            
            if os.path.exists(resulting_pdf):
                shutil.move(resulting_pdf, output_path)
                return True
            else:
                raise RuntimeError("LibreOffice conversion failed to produce a PDF output file.")
        except subprocess.CalledProcessError as e:
            # Re-raise with context so we can log it properly upstream
            raise RuntimeError(f"LibreOffice command failed: {e.stderr.decode('utf-8') if e.stderr else str(e)}")

def add_page_numbers(input_pdf_path, output_pdf_path):
    """
    Adds page numbers (Seite: X) to the bottom right of every page.
    """
    reader = PdfReader(input_pdf_path)
    writer = PdfWriter()
    
    total_pages = len(reader.pages)
    
    for i, page in enumerate(reader.pages):
        page_num = i + 1
        
        # Create a transient PDF with the page number using WeasyPrint
        # We use a transparent body so only the text shows
        html_content = f"""
        <html>
        <head>
            <style>
                @page {{ size: A4; margin: 0; }}
                body {{ margin: 0; padding: 0; }}
                .page-number {{
                    position: absolute;
                    bottom: 15mm;
                    right: 15mm;
                    font-family: sans-serif;
                    font-size: 12px;
                    color: #000;
                    padding: 2px 5px;
                }}
            </style>
        </head>
        <body>
            <div class="page-number">Seite: {page_num}</div>
        </body>
        </html>
        """
        
        # Generate watermark PDF in memory
        from io import BytesIO
        watermark_pdf = BytesIO()
        HTML(string=html_content).write_pdf(watermark_pdf)
        watermark_pdf.seek(0)
        
        # Merge watermark
        watermark_reader = PdfReader(watermark_pdf)
        watermark_page = watermark_reader.pages[0]
        
        page.merge_page(watermark_page)
        writer.add_page(page)
        
    with open(output_pdf_path, 'wb') as f:
        writer.write(f)

def extract_pdf_data(pdf_path):
    """
    Extracts metadata, text, and form fields from a PDF.
    Returns a dictionary with 'metadata' (dict), 'pages' (list of dicts), and 'form_fields' (list of dicts).
    """
    try:
        reader = PdfReader(pdf_path)
        data = {
            'metadata': {},
            'pages': [],
            'form_fields': []
        }
        
        # Metadata
        if reader.metadata:
            for key, value in reader.metadata.items():
                clean_key = key.replace('/', '')
                data['metadata'][clean_key] = str(value) if value else ""
                
        # Form Fields
        try:
            fields = reader.get_fields()
            if fields:
                for key, value in fields.items():
                    # Value is usually a dictionary or Field object
                    # We want the actual value ('/V')
                    field_val = None
                    if isinstance(value, dict):
                        field_val = value.get('/V')
                    
                    # Store as name/value pari
                    data['form_fields'].append({
                        'name': key,
                        'value': str(field_val) if field_val is not None else "" 
                    })
        except Exception as e:
            logger.warning(f"Failed to extract form fields from {pdf_path}: {e}")
                
        # Content
        for i, page in enumerate(reader.pages):
            page_text = ""
            try:
                page_text = page.extract_text()
            except Exception as e:
                logger.warning(f"Failed to extract text from page {i+1} of {pdf_path}: {e}")
                page_text = "[Error extracting text]"
            
            data['pages'].append({
                'id': i + 1,
                'text': page_text
            })
            
        return data
    except Exception as e:
        logger.error(f"Failed to extract info from PDF {pdf_path}: {e}")
        return None

def save_msg_as_eml(msg_path, output_eml_path):
    """
    Manually converts an MSG file (extract-msg object) to an EML file.
    """
    try:
        msg = extract_msg.Message(str(msg_path))
        
        email_msg = EmailMessage()
        
        # Copy Headers
        # Map common headers. extract-msg properties usually match closely or we access header dict.
        # Safer to use properties where available.
        if msg.subject: email_msg['Subject'] = msg.subject
        if msg.sender: email_msg['From'] = msg.sender
        if msg.to: email_msg['To'] = msg.to
        if msg.cc: email_msg['Cc'] = msg.cc
        if msg.bcc: email_msg['Bcc'] = msg.bcc
        if msg.date: email_msg['Date'] = msg.date
        
        # Determine Body
        # Prefer HTML, fallback to Text
        body_text = msg.body
        body_html = None
        try:
            body_html = msg.htmlBody
        except:
             pass # Some MSGs don't have HTML body or it fails
        
        if body_html:
            # Ensure body_html is string
            if isinstance(body_html, bytes):
                try:
                    body_html = body_html.decode('utf-8')
                except UnicodeDecodeError:
                    try:
                        body_html = body_html.decode('latin-1')
                    except:
                        logger.warning("Could not decode HTML body of MSG")
                        body_html = None
            
            if body_html:
                if body_text:
                    email_msg.set_content(body_text)
                    email_msg.add_alternative(body_html, subtype='html')
                else:
                    email_msg.set_content(body_html, subtype='html')
        else:
            if body_text:
                email_msg.set_content(body_text)
            else:
                email_msg.set_content("(No Body)")

        # Handle Attachments
        for attachment in msg.attachments:
            # extract-msg attachments have .data (bytes) and properties for filename
            try:
                data = attachment.data
                if not data: continue
                
                long_filename = getattr(attachment, 'longFilename', None)
                short_filename = getattr(attachment, 'shortFilename', None)
                filename = long_filename or short_filename or "untitled"
                
                # Guess mime type
                # attachment might have mimetype property? Not reliable.
                mime_type = filetype.guess_mime(data)
                if not mime_type:
                     mime_type, _ = mimetypes.guess_type(filename)
                
                if not mime_type:
                    mime_type = 'application/octet-stream'
                
                maintype, subtype = mime_type.split('/', 1)
                
                email_msg.add_attachment(
                    data,
                    maintype=maintype,
                    subtype=subtype,
                    filename=filename
                )
            except Exception as e:
                logger.warning(f"Skipping MSG attachment {attachment}: {e}")

        # Save to file
        with open(output_eml_path, 'wb') as f:
            f.write(email_msg.as_bytes())
            
        msg.close()
        return True

    except Exception as e:
        logger.error(f"Failed manual MSG->EML conversion: {e}")
        return False

# --- REFACTORED ATTACHMENT PROCESSING ---

def convert_attachment(filepath, output_pdf_path):
    """
    Converts a single file (attachment) to PDF.
    Handles: PDF, Images, Office, EML, MSG, ZIP.
    Returns True on success, False otherwise.
    """
    filepath = Path(filepath)
    output_pdf_path = Path(output_pdf_path)
    filename = filepath.name
    
    # Simple check for empty files
    if filepath.stat().st_size == 0:
        logger.warning(f"Skipping empty file: {filepath}")
        return False

    # Detect file type based on headers (Magic Bytes)
    kind = filetype.guess(str(filepath))
    detected_mime = kind.mime if kind else None
    ext = filepath.suffix.lower()

    # Fallback if filetype fails
    if not detected_mime:
        detected_mime, _ = mimetypes.guess_type(filepath)
    
    logger.info(f"Converting attachment {filename}: MIME={detected_mime}, Ext={ext}")

    try:
        # 1. PDF
        if detected_mime == 'application/pdf':
            shutil.copy(filepath, output_pdf_path)
            # Normalize to A4 will happen in caller or require explicit call? 
            # The caller `process_eml` loops result and normalizing. 
            # Let's ensure this function produces a valid PDF at `output_pdf_path`.
            return True

        # 2. Images
        elif detected_mime and detected_mime.startswith('image/'):
            if detected_mime in ['image/jpeg', 'image/png', 'image/gif', 'image/bmp', 'image/tiff', 'image/webp']:
                convert_image_to_pdf(filepath, output_pdf_path)
                return True
        
        # 3. Office Documents
        elif detected_mime in [
            'application/msword',
            'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            'application/vnd.ms-excel',
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'application/vnd.ms-powerpoint',
            'application/vnd.openxmlformats-officedocument.presentationml.presentation',
            'application/vnd.oasis.opendocument.text',
            'application/vnd.oasis.opendocument.spreadsheet',
            'application/vnd.oasis.opendocument.presentation'
        ] or ext in ['.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx', '.odt', '.ods', '.odp']:
            
            is_office = False
            # MIME check
            if detected_mime in [
                'application/msword',
                'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                'application/vnd.ms-excel',
                'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                'application/vnd.ms-powerpoint',
                'application/vnd.openxmlformats-officedocument.presentationml.presentation',
                'application/vnd.oasis.opendocument.text',
                'application/vnd.oasis.opendocument.spreadsheet',
                'application/vnd.oasis.opendocument.presentation'
            ] and ext != '.msg':
                is_office = True
            # Extension/Zip fallback
            elif detected_mime == 'application/zip' and ext in ['.docx', '.xlsx', '.pptx', '.odt', '.ods', '.odp']:
                 is_office = True
            elif ext in ['.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx', '.odt', '.ods', '.odp']:
                 is_office = True
            
            if is_office:
                return convert_office_to_pdf(str(filepath), str(output_pdf_path))

        # 4. EML Files (Nested)
        # filetype might identify as 'message/rfc822' or similar, or just text.
        # Strict check on extension or content?
        # filetype doesn't always detect EML well (it's text).
        if ext == '.eml' or detected_mime == 'message/rfc822':
             # Recursively process this EML
             logger.info(f"Recursively processing EML: {filename}")
             # process_eml creates a PDF from an EML.
             process_eml(str(filepath), str(output_pdf_path))
             return True

        # 5. MSG Files (Outlook)
        if ext == '.msg' or detected_mime == 'application/vnd.ms-outlook':
             logger.info(f"Converting MSG to EML then processing: {filename}")
             try:
                 with tempfile.TemporaryDirectory() as temp_msg_dir:
                     temp_eml_path = Path(temp_msg_dir) / f"{filepath.stem}.eml"
                     
                     if save_msg_as_eml(filepath, temp_eml_path):
                         # Recursively process the generated EML
                         process_eml(str(temp_eml_path), str(output_pdf_path))
                         return True
                     else:
                         return False
             except Exception as e:
                 logger.error(f"Failed to convert MSG {filename}: {e}")
                 return False

        # 6. ZIP / Archives
        if detected_mime == 'application/zip' or ext == '.zip':
             logger.info(f"Processing ZIP archive: {filename}")
             # We need to extract, convert all contents, and merge them into one PDF (output_pdf_path)
             with tempfile.TemporaryDirectory() as zip_temp_dir:
                 zip_path = Path(zip_temp_dir)
                 try:
                     with zipfile.ZipFile(filepath, 'r') as zf:
                         zf.extractall(zip_path)
                 except Exception as e:
                     logger.error(f"Failed to extract ZIP {filename}: {e}")
                     return False
                 
                 # Collect all valid PDFs
                 generated_pdfs = []
                 
                 # Walk strictly to find files
                 # We sort to have deterministic order
                 file_list = []
                 for root, dirs, files in os.walk(zip_path):
                     for f in files:
                         file_list.append(Path(root) / f)
                 file_list.sort(key=lambda x: str(x)) # Sort by path
                 
                 for i, subfile in enumerate(file_list):
                     # Skip __MACOSX and hidden files
                     if '__MACOSX' in subfile.parts or subfile.name.startswith('.'):
                         continue
                     
                     sub_output_pdf = zip_path / f"zip_part_{i:04d}.pdf"
                     
                     # Recursion!
                     if convert_attachment(subfile, sub_output_pdf):
                         # Ensure it is normalized to A4
                         sub_output_a4 = zip_path / f"zip_part_{i:04d}_a4.pdf"
                         scale_to_a4(str(sub_output_pdf), str(sub_output_a4))
                         generated_pdfs.append(str(sub_output_a4))
                 
                 if generated_pdfs:
                     # Merge
                     merger = PdfWriter()
                     for pdf in generated_pdfs:
                         merger.append(pdf)
                     merger.write(output_pdf_path)
                     return True
                 else:
                     logger.warning(f"ZIP {filename} contained no convertible files.")
                     return False

    except Exception as e:
        logger.error(f"Error converting {filename}: {e}")
        # Traceback could be helpful for debugging
        # import traceback
        # logger.error(traceback.format_exc())
        return False
    
    return False

def process_eml(eml_path, output_pdf_path):
    with open(eml_path, 'rb') as f:
        msg = email.message_from_binary_file(f, policy=policy.default)

    temp_dir = Path(tempfile.mkdtemp())
    pdf_parts = []
    
    # List to store analysis data for XML output
    analysis_data = {
        'source': Path(eml_path).name,
        'attachments': []
    }
    
    try:
        # 0. Extract Metadata
        subject = msg.get('Subject', '')
        from_ = msg.get('From', '')
        to_ = msg.get('To', '')
        cc_ = msg.get('Cc', '')
        bcc_ = msg.get('Bcc', '') # Often None
        date_ = msg.get('Date', '')
        
        analysis_data['subject'] = subject
        
        attachment_names = []
        for part in msg.iter_attachments():
            fn = part.get_filename()
            if fn:
                attachment_names.append(fn)
        
        # Construct Metadata HTML Block
        meta_html = f"""
        <div style="font-family: sans-serif; font-size: 14px; border-bottom: 2px solid #000; padding-bottom: 20px; margin-bottom: 30px;">
            <p style="margin: 2px 0;"><b>Von:</b> {from_}</p>
            <p style="margin: 2px 0;"><b>An:</b> {to_}</p>
        """
        if cc_:
            meta_html += f'<p style="margin: 2px 0;"><b>CC:</b> {cc_}</p>'
        if bcc_:
            meta_html += f'<p style="margin: 2px 0;"><b>BCC:</b> {bcc_}</p>'
        
        meta_html += f"""
            <p style="margin: 2px 0;"><b>Datum:</b> {date_}</p>
            <p style="margin: 2px 0;"><b>Betreff:</b> {subject}</p>
            <p style="margin: 2px 0;"><b>Anhänge:</b> {", ".join(attachment_names) if attachment_names else "Keine"}</p>
        </div>
        """

        # 1. Extract Body
        body_content = ""
        html_part = msg.get_body(preferencelist=('html'))
        if html_part:
            body_content = html_part.get_content()
            
            # Sanitization and Fixes
            soup = BeautifulSoup(body_content, 'html.parser')
            
            # Inject Metadata at top of body
            if soup.body:
                # Parse meta_html into a tag
                from bs4 import BeautifulSoup as BS
                meta_soup = BS(meta_html, 'html.parser')
                # Insert at beginning
                soup.body.insert(0, meta_soup)
            else:
                # If no body (weird), wrap it
                body_content = f"<html><body>{meta_html}{body_content}</body></html>"
                soup = BeautifulSoup(body_content, 'html.parser')
            
            # 1. Aggressive CSS Cleanup
            # Remove style tags that contain problematic rules causing WeasyPrint rendering issues
            # (e.g., hidden mix-blend-modes or mobile-specific max-heights for product images)
            for style in soup.find_all('style'):
                if style.string and ('mix-blend-mode' in style.string or '.productImage' in style.string):
                    style.decompose()

            # 2. Embed images as Base64 and normalize styles
            import ssl
            import urllib.request
            
            ctx = ssl.create_default_context()
            ctx.check_hostname = False
            ctx.verify_mode = ssl.CERT_NONE
            opener = urllib.request.build_opener(urllib.request.HTTPSHandler(context=ctx))
            opener.addheaders = [('User-Agent', 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.114 Safari/537.36')]
            urllib.request.install_opener(opener)

            for img_tag in soup.find_all('img'):
                # 2a. Embed src if present
                src = img_tag.get('src')
                if src and src.startswith('http'):
                    try:
                        logger.info(f"Downloading and embedding: {src}")
                        with urllib.request.urlopen(src, timeout=10) as response:
                            content_type = response.info().get_content_type()
                            data = response.read()
                            b64_data = base64.b64encode(data).decode('utf-8')
                            img_tag['src'] = f"data:{content_type};base64,{b64_data}"
                    except Exception as e:
                        logger.warning(f"Failed to embed image {src}: {e}")
                
                # 2b. Normalize styles based on context
                current_style = img_tag.get('style', '')
                classes = img_tag.get('class', [])
                
                # Amazon Product Image Fix: FORCE reset of style inline
                # We completely overwrite to remove potential restrictive max-widths/heights present in the email HTML
                if 'productImage' in classes:
                     logger.info(f"Fixing Amazon Product Image: {src}")
                     # Remove HTML attributes that might conflict
                     if img_tag.has_attr('width'): del img_tag['width']
                     if img_tag.has_attr('height'): del img_tag['height']
                     
                     # Force robust sizing for responsive layout without overflow
                     # width: auto + max-width: 100% allows intrinsic size but prevents overflow
                     img_tag['style'] = (
                         "display: block !important; "
                         "width: auto !important; "
                         "max-width: 100% !important; "
                         "height: auto !important; "
                         "box-sizing: border-box !important; "
                         "margin: 0 auto !important;"
                     )
                
                # General Fix (Steam etc): Needs display:block to escape line-height:0 containers
                elif 'display' not in current_style:
                     img_tag['style'] = f"{current_style}; display: block;"

            # 3. Fix bgcolor for WeasyPrint (convert legacy attribute to CSS)
            for tag in soup.find_all(attrs={"bgcolor": True}):
                bg_color = tag['bgcolor']
                current_style = tag.get('style', '')
                if 'background-color' not in current_style:
                    tag['style'] = f"{current_style}; background-color: {bg_color};"
            
            # 4. Inject global styles for rendering consistency (colors, visibility)
            override_style = soup.new_tag('style', type='text/css')
            override_style.string = """
                img {
                    visibility: visible !important;
                    opacity: 1 !important;
                    mix-blend-mode: normal !important;
                }
                /* Specific fix for Amazon product images */
                .productImage {
                    display: block !important;
                    width: 100% !important;
                    height: auto !important;
                    max-width: none !important;
                }
                /* Ensure background colors print */
                * {
                    -webkit-print-color-adjust: exact !important;
                    print-color-adjust: exact !important;
                }
            """
            if soup.head:
                soup.head.append(override_style)
            else:
                if not soup.body:
                    soup.append(soup.new_tag('body'))
                soup.body.insert(0, override_style)

            body_content = str(soup)
        else:
            text_part = msg.get_body(preferencelist=('plain'))
            # For plain text, we create a basic HTML wrapper
            # We inject the metadata block here too
            body_content = f"<html><body>{meta_html}<pre style='font-family: monospace; white-space: pre-wrap;'>{text_part.get_content() if text_part else 'No body content.'}</pre></body></html>"

        # 2. Extract Inline Images (CIDs) - already handled by existing code, but let's check
        # Existing CID code (below in original file) replaces src with file:// paths.
        # We need to make sure we don't break that.
        # The existing code is:
        # for part in msg.walk(): ... if cid: ... soup.find_all(...) ... img_tag['src'] = file://...
        # Since we modified soup above and regenerated body_content, the CIDs are still in there as "cid:..." strings.
        # We need to re-soup the body_content or move this logic up.
        # actually, the original code does steps sequentially.
        # Original Step 2: Extract Inline Images. It parses `body_content` AGAIN into a NEW soup.
        # So my changes to `body_content` (Base64 embedding) will persist, and the next block will parse it again to handle CIDs.
        # That is fine.
        
        # 3. Create Body PDF
        body_pdf_path = temp_dir / "00_body.pdf"
        # Use simple convert without custom fetcher since we embedded everything
        convert_html_to_pdf(body_content, str(body_pdf_path), base_url=str(temp_dir))
        pdf_parts.append(str(body_pdf_path))

        # 4. Attachments
        attach_idx = 1
        for part in msg.iter_attachments():
            filename = part.get_filename()
            if not filename:
                continue 
            
            # --- FILENAME SANITIZATION START ---
            # 1. Normalize Unicode (NFKD) and drop non-ASCII (removes UTF-8 chars like Umlaute -> base letter)
            filename = unicodedata.normalize('NFKD', filename).encode('ascii', 'ignore').decode('ascii')
            
            # 2. Remove spaces and remaining non-alphanumeric chars (keep dot, underscore, dash)
            # User requirement: "keine spaces" -> replace with underscore for readability
            filename = filename.replace(' ', '_')
            filename = re.sub(r'[^a-zA-Z0-9._-]', '', filename)
            
            # 3. Truncate to max 50 chars while preserving extension
            MAX_LEN = 50
            if len(filename) > MAX_LEN:
                stem, suffix = os.path.splitext(filename)
                # Ensure we have at least 1 char for stem
                limit = MAX_LEN - len(suffix)
                if limit < 1: 
                    limit = 1
                    # If extension itself is super long, we might still exceed 50, but we must protect extension.
                    # Or we hard truncate everything.
                    # User: "variable filename maximal 50 Zeichen gross sein". 
                    # Strict interpretation: trunc whole string.
                    # But cutting extension breaks file type detection later (suffix check).
                    # I will try to preserve suffix if possible.
                
                stem = stem[:limit]
                filename = stem + suffix
                
                # Final safety check if suffix was huge
                if len(filename) > MAX_LEN:
                     filename = filename[:MAX_LEN]
            
            # --- FILENAME SANITIZATION END ---
             
            
            filepath = temp_dir / filename
            try:    
                if part.get_content_type() == 'message/rfc822':
                     with open(filepath, 'wb') as att_f:
                        att_f.write(part.as_bytes())
                else:
                    with open(filepath, 'wb') as att_f:
                        att_f.write(part.get_content())
            except Exception as e:
                logger.error(f"Failed to write attachment {filename}: {str(e)}")
                continue
            
            output_part_path = temp_dir / f"{attach_idx:02d}_{filename}.pdf"
            
            success = False
            
            # Use universal converter
            try:
                success = convert_attachment(filepath, output_part_path)
                
                # If conversion was successful and resulted in a PDF, extract data if it's a PDF
                if success and os.path.exists(output_part_path):
                    # Check if the original attachment was a PDF or an Office file that got converted
                    # This logic was previously inside the attachment conversion block, now it's here.
                    # We need to re-detect the type of the *original* attachment to decide if we should extract data.
                    # For now, let's assume if it successfully converted to PDF, we can try to extract data.
                    # This might be too broad, but matches the spirit of "analyze PDF files that are processed".
                    
                    # Re-detect type of original attachment for analysis decision
                    original_kind = filetype.guess(str(filepath))
                    original_detected_mime = original_kind.mime if original_kind else None
                    original_ext = filepath.suffix.lower()

                    is_original_pdf_or_office = False
                    if original_detected_mime == 'application/pdf':
                        is_original_pdf_or_office = True
                    elif original_detected_mime in [
                        'application/msword', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                        'application/vnd.ms-excel', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                        'application/vnd.ms-powerpoint', 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
                        'application/vnd.oasis.opendocument.text', 'application/vnd.oasis.opendocument.spreadsheet',
                        'application/vnd.oasis.opendocument.presentation'
                    ] or original_ext in ['.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx', '.odt', '.ods', '.odp']:
                        is_original_pdf_or_office = True
                    
                    if is_original_pdf_or_office:
                        logger.info(f"Extracting XML data for: {filename}")
                        pdf_data = extract_pdf_data(str(output_part_path)) # Extract from the *converted* PDF
                        if pdf_data:
                            pdf_data['filename'] = filename
                            analysis_data['attachments'].append(pdf_data)
                
            except Exception as e:
                # DETAILED LOGGING as requested
                logger.error(f"FAILED to convert attachment. EML='{Path(eml_path).name}' | Attachment='{filename}' | Error='{e}'")
                success = False
            # --- CONVERSION LOGIC END ---
            
            
            
            if success and os.path.exists(output_part_path):
                # Normalize to A4 immediately
                normalized_path = str(output_part_path).replace('.pdf', '_a4.pdf')
                scale_to_a4(str(output_part_path), normalized_path)
                pdf_parts.append(normalized_path)
            else:
                logger.warning(f"Skipping attachment {filename}: Unsupported format or conversion failed.")
            
            attach_idx += 1

        # 5. Merge
        # Also need to make sure the Body PDF is A4
        body_a4 = str(body_pdf_path).replace('.pdf', '_a4.pdf')
        scale_to_a4(str(body_pdf_path), body_a4)
        pdf_parts[0] = body_a4

        merger = PdfWriter()
        for pdf in pdf_parts:
            merger.append(pdf)
        
        # Write to temporary file first
        temp_merged_pdf = temp_dir / "temp_merged.pdf"
        merger.write(str(temp_merged_pdf))
        
        # 6. Add Page Numbers
        add_page_numbers(str(temp_merged_pdf), output_pdf_path)
        logger.info(f"Successfully created PDF at {output_pdf_path}")
        
        # 7. Write Analysis XML
        try:
            root = ET.Element("EmailAnalysis")
            
            source_node = ET.SubElement(root, "Source")
            source_node.text = analysis_data.get('source', '')
            
            subject_node = ET.SubElement(root, "Subject")
            subject_node.text = analysis_data.get('subject', '')
            
            atts_node = ET.SubElement(root, "Attachments")
            
            for att_data in analysis_data['attachments']:
                att_node = ET.SubElement(atts_node, "Attachment", filename=att_data.get('filename', ''))
                
                # Metadata
                meta_node = ET.SubElement(att_node, "Metadata")
                for k, v in att_data.get('metadata', {}).items():
                    # Sanitize tag name (remove spaces, etc to be valid XML tag if needed, but best to use generic item with key attr or just children)
                    # Let's just use the key as tag if valid, otherwise Item
                    try:
                        node = ET.SubElement(meta_node, k.replace(' ', '_'))
                        node.text = v
                    except:
                        # Fallback
                        node = ET.SubElement(meta_node, "MetaItem", key=k)
                        node.text = v
                
                # Form Fields
                if att_data.get('form_fields'):
                    fields_node = ET.SubElement(att_node, "FormFields")
                    for field in att_data['form_fields']:
                        f_node = ET.SubElement(fields_node, "Field", name=field['name'])
                        f_node.text = field['value']

                # Content
                content_node = ET.SubElement(att_node, "Content")
                for page in att_data.get('pages', []):
                    p_node = ET.SubElement(content_node, "Page", id=str(page['id']))
                    p_node.text = page['text']
            
            # Save
            output_xml_path = Path(output_pdf_path).with_suffix('.xml')
            xml_str = minidom.parseString(ET.tostring(root)).toprettyxml(indent="    ")
            
            with open(output_xml_path, "w", encoding="utf-8") as f:
                f.write(xml_str)
            logger.info(f"Successfully created XML Analysis at {output_xml_path}")
            
        except Exception as e:
            logger.error(f"Failed to write XML analysis: {e}")

    finally:
        shutil.rmtree(temp_dir)

def convert_pdf_to_tiff(pdf_path, tiff_path):
    """
    Converts a PDF to a multipage TIFF with JPEG compression.
    """
    try:
        # Convert PDF pages to images
        # 200 DPI matches our generation DPI
        images = convert_from_path(pdf_path, dpi=DPI)
        
        if not images:
            logger.warning("No images extracted from PDF for TIFF conversion.")
            return

        # Save as Multipage TIFF
        # compression='jpeg' requires the original images to be compatible, usually RGB
        # We ensure RGB mode
        rgb_images = []
        for img in images:
            if img.mode != 'RGB':
                img = img.convert('RGB')
            rgb_images.append(img)
            
        rgb_images[0].save(
            tiff_path,
            compression="jpeg",
            save_all=True,
            append_images=rgb_images[1:]
        )
        logger.info(f"Successfully created TIFF at {tiff_path}")
        
    except Exception as e:
        logger.error(f"Failed to convert PDF to TIFF: {e}")
        # Hint about poppler if it looks like a missing dependency
        if "poppler" in str(e).lower():
            logger.error("Ensure 'poppler' is installed (brew install poppler OR apt-get install poppler-utils)")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Convert EML file to PDF (and optionally TIFF).")
    parser.add_argument("--input", "-i", required=True, help="Path to input .eml file or directory containing .eml files")
    parser.add_argument("--output", "-o", help="Path to output file (if input is file) or directory (if input is directory)")
    parser.add_argument("--format", "-f", choices=['pdf', 'tif'], default='pdf', 
                        help="Output format: 'pdf' (default) or 'tif' (generates both PDF and TIFF)")
    
    args = parser.parse_args()
    
    input_path = Path(args.input)
    
    if not input_path.exists():
        logger.error(f"Error: Input path {input_path} does not exist.")
        sys.exit(1)

    files_to_process = []
    if input_path.is_dir():
        # Batch mode: find all .eml files
        files_to_process = list(input_path.glob("*.eml"))
        if not files_to_process:
            logger.warning(f"No .eml files found in {input_path}")
    else:
        # Single file mode
        files_to_process = [input_path]

    # Determine output directory
    if args.output:
        output_arg = Path(args.output)
        # If input is dir, output should be dir (create if needed)
        if input_path.is_dir():
            if not output_arg.exists():
                output_arg.mkdir(parents=True, exist_ok=True)
            elif not output_arg.is_dir():
                 logger.error(f"Error: Input is a directory but output {output_arg} is a file.")
                 sys.exit(1)
            output_dir = output_arg
        else:
            # Input is file. 
            # Output can be a file path OR a directory.
            # We assume if it ends in .pdf it is a file path, otherwise a directory ?? 
            # Actually, to be safe: valid check.
            if output_arg.suffix.lower() == '.pdf':
                output_dir = output_arg.parent
                # Special case: direct filename provided for single file
                # We handle this inside loop? No, simpler to just set the target path.
            else:
                # Assume directory
                if not output_arg.exists():
                    output_arg.mkdir(parents=True, exist_ok=True)
                output_dir = output_arg
    else:
        # Default output to same directory as input file(s)
        if input_path.is_dir():
            output_dir = input_path
        else:
            output_dir = input_path.parent

    for eml_file in files_to_process:
        try:
            # Determine effective output PDF path
            base_name = eml_file.stem
            
            # If explicit output file path was given for single file input
            if not input_path.is_dir() and args.output and Path(args.output).suffix.lower() == '.pdf':
                target_pdf = Path(args.output)
            else:
                target_pdf = output_dir / f"{base_name}.pdf"
            
            logger.info(f"Processing {eml_file} -> {target_pdf}")
            process_eml(str(eml_file), str(target_pdf))
            
            # Generate TIFF if requested
            if args.format == 'tif':
                output_tiff = target_pdf.with_suffix('.tif')
                convert_pdf_to_tiff(str(target_pdf), str(output_tiff))
                
        except Exception as e:
            import traceback
            logger.error(f"Failed to process {eml_file}: {e}\n{traceback.format_exc()}")

