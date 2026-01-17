import os
import sys
import email
import tempfile
import mimetypes
import subprocess
import shutil
import base64
from email import policy
from pathlib import Path
from PIL import Image
from bs4 import BeautifulSoup

from pypdf import PdfWriter, PdfReader, Transformation, PageObject, PaperSize

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
    gtk3_paths = [
        r'C:\Program Files\GTK3-Runtime Win64\bin',
        r'C:\Program Files (x86)\GTK3-Runtime Win64\bin'
    ]
    current_path = os.environ.get('PATH', '')
    for p in gtk3_paths:
        if os.path.exists(p) and p not in current_path:
            os.environ['PATH'] = f"{p};{current_path}"

# Now import WeasyPrint which relies on gobject/pango/cairo
from weasyprint import HTML, CSS

# Constants
A4_WIDTH_MM = 210
A4_HEIGHT_MM = 297
DPI = 200 # User requested DPI (affects initial rasterization if needed, mostly for image scaling)
# A4 in points (1/72 inch) - PDF standard
A4_WIDTH_PT = 595.28
A4_HEIGHT_PT = 841.89

def setup_logger():
    import logging
    logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
    return logging.getLogger(__name__)

logger = setup_logger()

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
                logger.error("LibreOffice conversion failed to produce a PDF.")
                return False
        except subprocess.CalledProcessError as e:
            logger.error(f"LibreOffice failed: {e}")
            return False

def process_eml(eml_path, output_pdf_path):
    with open(eml_path, 'rb') as f:
        msg = email.message_from_binary_file(f, policy=policy.default)

    temp_dir = Path(tempfile.mkdtemp())
    pdf_parts = []
    
    try:
        # 1. Extract Body
        body_content = ""
        html_part = msg.get_body(preferencelist=('html'))
        if html_part:
            body_content = html_part.get_content()
            
            # Sanitization and Fixes
            soup = BeautifulSoup(body_content, 'html.parser')
            
            # 1. Remove mix-blend-mode
            if soup.find('style'):
                for style in soup.find_all('style'):
                    if style.string:
                        style.string = style.string.replace('mix-blend-mode:multiply', 'mix-blend-mode:normal')
                        style.string = style.string.replace('mix-blend-mode:initial', 'mix-blend-mode:normal')

            # 2. Embed images as Base64 to ensure WeasyPrint renders them
            import ssl
            import urllib.request
            
            ctx = ssl.create_default_context()
            ctx.check_hostname = False
            ctx.verify_mode = ssl.CERT_NONE
            opener = urllib.request.build_opener(urllib.request.HTTPSHandler(context=ctx))
            opener.addheaders = [('User-Agent', 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.114 Safari/537.36')]
            urllib.request.install_opener(opener)

            for img_tag in soup.find_all('img', src=True):
                src = img_tag['src']
                if src.startswith('http'):
                    try:
                        logger.info(f"Downloading and embedding: {src}")
                        with urllib.request.urlopen(src, timeout=10) as response:
                            content_type = response.info().get_content_type()
                            data = response.read()
                            b64_data = base64.b64encode(data).decode('utf-8')
                            img_tag['src'] = f"data:{content_type};base64,{b64_data}"
                    except Exception as e:
                        logger.warning(f"Failed to embed image {src}: {e}")
            
            # 3. Specific fix for Amazon product images in tables (often have 1px/0px sizing issues in PDF)
            # Find images that look like product images and ensure they have a size
            for img in soup.find_all('img', class_='productImage'):
                 img['style'] = "display: block; max-width: 100%; height: auto;"

            body_content = str(soup)
        else:
            text_part = msg.get_body(preferencelist=('plain'))
            if text_part:
                body_content = f"<html><body><pre>{text_part.get_content()}</pre></body></html>"
            else:
                body_content = "<html><body><p>No body content found.</p></body></html>"

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
                continue # Skip inline processing if already done? 
                # Actually attachments might be inline images too if they were not referenced by CID or if policy differs.
                # policy.default usually handles this separation.
            
            filepath = temp_dir / filename
            with open(filepath, 'wb') as att_f:
                att_f.write(part.get_content())
            
            ext = filepath.suffix.lower()
            output_part_path = temp_dir / f"{attach_idx:02d}_{filename}.pdf"
            
            success = False
            if ext in ['.pdf']:
                success = True # Already PDF, just verify/copy
                shutil.copy(filepath, output_part_path)
            elif ext in ['.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.gif']:
                try:
                    convert_image_to_pdf(filepath, output_part_path)
                    success = True
                except Exception as e:
                    logger.error(f"Failed to convert image {filename}: {e}")
            elif ext in ['.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx']:
                success = convert_office_to_pdf(str(filepath), str(output_part_path))
            
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
        
        merger.write(output_pdf_path)
        logger.info(f"Successfully created PDF at {output_pdf_path}")

    finally:
        shutil.rmtree(temp_dir)

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python converter.py <input.eml> <output.pdf>")
        sys.exit(1)
    
    input_eml = sys.argv[1]
    output_pdf = sys.argv[2]
    
    if not os.path.exists(input_eml):
        print(f"Error: Input file {input_eml} does not exist.")
        sys.exit(1)
        
    process_eml(input_eml, output_pdf)
