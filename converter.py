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

from loguru import logger

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

def process_eml(eml_path, output_pdf_path):
    with open(eml_path, 'rb') as f:
        msg = email.message_from_binary_file(f, policy=policy.default)

    temp_dir = Path(tempfile.mkdtemp())
    pdf_parts = []
    
    try:
        # 0. Extract Metadata
        subject = msg.get('Subject', '')
        from_ = msg.get('From', '')
        to_ = msg.get('To', '')
        cc_ = msg.get('Cc', '')
        bcc_ = msg.get('Bcc', '') # Often None
        date_ = msg.get('Date', '')
        
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
        
        # Write to temporary file first
        temp_merged_pdf = temp_dir / "temp_merged.pdf"
        merger.write(str(temp_merged_pdf))
        
        # 6. Add Page Numbers
        add_page_numbers(str(temp_merged_pdf), output_pdf_path)
        logger.info(f"Successfully created PDF at {output_pdf_path}")

    finally:
        shutil.rmtree(temp_dir)

import argparse
from pdf2image import convert_from_path

# ... imports ...

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
    parser.add_argument("input_eml", help="Path to input .eml file")
    parser.add_argument("output_pdf", help="Path to output .pdf file")
    parser.add_argument("--format", choices=['pdf', 'tif'], default='pdf', 
                        help="Output format: 'pdf' (default) or 'tif' (generates both PDF and TIFF)")
    
    args = parser.parse_args()
    
    if not os.path.exists(args.input_eml):
        print(f"Error: Input file {args.input_eml} does not exist.")
        sys.exit(1)
        
    # Generate PDF
    process_eml(args.input_eml, args.output_pdf)
    
    # Generate TIFF if requested
    if args.format == 'tif':
        # Derive TIFF filename from PDF filename
        base_name = os.path.splitext(args.output_pdf)[0]
        output_tiff = f"{base_name}.tif"
        convert_pdf_to_tiff(args.output_pdf, output_tiff)
