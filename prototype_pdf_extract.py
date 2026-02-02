from pypdf import PdfReader
import xml.etree.ElementTree as ET
from xml.dom import minidom
import sys
import os

def extract_to_xml(pdf_path, output_xml_path):
    try:
        reader = PdfReader(pdf_path)
        
        root = ET.Element("PdfAnalysis")
        
        # Metadata
        meta = reader.metadata
        meta_node = ET.SubElement(root, "Metadata")
        if meta:
            for key, value in meta.items():
                # Clean key (remove slash)
                clean_key = key.replace('/', '')
                item = ET.SubElement(meta_node, clean_key)
                item.text = str(value) if value else ""

        # Pages / Text
        content_node = ET.SubElement(root, "Content")
        for i, page in enumerate(reader.pages):
            page_node = ET.SubElement(content_node, "Page", id=str(i+1))
            text_node = ET.SubElement(page_node, "Text")
            try:
                text = page.extract_text()
                text_node.text = text
            except Exception as e:
                text_node.text = f"Error extracting text: {e}"

        # Pretty print
        xml_str = minidom.parseString(ET.tostring(root)).toprettyxml(indent="    ")
        
        with open(output_xml_path, "w", encoding="utf-8") as f:
            f.write(xml_str)
            
        print(f"Successfully extracted to {output_xml_path}")
        
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    # Create a dummy PDF for testing if none exists
    from pypdf import PdfWriter
    dummy_pdf = "test_extract.pdf"
    
    writer = PdfWriter()
    writer.add_blank_page(width=72, height=72)
    # We can't easily add text with pypdf pure python without fonts, but metadata will be there.
    # Actually pypdf is mostly for manipulation. 
    # For the sake of the prototype, I'll rely on the existing 'output_ods.pdf' or similar if it exists, 
    # or just create a blank one with metadata.
    
    writer.add_metadata({
        '/Title': 'Test PDF',
        '/Author': 'Antigravity'
    })
    
    with open(dummy_pdf, "wb") as f:
        writer.write(f)
        
    extract_to_xml(dummy_pdf, "test_extract.xml")
    
    # Read back to verify
    with open("test_extract.xml", "r") as f:
        print(f.read())
        
    if os.path.exists(dummy_pdf):
        os.remove(dummy_pdf)
    if os.path.exists("test_extract.xml"):
        os.remove("test_extract.xml")
