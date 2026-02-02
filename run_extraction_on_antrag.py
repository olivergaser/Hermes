import sys
import os
import xml.etree.ElementTree as ET
from xml.dom import minidom
from pypdf import PdfReader

# Ensure pypdf is available
try:
    from pypdf import PdfReader
except ImportError:
    print("Error: pypdf not installed.")
    sys.exit(1)

def extract_pdf_data(pdf_path, output_xml_path):
    print(f"Extracting from {pdf_path}...")
    try:
        reader = PdfReader(pdf_path)
        root = ET.Element("PdfAnalysis")
        
        # Metadata
        meta_node = ET.SubElement(root, "Metadata")
        if reader.metadata:
            for key, value in reader.metadata.items():
                clean_key = key.replace('/', '')
                item = ET.SubElement(meta_node, clean_key)
                item.text = str(value) if value else ""

        # Form Fields
        try:
            fields = reader.get_fields()
            if fields:
                fields_node = ET.SubElement(root, "FormFields")
                for key, value in fields.items():
                    field_val = None
                    if isinstance(value, dict):
                        field_val = value.get('/V')
                    
                    f_node = ET.SubElement(fields_node, "Field", name=key)
                    f_node.text = str(field_val) if field_val is not None else ""
        except Exception as e:
            print(f"Warning: Form field extraction failed: {e}")

        # Content
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
    pdf_path = os.path.abspath(r"Testdateien\Antrag.pdf")
    output_xml = "antrag_extraction.xml"
    
    if not os.path.exists(pdf_path):
        print(f"File not found: {pdf_path}")
        sys.exit(1)
        
    extract_pdf_data(pdf_path, output_xml)
