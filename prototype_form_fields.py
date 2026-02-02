from pypdf import PdfReader
import os
import sys

def extract_fields(pdf_path):
    try:
        reader = PdfReader(pdf_path)
        fields = reader.get_fields()
        
        if fields:
            print(f"Found {len(fields)} fields.")
            for key, value in fields.items():
                # Value is a dict with '/V' for value, '/T' for field name etc, or usually the value is directly accessible or nested.
                # pypdf returns a dict/Field object. Let's inspect it.
                val = value.get('/V', 'N/A')
                print(f"Key: {key} | Value: {val}")
        else:
            print("No form fields found.")
            
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    pdf_path = os.path.abspath(r"Testdateien\Antrag.pdf")
    if os.path.exists(pdf_path):
        extract_fields(pdf_path)
    else:
        print("PDF not found")
