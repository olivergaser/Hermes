import unittest
from unittest.mock import MagicMock, patch
import os
import sys
import xml.etree.ElementTree as ET
from pypdf import PdfWriter

# Mock WeasyPrint before it is imported by converter
sys.modules['weasyprint'] = MagicMock()

# Add parent directory to path so we can import converter
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from converter import process_eml
import email
from email.message import EmailMessage

class TestPdfExtraction(unittest.TestCase):
    @patch('converter.convert_html_to_pdf')
    @patch('converter.add_page_numbers')
    @patch('converter.scale_to_a4')
    @patch('converter.convert_image_to_pdf')
    @patch('converter.convert_office_to_pdf')
    def test_xml_generation(self, mock_office, mock_img, mock_scale, mock_page_nums, mock_html):
        # Setup mocks to create dummy files so file existence checks pass
        def create_dummy_pdf(input_path, output_path, **kwargs):
            # Create a minimal valid PDF at output_path
            writer = PdfWriter()
            writer.add_blank_page(width=100, height=100)
            with open(output_path, 'wb') as f:
                writer.write(f)
        
        mock_html.side_effect = create_dummy_pdf
        mock_scale.side_effect = create_dummy_pdf
        mock_page_nums.side_effect = create_dummy_pdf
        mock_img.side_effect = create_dummy_pdf
        
        # Use real PDF from Testdateien
        real_pdf_path = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'Testdateien', 'Antrag.pdf'))
        if not os.path.exists(real_pdf_path):
            self.skipTest(f"Test PDF not found at {real_pdf_path}")
            
        try:
            # Create EML with this real PDF attached
            msg = EmailMessage()
            msg['Subject'] = 'Test XML Generation with Real PDF'
            msg['From'] = 'sender@example.com'
            msg['To'] = 'recipient@example.com'
            msg.set_content('Body text')
            
            with open(real_pdf_path, 'rb') as f:
                pdf_content = f.read()
            
            # Use the real filename
            msg.add_attachment(pdf_content, maintype='application', subtype='pdf', filename='Antrag.pdf')
            
            eml_path = 'test_real_pdf.eml'
            with open(eml_path, 'wb') as f:
                f.write(msg.as_bytes())
                
            output_pdf = 'output_real_test.pdf'
            output_xml = 'output_real_test.xml'
            
            # Execute
            process_eml(eml_path, output_pdf)
            
            # Assert XML creation
            self.assertTrue(os.path.exists(output_xml), "XML file was not created")
            
            # Parse XML and check content
            tree = ET.parse(output_xml)
            root = tree.getroot()
            
            self.assertEqual(root.tag, 'EmailAnalysis')
            
            attachments = root.find('Attachments')
            self.assertIsNotNone(attachments)
            att = attachments.find('Attachment')
            self.assertIsNotNone(att)
            self.assertEqual(att.get('filename'), 'Antrag.pdf')
            
            # Verify we extracted SOME content
            content = att.find('Content')
            self.assertIsNotNone(content)
            # We don't know exact content of Antrag.pdf, but let's check basic structure
            pages = content.findall('Page')
            self.assertTrue(len(pages) > 0, "No pages extracted from PDF")
            
        finally:
            # Cleanup
            if os.path.exists(eml_path):
                os.remove(eml_path)
            if os.path.exists(output_pdf):
                os.remove(output_pdf)
            if os.path.exists(output_xml):
                os.remove(output_xml)

if __name__ == '__main__':
    unittest.main()
