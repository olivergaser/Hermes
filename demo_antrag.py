import os
import sys
import subprocess
from email.message import EmailMessage

# Paths
pdf_path = os.path.abspath(r"Testdateien\Antrag.pdf")
eml_path = "antrag_demo.eml"
converter_script = "converter.py"
python_exe = sys.executable

if not os.path.exists(pdf_path):
    print(f"Error: {pdf_path} not found.")
    sys.exit(1)

# Create EML
print(f"Creating {eml_path} with attachment {pdf_path}...")
msg = EmailMessage()
msg['Subject'] = 'Demo Antrag PDF Extraction'
msg['From'] = 'demo@example.com'
msg['To'] = 'user@example.com'
msg.set_content('Please find the attached application form.')

with open(pdf_path, 'rb') as f:
    pdf_data = f.read()
    
msg.add_attachment(pdf_data, maintype='application', subtype='pdf', filename='Antrag.pdf')

with open(eml_path, 'wb') as f:
    f.write(msg.as_bytes())

# Run Converter
print("Running converter...")
cmd = [python_exe, converter_script, "-i", eml_path]
subprocess.run(cmd, check=True)

print("Done.")
