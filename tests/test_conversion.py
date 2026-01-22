import sys
import os
import glob
import subprocess
import unittest
import shutil

class TestEmlConversion(unittest.TestCase):
    def setUp(self):
        self.eml_dir = os.path.abspath(r"z:\ALL\Hermes-Testdaten")
        self.converter_script = os.path.abspath("converter.py")
        self.venv_python = sys.executable

    def test_convert_all_emls(self):
        eml_files = glob.glob(os.path.join(self.eml_dir, "*.eml"))
        if not eml_files:
            self.fail("No .eml files found in 'eml' directory to test.")

        for eml_file in eml_files:
            base_name = os.path.splitext(os.path.basename(eml_file))[0]
            output_pdf = os.path.join(self.eml_dir, f"{base_name}.pdf")
            
            # Clean up previous run
            if os.path.exists(output_pdf):
                os.remove(output_pdf)

            print(f"Testing conversion of: {eml_file}")
            
            # Run conversion
            cmd = [self.venv_python, self.converter_script, eml_file, output_pdf]
            env = os.environ.copy()
            # Ensure dyld fallback for mac if needed, though script handles it internally mostly
            if 'DYLD_FALLBACK_LIBRARY_PATH' not in env:
                 env['DYLD_FALLBACK_LIBRARY_PATH'] = '/opt/homebrew/lib:/usr/local/lib'

            result = subprocess.run(cmd, capture_output=True, text=True, env=env)
            
            # Check return code
            if result.returncode != 0:
                print(f"STDOUT: {result.stdout}")
                print(f"STDERR: {result.stderr}")
            self.assertEqual(result.returncode, 0, f"Conversion failed for {eml_file}")

            # Check if PDF exists and is not empty
            self.assertTrue(os.path.exists(output_pdf), f"PDF was not created: {output_pdf}")
            self.assertGreater(os.path.getsize(output_pdf), 1000, f"PDF seems too small (empty?): {output_pdf}")
            print(f"Successfully created: {output_pdf}")

if __name__ == '__main__':
    unittest.main()
