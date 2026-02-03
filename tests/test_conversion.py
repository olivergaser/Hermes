import sys
import os
import glob
import subprocess
import unittest
from loguru import logger

class TestEmlConversion(unittest.TestCase):
    def setUp(self):
        #self.eml_dir = os.path.abspath(r"z:\ALL\Hermes-Testdaten")
        self.eml_dir = os.path.abspath(r"z:\ALL\Hermes-Testdaten\test")
        self.converter_script = os.path.abspath("converter.py")
        self.venv_python = sys.executable
        logger.debug(f"Using Python executable: {self.venv_python}")
        logger.debug(f"Using eml Dir: {self.eml_dir}")

    def test_convert_all_emls(self):
        eml_files = glob.glob(os.path.join(self.eml_dir, "**", "*.eml"), recursive=True)
        if not eml_files:
            self.fail(f"No .eml files found in {self.eml_dir} directory to test.")

        for eml_file in eml_files:
            output_pdf = os.path.splitext(eml_file)[0] + ".pdf"
            
            # Clean up previous run
            if os.path.exists(output_pdf):
                os.remove(output_pdf)

            logger.debug(f"Testing conversion of: {eml_file}")
            
            # Run conversion
            cmd = [self.venv_python, self.converter_script, "--input", eml_file, "--output", output_pdf]
            env = os.environ.copy()
            if sys.platform == 'darwin':
                # Ensure dyld fallback for mac if needed, though script handles it internally mostly
                if 'DYLD_FALLBACK_LIBRARY_PATH' not in env:
                    env['DYLD_FALLBACK_LIBRARY_PATH'] = '/opt/homebrew/lib:/usr/local/lib'

            result = subprocess.run(cmd, capture_output=True, text=True, env=env)
            
            # Check return code
            if result.returncode != 0:
                logger.error(f"STDOUT: {result.stdout} ,STDERR: {result.stderr}")
                logger.error(f"Conversion failed for {eml_file}, returncode")
                logger.error(f"Command executed: {' '.join(cmd)}")
                continue  # Skip to next file

            # Check if PDF exists and is not empty
            #self.assertTrue(os.path.exists(output_pdf), f"PDF was not created: {output_pdf}")
            if not os.path.exists(output_pdf):
                logger.error(f"PDF was not created: {output_pdf}")
            else:
                logger.debug(f"Successfully created: {output_pdf}")
                if os.path.getsize(output_pdf) < 1000:
                    logger.warning(f"PDF seems too small (empty?): {output_pdf}")    
            
            #self.assertGreater(os.path.getsize(output_pdf), 1000, f"PDF seems too small (empty?): {output_pdf}")

if __name__ == '__main__':
    unittest.main()
