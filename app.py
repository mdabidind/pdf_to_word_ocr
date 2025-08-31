import os
import sys
import tempfile
import subprocess
from pathlib import Path
from docx import Document
import pdf2image
import pytesseract
from pdf2docx import Converter
import pdfplumber

class PDFToWordConverter:
    def __init__(self):
        self.setup_tool_paths()
        
    def setup_tool_paths(self):
        """Configure all tool paths"""
        self.tools_dir = Path("tools")
        
        # Poppler paths
        self.poppler_path = self.tools_dir / "poppler" / "bin"
        if self.poppler_path.exists():
            os.environ['PATH'] += os.pathsep + str(self.poppler_path)
        
        # Tesseract path
        self.tesseract_path = self.tools_dir / "tesseract" / "tesseract.exe"
        if self.tesseract_path.exists():
            pytesseract.pytesseract.tesseract_cmd = str(self.tesseract_path)
        
        # Java path for tabula
        self.java_path = self.tools_dir / "java" / "bin" / "java.exe"
        self.tabula_jar = self.tools_dir / "tabula.jar"
    
    def is_digital_pdf(self, pdf_path):
        """Check if PDF has selectable text"""
        try:
            with pdfplumber.open(pdf_path) as pdf:
                for i in range(min(3, len(pdf.pages))):
                    text = pdf.pages[i].extract_text()
                    if text and len(text.strip()) > 50:
                        return True
            return False
        except:
            return False
    
    def extract_text_with_ocr(self, pdf_path, output_docx):
        """High-quality OCR extraction"""
        try:
            images = pdf2image.convert_from_path(
                pdf_path,
                dpi=300,
                poppler_path=str(self.poppler_path) if self.poppler_path.exists() else None
            )
            
            doc = Document()
            
            for i, image in enumerate(images):
                with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as temp_img:
                    image.save(temp_img.name, 'PNG')
                    
                    text = pytesseract.image_to_string(
                        temp_img.name,
                        lang='eng',
                        config='--psm 6'
                    )
                
                if text.strip():
                    doc.add_paragraph(text)
                
                if i < len(images) - 1:
                    doc.add_page_break()
                
                os.unlink(temp_img.name)
            
            doc.save(output_docx)
            return True
            
        except Exception as e:
            print(f"OCR failed: {e}")
            return False
    
    def convert_digital_pdf(self, pdf_path, output_docx):
        """Convert digital PDF"""
        try:
            cv = Converter(pdf_path)
            cv.convert(output_docx)
            cv.close()
            return True
        except Exception as e:
            print(f"Digital conversion failed: {e}")
            return False
    
    def convert_pdf_to_word(self, pdf_path, output_docx):
        """Main conversion method"""
        print(f"Converting: {pdf_path}")
        
        if self.is_digital_pdf(pdf_path):
            print("Using digital conversion")
            return self.convert_digital_pdf(pdf_path, output_docx)
        else:
            print("Using OCR conversion")
            return self.extract_text_with_ocr(pdf_path, output_docx)

def main():
    if len(sys.argv) != 3:
        print("Usage: python app.py input.pdf output.docx")
        return
    
    input_pdf = sys.argv[1]
    output_docx = sys.argv[2]
    
    if not os.path.exists(input_pdf):
        print("Input file not found")
        return
    
    converter = PDFToWordConverter()
    success = converter.convert_pdf_to_word(input_pdf, output_docx)
    
    if success:
        print(f"Success! Created: {output_docx}")
    else:
        print("Conversion failed")

if __name__ == "__main__":
    main()
