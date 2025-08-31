import os
import sys
import tempfile
import subprocess
import shutil
from pathlib import Path
from docx import Document
from docx.shared import Inches
import pdf2image
import pytesseract
from pdf2docx import Converter
import pdfplumber
import fitz  # PyMuPDF

class PDFToWordConverter:
    def __init__(self):
        # Set up tool paths
        self.tools_dir = Path("tools")
        self.setup_tool_paths()
        
    def setup_tool_paths(self):
        """Configure all tool paths"""
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
                # Check first 3 pages for text
                for i in range(min(3, len(pdf.pages))):
                    text = pdf.pages[i].extract_text()
                    if text and len(text.strip()) > 100:
                        return True
            return False
        except:
            return False
    
    def extract_text_with_ocr(self, pdf_path, output_docx):
        """High-quality OCR extraction with Tesseract"""
        try:
            # Convert PDF to high-resolution images
            images = pdf2image.convert_from_path(
                pdf_path,
                dpi=400,  # High DPI for quality
                poppler_path=str(self.poppler_path) if self.poppler_path.exists() else None,
                grayscale=True  # Better for OCR
            )
            
            doc = Document()
            
            for i, image in enumerate(images):
                # Save temporary image
                with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as temp_img:
                    image.save(temp_img.name, 'PNG', dpi=(400, 400))
                    
                    # Perform OCR with optimized settings
                    text = pytesseract.image_to_string(
                        temp_img.name,
                        lang='eng',
                        config='--psm 6 -c preserve_interword_spaces=1'
                    )
                
                # Add text to document
                if text.strip():
                    paragraph = doc.add_paragraph(text)
                    paragraph.style = 'Normal'
                
                # Add page break except for last page
                if i < len(images) - 1:
                    doc.add_page_break()
                
                # Clean up temp file
                os.unlink(temp_img.name)
            
            doc.save(output_docx)
            return True
            
        except Exception as e:
            print(f"OCR extraction failed: {e}")
            return False
    
    def extract_tables(self, pdf_path, output_docx):
        """Extract tables using multiple methods"""
        try:
            doc = Document(output_docx)
            tables_found = False
            
            # Method 1: Try tabula with Java
            if self.java_path.exists() and self.tabula_jar.exists():
                try:
                    with tempfile.NamedTemporaryFile(suffix='.csv', delete=False) as temp_csv:
                        cmd = [
                            str(self.java_path),
                            '-jar', str(self.tabula_jar),
                            '-f', 'CSV',
                            '-o', temp_csv.name,
                            '-p', 'all',
                            '-l',  # Lattice mode for tables
                            str(pdf_path)
                        ]
                        
                        result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
                        
                        if result.returncode == 0:
                            with open(temp_csv.name, 'r', encoding='utf-8') as f:
                                csv_content = f.read()
                                if csv_content.strip():
                                    tables_found = True
                                    doc.add_paragraph("Extracted Tables:").bold = True
                                    doc.add_paragraph(csv_content)
                    
                    os.unlink(temp_csv.name)
                except:
                    pass
            
            # Method 2: Try pdfplumber as fallback
            if not tables_found:
                try:
                    with pdfplumber.open(pdf_path) as pdf:
                        for page_num, page in enumerate(pdf.pages):
                            tables = page.extract_tables()
                            if tables:
                                tables_found = True
                                doc.add_paragraph(f"Tables from page {page_num + 1}:").bold = True
                                
                                for table in tables:
                                    for row in table:
                                        if any(cell is not None for cell in row):
                                            doc.add_paragraph(" | ".join(str(cell) if cell else "" for cell in row))
                except:
                    pass
            
            if tables_found:
                doc.save(output_docx)
            
            return tables_found
            
        except Exception as e:
            print(f"Table extraction failed: {e}")
            return False
    
    def convert_digital_pdf(self, pdf_path, output_docx):
        """Convert digital PDF with layout preservation"""
        try:
            cv = Converter(pdf_path)
            cv.convert(output_docx, start=0, end=None)
            cv.close()
            return True
        except Exception as e:
            print(f"Digital conversion failed: {e}")
            return False
    
    def enhance_quality(self, output_docx):
        """Enhance document quality for Xerox-like output"""
        try:
            doc = Document(output_docx)
            
            # Set better document properties
            core_props = doc.core_properties
            core_props.title = "Converted Document"
            core_props.subject = "PDF to Word Conversion"
            
            # Set better page margins
            sections = doc.sections
            for section in sections:
                section.top_margin = Inches(1)
                section.bottom_margin = Inches(1)
                section.left_margin = Inches(1)
                section.right_margin = Inches(1)
            
            doc.save(output_docx)
            return True
            
        except Exception as e:
            print(f"Quality enhancement failed: {e}")
            return False
    
    def convert_pdf_to_word(self, pdf_path, output_docx):
        """Main conversion method with fallback strategies"""
        print(f"Starting conversion: {pdf_path}")
        
        # Try digital conversion first
        if self.is_digital_pdf(pdf_path):
            print("Detected digital PDF - using layout preservation...")
            if self.convert_digital_pdf(pdf_path, output_docx):
                print("Digital conversion successful!")
                # Still try to extract tables
                self.extract_tables(pdf_path, output_docx)
                self.enhance_quality(output_docx)
                return True
        
        # Fallback to high-quality OCR
        print("Using high-quality OCR conversion...")
        if self.extract_text_with_ocr(pdf_path, output_docx):
            print("OCR conversion successful!")
            # Extract tables from OCR result
            self.extract_tables(pdf_path, output_docx)
            self.enhance_quality(output_docx)
            return True
        
        print("All conversion methods failed")
        return False

def main():
    if len(sys.argv) != 3:
        print("Usage: python app.py <input_pdf> <output_docx>")
        sys.exit(1)
    
    input_pdf = sys.argv[1]
    output_docx = sys.argv[2]
    
    # Validate input file
    if not os.path.exists(input_pdf):
        print(f"Error: Input file '{input_pdf}' not found")
        sys.exit(1)
    
    if not input_pdf.lower().endswith('.pdf'):
        print("Error: Input file must be a PDF")
        sys.exit(1)
    
    # Initialize converter
    converter = PDFToWordConverter()
    
    # Perform conversion
    success = converter.convert_pdf_to_word(input_pdf, output_docx)
    
    if success:
        print(f"Conversion successful! Output: {output_docx}")
        sys.exit(0)
    else:
        print("Conversion failed")
        sys.exit(1)

if __name__ == "__main__":
    main()
