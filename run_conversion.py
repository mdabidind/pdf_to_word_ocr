#!/usr/bin/env python3
"""
Standalone script to convert all PDFs in inputs/ folder to outputs/ folder
"""
import os
import sys
import glob
from convert_all_in_one import convert_pdf_to_docx

def main():
    # Create outputs directory if it doesn't exist
    os.makedirs('outputs', exist_ok=True)
    
    # Convert all PDFs in inputs folder
    pdf_files = glob.glob('inputs/*.pdf')
    
    if not pdf_files:
        print("No PDF files found in inputs/ folder")
        return
    
    for pdf_path in pdf_files:
        filename = os.path.basename(pdf_path)
        docx_filename = os.path.splitext(filename)[0] + '.docx'
        output_path = os.path.join('outputs', docx_filename)
        
        print(f"Converting {pdf_path} to {output_path}")
        success = convert_pdf_to_docx(pdf_path, output_path)
        
        if success:
            print(f"✓ Successfully converted {filename}")
        else:
            print(f"✗ Failed to convert {filename}")

if __name__ == "__main__":
    main()
