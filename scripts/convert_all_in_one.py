# Your original convert_all_in_one.py content goes here
# But let's add a main function at the bottom:

if __name__ == "__main__":
    import sys
    if len(sys.argv) != 3:
        print("Usage: python convert_all_in_one.py <input_pdf> <output_docx>")
        sys.exit(1)
        
    input_pdf = sys.argv[1]
    output_docx = sys.argv[2]
    
    success = convert_pdf_to_docx(input_pdf, output_docx)
    sys.exit(0 if success else 1)
