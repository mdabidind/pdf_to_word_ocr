# PDF to Word Converter with OCR

Automatically convert PDF files to Word documents using GitHub Actions.

## How to Use

1. **Upload PDF**: Add your PDF file to the `inputs/` folder
2. **Automatic Conversion**: The system will automatically convert your file within 2-3 minutes
3. **Download**: Get your converted Word document from the `outputs/` folder

## Features

- Automatic OCR for scanned PDFs
- Table extraction and formatting
- Layout preservation
- Batch processing support

## Supported Formats

- Input: PDF files (.pdf)
- Output: Word documents (.docx)

## Processing Time

Conversion typically takes 2-5 minutes depending on:
- File size
- Number of pages
- Complexity of content

## Notes

- Maximum file size: 100MB recommended
- Complex layouts may require manual adjustment
- Check the `outputs/` folder for your converted files
