# PDF to Word Converter on GitHub

This repository uses GitHub Actions to convert PDF files to Word documents automatically.

## How to Use

1. Upload your PDF file to the `inputs/` folder
2. Wait for the GitHub Action to run (takes 2-3 minutes)
3. Download your converted Word document from the `outputs/` folder

## How It Works

- When you add a PDF to the `inputs/` folder, GitHub Actions automatically triggers
- The workflow installs all dependencies (Python packages + system tools)
- Your PDF is converted to DOCX using the conversion script
- The converted file is saved to the `outputs/` folder
- You can download the result directly from GitHub

## Limitations

- Maximum file size: GitHub has limits on repository size and workflow run time
- No web interface: This is an automated process, not a live website
- Conversion may take 2-5 minutes depending on PDF complexity
