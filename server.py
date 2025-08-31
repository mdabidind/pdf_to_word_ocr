from flask import Flask, request, jsonify, send_file
import os
import uuid
from app import PDFToWordConverter

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024

UPLOAD_DIR = 'uploads'
OUTPUT_DIR = 'outputs'

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

@app.route('/')
def serve_index():
    try:
        with open('index.html', 'r', encoding='utf-8') as f:
            return f.read()
    except:
        return "HTML file not found", 404

@app.route('/convert', methods=['POST'])
def convert():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No file selected'}), 400
    
    if not file.filename.lower().endswith('.pdf'):
        return jsonify({'error': 'Please upload a PDF file'}), 400

    # Generate unique IDs
    file_id = str(uuid.uuid4())
    input_path = os.path.join(UPLOAD_DIR, f"{file_id}.pdf")
    output_path = os.path.join(OUTPUT_DIR, f"{file_id}.docx")
    
    try:
        # Save uploaded file
        file.save(input_path)
        
        # Convert PDF to Word
        converter = PDFToWordConverter()
        success = converter.convert_pdf_to_word(input_path, output_path)
        
        if success and os.path.exists(output_path):
            return jsonify({
                'success': True,
                'download_url': f'/download/{file_id}',
                'filename': file.filename.replace('.pdf', '.docx')
            })
        else:
            return jsonify({'error': 'Conversion failed'}), 500
            
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    finally:
        # Clean up
        if os.path.exists(input_path):
            os.remove(input_path)

@app.route('/download/<file_id>')
def download(file_id):
    output_path = os.path.join(OUTPUT_DIR, f"{file_id}.docx")
    
    if os.path.exists(output_path):
        return send_file(
            output_path,
            as_attachment=True,
            download_name=f"converted_{file_id}.docx"
        )
    else:
        return jsonify({'error': 'File not found'}), 404

if __name__ == '__main__':
    print("Starting PDF to Word Converter Server...")
    print("Server running at http://localhost:5000")
    app.run(host='0.0.0.0', port=5000, debug=True)
