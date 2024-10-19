from flask import Flask, request, render_template, send_from_directory, url_for, jsonify
import os
import fitz  # PyMuPDF
from docx import Document

app = Flask(__name__)

# Ensure 'uploads/' directory exists
if not os.path.exists('uploads'):
    os.makedirs('uploads')

@app.route('/')
def index():
    return render_template('test.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': "No file part in the request"}), 400

    file = request.files['file']

    if file.filename == '':
        return jsonify({'error': "No selected file"}), 400

    if file and file.filename.endswith('.pdf'):
        # Save the uploaded PDF file
        pdf_path = os.path.join('uploads', file.filename)
        file.save(pdf_path)

        try:
            # Create a new DOCX document
            docx_filename = file.filename.replace('.pdf', '.docx')
            docx_path = os.path.join('uploads', docx_filename)
            doc = Document()

            # Open the PDF and extract text using PyMuPDF
            pdf_document = fitz.open(pdf_path)

            for page_num in range(len(pdf_document)):
                page = pdf_document.load_page(page_num)
                text = page.get_text("text")
                doc.add_paragraph(text)

            # Save the DOCX file
            doc.save(docx_path)

            # Check if the file was saved
            if os.path.exists(docx_path):
                print(f"DOCX file saved at: {docx_path}")  # Debugging statement

            # Generate download URL for the converted DOCX file
            download_url = url_for('download_converted_file', filename=docx_filename)
            print(f"Download URL: {download_url}")  # Debugging statement

            # Return JSON response with download link
            return jsonify({'download_link': download_url}), 200

        except Exception as e:
            return jsonify({'error': f"An error occurred while processing the PDF: {str(e)}"}), 500

    return jsonify({'error': "Unsupported file type. Please upload a PDF."}), 400

@app.route('/download/<filename>')
def download_converted_file(filename):
    uploads_dir = os.path.abspath('uploads')
    return send_from_directory(directory=uploads_dir, path=filename, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
