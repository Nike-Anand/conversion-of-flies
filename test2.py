from flask import Flask, request, render_template, send_file
import pdfkit
import os
from docx import Document

app = Flask(__name__)

# Set the path for wkhtmltopdf executable
config = pdfkit.configuration(wkhtmltopdf=r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe')

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return "No file part"

    file = request.files['file']

    if file.filename == '':
        return "No selected file"

    if file and file.filename.endswith('.docx'):
        # Save the uploaded DOCX file
        docx_path = os.path.join('uploads', file.filename)
        file.save(docx_path)

        # Convert DOCX to HTML
        document = Document(docx_path)
        html_content = "<html><body>"

        for paragraph in document.paragraphs:
            html_content += f"<p>{paragraph.text}</p>"

        html_content += "</body></html>"

        # Convert HTML to PDF
        pdf_path = os.path.join('uploads', file.filename.replace('.docx', '.pdf'))
        pdfkit.from_string(html_content, pdf_path, configuration=config)

        return send_file(pdf_path, as_attachment=True)

    return "File type not supported. Please upload a DOCX file."

if __name__ == '__main__':
    if not os.path.exists('uploads'):
        os.makedirs('uploads')
    app.run(debug=True)
