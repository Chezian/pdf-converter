from flask import Flask, request, render_template, send_file
from werkzeug.utils import secure_filename
import os
from fpdf import FPDF
from PIL import Image
import pandas as pd
import tempfile
import docx
import traceback

app = Flask(__name__)

# Upload folder setup
UPLOAD_FOLDER = os.path.join(os.getcwd(), "uploads")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def convert():
    try:
        file = request.files.get('file')
        if not file or file.filename == '':
            return "No file selected"

        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)

        ext = os.path.splitext(filename)[1].lower()
        output_pdf = os.path.join(app.config['UPLOAD_FOLDER'], f"{os.path.splitext(filename)[0]}.pdf")

        pdf = FPDF()
        pdf.add_page()
        pdf.set_auto_page_break(auto=True, margin=15)
        pdf.set_font("Arial", size=12)

        if ext == ".txt":
            with open(filepath, 'r', encoding='utf-8') as f:
                text = f.read()
            for line in text.split('\n'):
                pdf.cell(200, 10, txt=line, ln=1)
            pdf.output(output_pdf)

        elif ext == ".csv":
            df = pd.read_csv(filepath)
            text = df.to_string(index=False)
            for line in text.split('\n'):
                pdf.cell(200, 8, txt=line, ln=1)
            pdf.output(output_pdf)

        elif ext == ".xlsx":
            df = pd.read_excel(filepath)
            text = df.to_string(index=False)
            for line in text.split('\n'):
                pdf.cell(200, 8, txt=line, ln=1)
            pdf.output(output_pdf)

        elif ext == ".docx":
            doc_file = docx.Document(filepath)
            text = '\n'.join([para.text for para in doc_file.paragraphs])
            for line in text.split('\n'):
                pdf.cell(200, 10, txt=line, ln=1)
            pdf.output(output_pdf)

        elif ext in [".jpg", ".jpeg", ".png"]:
            image = Image.open(filepath)
            rgb_image = image.convert('RGB')
            rgb_image.save(output_pdf)

        else:
            return "Unsupported file format"

        if not os.path.exists(output_pdf):
            return "Error: PDF was not created."

        return send_file(output_pdf, as_attachment=True)

    except Exception as e:
        traceback.print_exc()
        return f"Error during conversion: {e}"

    finally:
        # Clean up uploaded file (but keep output_pdf for download)
        if os.path.exists(filepath):
            try:
                os.remove(filepath)
            except:
                pass

if __name__ == '__main__':
    app.run(debug=True)
