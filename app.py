from flask import Flask, request, render_template, send_file
from werkzeug.utils import secure_filename
import os
from fpdf import FPDF
from PIL import Image
import pandas as pd
import tempfile

app = Flask(__name__)
UPLOAD_FOLDER = tempfile.gettempdir()
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def convert():
    file = request.files['file']
    if not file:
        return "No file uploaded"

    filename = secure_filename(file.filename)
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(filepath)

    ext = os.path.splitext(filename)[1].lower()
    output_pdf = os.path.join(app.config['UPLOAD_FOLDER'], f"{os.path.splitext(filename)[0]}.pdf")

    try:
        if ext == ".txt":
            with open(filepath, 'r', encoding='utf-8') as f:
                text = f.read()
            pdf = FPDF()
            pdf.add_page()
            pdf.set_auto_page_break(auto=True, margin=15)
            pdf.set_font("Arial", size=12)
            for line in text.split('\n'):
                pdf.cell(200, 10, txt=line, ln=1)
            pdf.output(output_pdf)

        elif ext in [".jpg", ".jpeg", ".png"]:
            image = Image.open(filepath)
            rgb_image = image.convert('RGB')
            rgb_image.save(output_pdf)

        elif ext == ".csv":
            df = pd.read_csv(filepath)
            text = df.to_string(index=False)
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Arial", size=10)
            for line in text.split('\n'):
                pdf.cell(200, 8, txt=line, ln=1)
            pdf.output(output_pdf)

        elif ext == ".xlsx":
            df = pd.read_excel(filepath)
            text = df.to_string(index=False)
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Arial", size=10)
            for line in text.split('\n'):
                pdf.cell(200, 8, txt=line, ln=1)
            pdf.output(output_pdf)

        else:
            return "Unsupported file format"

        return send_file(output_pdf, as_attachment=True)

    except Exception as e:
        return f"Error during conversion: {e}"

    finally:
        os.remove(filepath)
