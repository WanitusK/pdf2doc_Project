import os
from io import BytesIO
from flask import Flask, request, send_file, render_template
from pdf2docx import Converter

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        if 'file' not in request.files:
            return 'No file uploaded', 400

        file = request.files['file']
        if file.filename == '':
            return 'No selected file', 400

        if not file.filename.lower().endswith('.pdf'):
            return 'Only PDF files are allowed', 400

        # ตั้งชื่อไฟล์
        original_filename = file.filename
        base_name = os.path.splitext(original_filename)[0]
        input_path = f'{base_name}_temp.pdf'
        output_path = f'{base_name}_converted.docx'

        # เซฟไฟล์ PDF ชั่วคราว
        file.save(input_path)

        try:
            # แปลง PDF → DOCX
            cv = Converter(input_path)
            cv.convert(output_path)
            cv.close()

            # โหลดไฟล์ .docx เข้า memory
            with open(output_path, 'rb') as f:
                docx_data = f.read()
            buffer = BytesIO(docx_data)
            buffer.seek(0)

        except Exception as e:
            return f'Conversion error: {str(e)}', 500

        finally:
            # ลบไฟล์ทั้งหมด
            if os.path.exists(input_path):
                os.remove(input_path)
            if os.path.exists(output_path):
                os.remove(output_path)

        # ส่ง buffer memory ให้ดาวน์โหลด
        return send_file(
            buffer,
            as_attachment=True,
            download_name=f"{base_name}.docx",
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )

    return render_template('index.html')


if __name__ == '__main__':
    app.run(debug=True)
