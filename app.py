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

        input_path = 'uploaded.pdf'
        output_path = 'converted.docx'
        file.save(input_path)

        try:
            cv = Converter(input_path)
            cv.convert(output_path)
            cv.close()
        except Exception as e:
            return f'Conversion error: {str(e)}', 500

        return send_file(output_path, as_attachment=True)

    return render_template('index.html')


if __name__ == '__main__':
    app.run(debug=True)
