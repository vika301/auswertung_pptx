import sys
import os
from flask import Flask, render_template, request, send_file
from werkzeug.utils import secure_filename
from io import BytesIO

sys.path.append(os.path.join(os.path.dirname(__file__), 'utils'))

from pptx_logic import modify_presentation

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)


@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        pptx_file = request.files['file']

        if not pptx_file:
            return "Bitte PowerPoint hochladen."

        pptx_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(pptx_file.filename))
        pptx_file.save(pptx_path)

        output_path = os.path.join(app.config['UPLOAD_FOLDER'], "output.pptx")

        # Remove logo and process the presentation
        modify_presentation(pptx_path, output_path)

        return send_file(output_path, as_attachment=True, download_name="bearbeitet.pptx")

    return render_template('index.html')


if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(debug=True, host='0.0.0.0', port=port)
