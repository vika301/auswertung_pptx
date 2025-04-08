import sys
import os
from flask import Flask, render_template, request, send_file
from werkzeug.utils import secure_filename
from io import BytesIO

# Füge den utils-Ordner zum sys.path hinzu, damit der Import funktioniert
sys.path.append(os.path.join(os.path.dirname(__file__), 'utils'))

from pptx_logic import modify_presentation

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# Definiere den festen Pfad für das Logo im static-Ordner
logo_path = os.path.join(app.static_folder, 'logo.jpg')  # Name des Logos, z.B. logo.jpg


@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        pptx_file = request.files['file']  # Der Name des Files im HTML-Formular ist "file"

        if not pptx_file:
            return "Bitte PowerPoint hochladen."

        # Save the PowerPoint file
        pptx_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(pptx_file.filename))
        pptx_file.save(pptx_path)

        # Speicherort der bearbeiteten PowerPoint-Datei
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], "output.pptx")

        # Funktion zur Bearbeitung der PowerPoint-Präsentation
        modify_presentation(pptx_path, output_path, logo_path)

        # Senden der bearbeiteten PowerPoint-Datei zum Download
        return send_file(output_path, as_attachment=True, download_name="bearbeitet.pptx")

    return render_template('index.html')


if __name__ == '__main__':
    app.run(debug=True)
