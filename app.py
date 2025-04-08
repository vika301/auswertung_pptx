import sys  # Wir importieren sys, um den Pfad hinzuzufügen
import os
from flask import Flask, render_template, request, send_file
from werkzeug.utils import secure_filename
from io import BytesIO

# Füge den utils-Ordner zum sys.path hinzu, damit der Import funktioniert
sys.path.append(os.path.join(os.path.dirname(__file__), 'utils'))

# Importiere die PowerPoint-Verarbeitungslogik
from pptx_logic import modify_presentation

# Flask App initialisieren
app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'

# Sicherstellen, dass der Upload-Ordner existiert
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# Pfad zum festen Logo im static-Ordner
logo_path = os.path.join(app.static_folder, 'logo.jpg')


@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        # PowerPoint-Datei empfangen
        pptx_file = request.files['file']

        if not pptx_file:
            return "Bitte PowerPoint hochladen."

        # Speichern der PowerPoint-Datei
        pptx_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(pptx_file.filename))
        pptx_file.save(pptx_path)

        # Speicherort für die bearbeitete PowerPoint
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], "output.pptx")

        # Anwendung der PowerPoint-Verarbeitungslogik (Hinzufügen des Logos)
        modify_presentation(pptx_path, output_path, logo_path)

        # Überarbeitete PowerPoint-Datei als Download zurückgeben
        return send_file(output_path, as_attachment=True, download_name="bearbeitet.pptx")

    # Bei GET-Anfragen das HTML-Formular anzeigen
    return render_template('index.html')


if __name__ == '__main__':
    # Den richtigen Port für Render festlegen
    port = int(os.environ.get("PORT", 5000))
    app.run(debug=True, host='0.0.0.0', port=port)
