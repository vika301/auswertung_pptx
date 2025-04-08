from pptx import Presentation

def modify_presentation(input_pptx, output_pptx, logo_path=None):
    try:
        # Lade die Eingabepowerpoint
        prs = Presentation(input_pptx)

        # Präsentation speichern ohne Logo
        prs.save(output_pptx)
    except Exception as e:
        print(f"Fehler bei der Bearbeitung der Präsentation: {e}")
