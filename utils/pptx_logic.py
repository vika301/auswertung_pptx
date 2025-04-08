from pptx import Presentation


def modify_presentation(input_pptx, output_pptx, logo_path):
    # Lade die Eingabepowerpoint
    prs = Presentation(input_pptx)

    # Logo auf jeder Folie hinzufügen
    for slide in prs.slides:
        add_logo(prs, slide, logo_path)

    # Präsentation speichern
    prs.save(output_pptx)


def add_logo(prs, slide, logo_path):
    # Berechne die Position für das Logo (obere rechte Ecke)
    logo_left = prs.slide_width - 100  # Position vom linken Rand (100px Abstand)
    logo_top = 0  # Position vom oberen Rand (0px)

    # Logo als Bild hinzufügen
    slide.shapes.add_picture(logo_path, logo_left, logo_top, width=100)
