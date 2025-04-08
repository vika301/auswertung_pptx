from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
from pptx.enum.chart import XL_TICK_MARK, XL_TICK_LABEL_POSITION
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from PIL import Image

def add_bullet_textbox(prs, slide, text_list, font_name="Frutiger 44 Light", font_size=Pt(16), font_color=RGBColor(0, 0, 0)):
    """
    Fügt eine TextBox mit Spiegelstrichen auf der rechten Hälfte der Folie hinzu.

    :param slide: Die Folie, auf die die TextBox eingefügt wird.
    :param text_list: Eine Liste von Strings, die als Spiegelstriche hinzugefügt werden.
    :param font_name: Die gewünschte Schriftart (Standard: Frutiger 44 Light).
    :param font_size: Die Schriftgröße (Standard: 16pt).
    :param font_color: Die Schriftfarbe (Standard: Schwarz).
    """
    # Position der TextBox (rechte Hälfte der Folie)
    left = Inches(7.5)    # Start auf der rechten Seite der Folie
    top = Inches(1.5)     # Obere Position
    width = Inches(5)  # Breite der TextBox
    height = Inches(5) # Höhe der TextBox

    # TextBox hinzufügen
    textbox = slide.shapes.add_textbox(left, top, width, height)
    text_frame = textbox.text_frame
    text_frame.word_wrap = True
    text_frame.vertical_anchor = MSO_ANCHOR.TOP  # Text oben ausrichten

    # Ersten Absatz hinzufügen (erster Spiegelstrich)
    for i, text in enumerate(text_list):
        if i == 0:
            p = text_frame.paragraphs[0]  # Erster Absatz
        else:
            p = text_frame.add_paragraph()  # Neue Zeile für weitere Spiegelstriche
        p.text = f"• {text}"  # Spiegelstrich hinzufügen
        p.font.size = font_size
        p.font.name = font_name
        p.font.color.rgb = font_color
        p.font.bold = True  # **Text fett setzen**
        p.alignment = PP_ALIGN.LEFT  # Links ausrichten

def set_background_color(slide, color):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color

def color_bottom_slide(prs, slide, color):
    slide_height = prs.slide_height
    height = slide_height // 20  # Bottom 1/20th of the slide
    shape = slide.shapes.add_shape(
        1, Inches(0), slide_height - height, prs.slide_width, height  # 1 corresponds to MSO_SHAPE.RECTANGLE
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.color.rgb = color  # No border color

# Diese Funktion entfernt das Hinzufügen des Logos
def apply_design_modifications(prs, background_color, bottom_color):
    for slide in prs.slides:
        set_background_color(slide, background_color)
        color_bottom_slide(prs, slide, bottom_color)

def modify_bar_chart_colors(chart, colors):
    for i, series in enumerate(chart.plots[0].series):
        series.format.fill.solid()
        series.format.fill.fore_color.rgb = colors[i % len(colors)]
        for point in series.points:
            point.format.fill.solid()
            point.format.fill.fore_color.rgb = colors[i % len(colors)]

def modify_pie_chart_colors(chart, colors):
    for i, point in enumerate(chart.plots[0].series[0].points):
        point.format.fill.solid()
        point.format.fill.fore_color.rgb = colors[i % len(colors)]

def modify_bar_chart(chart, shape, prs, font_color, colors):
    modify_bar_chart_colors(chart, colors)
    chart.value_axis.minimum_scale = 0  # Minimum value of the axis set to 0
    chart.value_axis.maximum_scale = 1  # Maximum value of the axis set to 1
    shape.top = int(prs.slide_height / 5)  # Shift the chart down
    shape.width = int(prs.slide_width / 2)  # Set width of the chart

def modify_pie_chart(chart, shape, prs, font_color, colors):
    modify_pie_chart_colors(chart, colors)
    shape.left = 0  # Position the chart on the left half of the slide
    shape.top = int(prs.slide_height / 5)  # Shift the chart down
    shape.width = int(prs.slide_width / 2)  # Set the width of the chart

# Main function to modify the presentation (no logo added)
def modify_presentation(input_pptx, output_pptx):
    prs = Presentation(input_pptx)

    background_color = RGBColor(174, 177, 192)  # Background color
    bottom_color = RGBColor(255, 105, 86)  # Color for the bottom bar

    chart_colors = [
        RGBColor(255, 105, 86),  # Red
        RGBColor(232, 230, 215),  # Beige
        RGBColor(255, 255, 255),  # White
        RGBColor(0, 0, 0)  # Black
    ]

    # Apply design modifications to all slides
    apply_design_modifications(prs, background_color, bottom_color)

    for slide in prs.slides:
        text_items = ["Erster Punkt", "Zweiter Punkt"]
        add_bullet_textbox(prs, slide, text_items)

        for shape in slide.shapes:
            if shape.has_chart:
                chart = shape.chart
                chart_type = chart.chart_type

                if chart_type in [XL_CHART_TYPE.COLUMN_CLUSTERED, XL_CHART_TYPE.COLUMN_STACKED, XL_CHART_TYPE.BAR_CLUSTERED]:
                    modify_bar_chart(chart, shape, prs, RGBColor(0, 0, 0), chart_colors)
                elif chart_type == XL_CHART_TYPE.PIE:
                    modify_pie_chart(chart, shape, prs, RGBColor(0, 0, 0), chart_colors)

    prs.save(output_pptx)
