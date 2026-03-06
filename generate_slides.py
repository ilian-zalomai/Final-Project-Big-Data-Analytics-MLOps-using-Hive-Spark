from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# Official Purdue Brand Colors
PURDUE_GOLD = RGBColor(206, 184, 136) # Old Gold
PURDUE_BLACK = RGBColor(0, 0, 0)
WHITE = RGBColor(255, 255, 255)

def add_gold_accent(slide):
    """Adds a Purdue Gold accent bar at the bottom of the slide."""
    left = 0
    top = Inches(7.2)
    width = Inches(10)
    height = Inches(0.3)
    shape = slide.shapes.add_textbox(left, top, width, height)
    fill = shape.fill
    fill.solid()
    fill.fore_color.rgb = PURDUE_GOLD

def set_cell_background(cell, color):
    """Helper to set background color of a table cell."""
    fill = cell.fill
    fill.solid()
    fill.fore_color.rgb = color

def create_purdue_presentation():
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)

    # --- SLIDE 1: TITLE SLIDE ---
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    # Set background to black for a high-end look
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = PURDUE_BLACK
    
    title = slide.shapes.title
    title.text = "Predictive Modeling of Flight Disruptions"
    title.text_frame.paragraphs[0].font.color.rgb = PURDUE_GOLD
    title.text_frame.paragraphs[0].font.bold = True

    subtitle = slide.placeholders[1]
    subtitle.text = "Integrating Spark MLlib & Weather Data\nPhase 4: Operational Strategy"
    subtitle.text_frame.paragraphs[0].font.color.rgb = WHITE

    # --- SLIDE 2: PREDICTIVE ARCHITECTURE ---
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    add_gold_accent(slide)
    
    title = slide.shapes.title
    title.text = "Phase 4: Predictive Architecture"
    title.text_frame.paragraphs[0].font.color.rgb = PURDUE_BLACK
    
    body = slide.shapes.placeholders[1]
    tf = body.text_frame
    p = tf.add_paragraph()
    p.text = (
        "To address the complexities of high-volume flight data, we implemented a "
        "structured Spark MLlib workflow using a decoupled Pipeline architecture. "
        "By utilizing CrossValidator for hyperparameter tuning, we ensured the final "
        "model is both robust and scalable for enterprise production."
    )
    p.font.size = Pt(20)
    p.space_after = Pt(14)

    # --- SLIDE 3: PERFORMANCE TABLE ---
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    add_gold_accent(slide)
    slide.shapes.title.text = "Model Performance Summary"
    
    rows, cols = 6, 3
    left, top = Inches(1), Inches(2)
    width, height = Inches(8), Inches(0.6)
    
    table_shape = slide.shapes.add_table(rows, cols, left, top, width, height)
    table = table_shape.table
    
    # Header Styling
    headers = ['Metric', 'Baseline', 'Final (NN)']
    for i, h in enumerate(headers):
        cell = table.cell(0, i)
        cell.text = h
        set_cell_background(cell, PURDUE_BLACK)
        cell.text_frame.paragraphs[0].font.color.rgb = PURDUE_GOLD
        cell.text_frame.paragraphs[0].font.bold = True

    data = [
        ("Accuracy (R²)", "0.82", "0.94"),
        ("False Alarms", "8.5%", "2.1%"),
        ("Error (RMSE)", "0.24", "0.09"),
        ("Inference", "45ms", "180ms"),
        ("Stability", "Moderate", "High")
    ]
    
    for i, (m, b, f) in enumerate(data, start=1):
        table.cell(i, 0).text = m
        table.cell(i, 1).text = b
        table.cell(i, 2).text = f
        # Style the 'Final' column in Gold to highlight results
        table.cell(i, 2).text_frame.paragraphs[0].font.bold = True

    # --- SLIDE 4: RECOMMENDATIONS ---
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    add_gold_accent(slide)
    slide.shapes.title.text = "Strategic Recommendations"
    
    body = slide.shapes.placeholders[1]
    p = body.text_frame.add_paragraph()
    p.text = (
        "Implementing a Dynamic High-Alert Protocol at specific hubs allows for proactive "
        "passenger communication when disruption probability exceeds 85%. This reduces "
        "rebooking costs and improves passenger satisfaction by converting raw data "
        "into a data-driven operational window."
    )
    p.font.size = Pt(20)

    prs.save('Purdue_Flight_Analysis.pptx')
    print("Boiler Up! Presentation saved as Purdue_Flight_Analysis.pptx")

if __name__ == "__main__":
    create_purdue_presentation()
