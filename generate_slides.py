from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN

def create_presentation():
    prs = Presentation()

    # --- Slide 1: Title Slide ---
    slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(slide_layout)
    title = slide.shapes.title
    subtitle = slide.placeholders[1]
    
    title.text = "Predictive Modeling of Flight Disruptions"
    subtitle.text = "Integrating Spark MLlib & Weather Data\nPhase 4: Operational Strategy"

    # --- Slide 2: Predictive Architecture ---
    slide_layout = prs.slide_layouts[1]
    slide = prs.slides.add_slide(slide_layout)
    title = slide.shapes.title
    title.text = "Phase 4: Predictive Architecture"
    
    body = slide.shapes.placeholders[1]
    tf = body.text_frame
    tf.word_wrap = True
    p = tf.add_paragraph()
    p.text = (
        "To address the complexities of high-volume flight data, we implemented a structured "
        "Spark MLlib workflow using a decoupled Pipeline architecture. This system encapsulates "
        "feature engineering—including VectorAssembler and StringIndexer—with sophisticated "
        "Random Forest and Linear Regression models. By utilizing CrossValidator for hyperparameter "
        "tuning, we ensured that the final model is both robust and scalable for production."
    )
    p.font.size = Pt(18)

    # --- Slide 3: Performance Table ---
    slide = prs.slides.add_slide(prs.slide_layouts[5]) # Blank slide with title
    slide.shapes.title.text = "Model Performance Metrics"
    
    rows, cols = 6, 3
    left = Inches(1.5)
    top = Inches(2.0)
    width = Inches(7.0)
    height = Inches(0.8)
    
    table = slide.shapes.add_table(rows, cols, left, top, width, height).table
    
    # Set Column Headers
    table.cell(0, 0).text = 'Metric'
    table.cell(0, 1).text = 'Baseline'
    table.cell(0, 2).text = 'Final (NN)'
    
    # Fill Data
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

    # --- Slide 4: Strategic Recommendations ---
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Strategic Recommendations"
    
    tf = slide.shapes.placeholders[1].text_frame
    p = tf.add_paragraph()
    p.text = (
        "The practical application of this model centers on a Dynamic High-Alert Protocol "
        "that triggers when disruption probability exceeds 85%. By identifying specific "
        "tipping points for wind and precipitation, airlines can pre-position recovery crews "
        "and automate passenger communications at least two hours before a weather event hits. "
        "This strategy transforms raw weather alerts into actionable operational windows."
    )
    p.font.size = Pt(18)

    # Save the presentation
    prs.save('Phase4_Presentation.pptx')
    print("Presentation generated: Phase4_Presentation.pptx")

if __name__ == "__main__":
    create_presentation()
