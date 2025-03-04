import streamlit as st
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.util import Inches
from pyngrok import ngrok

# Set Ngrok authentication token (Replace with your actual token)
NGROK_AUTH_TOKEN = "2tq7y9AH9zfvbcCyfQpNtWrSnf3_5xaS9qN23NFvxyWBwtpqR"  # Replace with your actual Ngrok token
ngrok.set_auth_token(NGROK_AUTH_TOKEN)

# Start Ngrok tunnel for Streamlit (port 8501)
ngrok.kill()  # Kills existing tunnels
public_url = ngrok.connect(8501).public_url
st.write(f"🌍 Public URL: [Click Here]({public_url})")

def add_title_slide(prs, title, subtitle):
    """Add a title slide with a title and subtitle."""
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = title
    slide.placeholders[1].text = subtitle

def add_chart_slide(prs, title, chart_data):
    """Add a chart slide."""
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = title
    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(1), Inches(1.5), Inches(6), Inches(4.5), chart_data
    ).chart
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM

def add_text_box_slide(prs, title, text):
    """Add a slide with a text box."""
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = title
    textbox = slide.shapes.add_textbox(Inches(1), Inches(1.5), Inches(6), Inches(4))
    textbox.text_frame.text = text

def generate_presentation(topic, description):
    """Generate a PowerPoint presentation based on user input."""
    prs = Presentation()
    add_title_slide(prs, topic, "Generated using Streamlit and python-pptx")
    
    # Add a sample chart
    chart_data = CategoryChartData()
    chart_data.categories = ['A', 'B', 'C']
    chart_data.add_series('Series 1', (10, 20, 30))
    add_chart_slide(prs, "Sample Chart", chart_data)

    # Add a text slide with user description
    add_text_box_slide(prs, "About " + topic, description)

    # Save file
    filename = "generated_presentation.pptx"
    prs.save(filename)
    return filename

# Streamlit UI
st.title("📊 Generate PowerPoint Presentation")

topic = st.text_input("Enter Topic:")
description = st.text_area("Enter Description:")

if st.button("Generate PPT"):
    if topic and description:
        pptx_file = generate_presentation(topic, description)
        with open(pptx_file, "rb") as file:
            st.download_button(label="📥 Download Presentation", data=file, file_name=pptx_file, mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
    else:
        st.warning("⚠️ Please enter both a topic and description.")
