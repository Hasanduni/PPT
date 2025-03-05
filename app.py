import streamlit as st
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.util import Inches
import requests
import os

def get_text_from_huggingface(prompt):
    API_URL = "https://api-inference.huggingface.co/models/facebook/bart-large-cnn"
    HF_API_KEY = os.getenv("HF_API_KEY")  # Load API key from environment variable
    
    if not HF_API_KEY:
        return "Error: API key is missing."

    headers = {"Authorization": f"Bearer {HF_API_KEY}"}
    payload = {"inputs": prompt}

    response = requests.post(API_URL, headers=headers, json=payload)

    if response.status_code == 200:
        try:
            return response.json()[0]['summary_text']
        except (KeyError, IndexError):
            return "Error: Unexpected response format."
    else:
        return f"Error: {response.status_code} - {response.text}"

def add_title_slide(prs, title, subtitle):
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = title
    slide.placeholders[1].text = subtitle

def add_chart_slide(prs, title, chart_data):
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = title
    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(1), Inches(1.5), Inches(6), Inches(4.5), chart_data
    ).chart
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM

def add_text_box_slide(prs, title, text):
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = title
    textbox = slide.shapes.add_textbox(Inches(1), Inches(1.5), Inches(6), Inches(4))
    textbox.text_frame.text = text

def generate_presentation(topic):
    prs = Presentation()
    add_title_slide(prs, topic, "Generated using Streamlit and python-pptx")
    
    description = get_text_from_huggingface(f"Generate a brief summary about {topic}")
    
    chart_data = CategoryChartData()
    chart_data.categories = ['A', 'B', 'C']
    chart_data.add_series('Series 1', (10, 20, 30))
    add_chart_slide(prs, "Sample Chart", chart_data)
    
    add_text_box_slide(prs, "About " + topic, description)
    
    filename = "generated_presentation.pptx"
    prs.save(filename)
    return filename

st.title("📊 Generate PowerPoint Presentation")
topic = st.text_input("Enter Topic:")

if st.button("Generate PPT"):
    if topic:
        pptx_file = generate_presentation(topic)
        with open(pptx_file, "rb") as file:
            st.download_button(label="📥 Download Presentation", data=file, file_name=pptx_file, mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
    else:
        st.warning("⚠️ Please enter a topic.")
