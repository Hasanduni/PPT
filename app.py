import streamlit as st
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.util import Inches
import requests
import os
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

def get_text_from_huggingface(prompt):
    API_URL = "https://api-inference.huggingface.co/models/t5-small"

    
    HF_API_KEY = os.getenv("HF_API_KEY")

    if not HF_API_KEY:
        print("Error: API key is missing.")
    else:
        print("API Key Loaded Successfully:", HF_API_KEY[:5] + "****")  # Masking for security

    headers = {"Authorization": f"Bearer {HF_API_KEY}"}
    payload = {"inputs": prompt}

    try:
        response = requests.post(API_URL, headers=headers, json=payload, timeout=10)
        response.raise_for_status()  # Raise error for bad responses (4xx, 5xx)
        summary = response.json()
        return summary[0].get("summary_text", "Error: Unexpected response format.")
    except requests.exceptions.RequestException as e:
        return f"Error: {str(e)}"

def add_title_slide(prs, title, subtitle):
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = title
    if slide.placeholders and len(slide.placeholders) > 1:
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
    
    add_text_box_slide(prs, f"About {topic}", description)
    
    filename = "generated_presentation.pptx"
    prs.save(filename)
    return filename

st.title("📊 Generate PowerPoint Presentation")
topic = st.text_input("Enter Topic:")

if st.button("Generate PPT"):
    if topic:
        pptx_file = generate_presentation(topic)
        with open(pptx_file, "rb") as file:
            st.download_button(
                label="📥 Download Presentation", 
                data=file, 
                file_name=pptx_file, 
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
        os.remove(pptx_file)  # Cleanup after download
    else:
        st.warning("⚠️ Please enter a topic.")
