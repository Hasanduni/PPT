import streamlit as st
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.util import Inches
import openai
import os

# Define your OpenAI API Key
OPENAI_API_KEY = ""  # Replace with your OpenAI API Key

# Initialize OpenAI client
client = openai.Client(api_key=OPENAI_API_KEY)

# Function to generate text from OpenAI API
def get_text_from_openai(prompt):
    try:
        response = client.chat.completions.create(
            model="gpt-3.5-turbo",  # Use "gpt-4" if needed
            messages=[{"role": "user", "content": prompt}],
            max_tokens=150
        )
        return response.choices[0].message.content
    except Exception as e:
        return f"Error: {str(e)}"

# Function to add title slide
def add_title_slide(prs, title, subtitle):
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = title
    if slide.placeholders and len(slide.placeholders) > 1:
        slide.placeholders[1].text = subtitle

# Function to add a chart slide
def add_chart_slide(prs, title, chart_data):
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = title
    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(1), Inches(1.5), Inches(6), Inches(4.5), chart_data
    ).chart
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM

# Function to add text box slide
def add_text_box_slide(prs, title, text):
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = title
    textbox = slide.shapes.add_textbox(Inches(1), Inches(1.5), Inches(6), Inches(4))
    textbox.text_frame.text = text

# Function to generate a PowerPoint presentation
def generate_presentation(topic):
    prs = Presentation()
    add_title_slide(prs, topic, "Generated using Streamlit and python-pptx")
    
    # Generate text summary using OpenAI
    description = get_text_from_openai(f"Generate a brief summary about {topic}")
    
    # Sample chart data
    chart_data = CategoryChartData()
    chart_data.categories = ['A', 'B', 'C']
    chart_data.add_series('Series 1', (10, 20, 30))
    add_chart_slide(prs, "Sample Chart", chart_data)
    
    add_text_box_slide(prs, f"About {topic}", description)
    
    filename = "generated_presentation.pptx"
    prs.save(filename)
    return filename

# Streamlit UI
st.title("📊 Generate PowerPoint Presentation with OpenAI")
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
