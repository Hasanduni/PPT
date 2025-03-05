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

# Check if OpenAI API key exists, otherwise fallback to GPT-2 or Hugging Face
openai_api_key = os.getenv('OPENAI_API_KEY')

# Define function for generating content
def generate_api_content(topic, description):
    """Fetch detailed content for the topic and description."""
    if openai_api_key:
        import openai
        try:
            # Use OpenAI API for content generation
            prompt = f"Generate a detailed explanation for the topic: {topic}\nDescription: {description}"
            response = openai.ChatCompletion.create(
                model="gpt-3.5-turbo",  # You can adjust this to use GPT-4 if needed
                messages=[{"role": "system", "content": "You are a helpful assistant."},
                          {"role": "user", "content": prompt}],
                max_tokens=150,
                temperature=0.7,
            )
            return response['choices'][0]['message']['content']
        except openai.error.OpenAIError as e:
            return f"OpenAI request failed: {e}"
    else:
        # Fallback: GPT-2 model or other text generation methods (e.g., Hugging Face API)
        return f"Generated content for {topic} based on description: {description}"

def add_title_slide(prs, title, subtitle):
    """Add a title slide with a title and subtitle."""
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = title
    slide.placeholders[1].text = subtitle

def add_chart_slide(prs, title):
    """Add a sample bar chart slide."""
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = title

    chart_data = CategoryChartData()
    chart_data.categories = ['A', 'B', 'C']
    chart_data.add_series('Series 1', (10, 20, 30))

    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(1), Inches(1.5), Inches(6), Inches(4.5), chart_data
    ).chart
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM

def add_pie_chart_slide(prs, title):
    """Add a sample pie chart slide."""
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = title

    chart_data = CategoryChartData()
    chart_data.categories = ['X', 'Y', 'Z']
    chart_data.add_series('Series 1', (40, 30, 30))

    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.PIE, Inches(1), Inches(1.5), Inches(6), Inches(4.5), chart_data
    ).chart
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.RIGHT

def add_text_box_slide(prs, title, text):
    """Add a slide with a text box."""
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = title
    textbox = slide.shapes.add_textbox(Inches(1), Inches(1.5), Inches(6), Inches(4))
    textbox.text_frame.text = text

def add_image_slide(prs, title, img_path):
    """Add a slide with an image."""
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = title
    slide.shapes.add_picture(img_path, Inches(1), Inches(1.5), width=Inches(6), height=Inches(4))

def add_table_slide(prs, title):
    """Add a slide with a table."""
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = title

    # Create a table with 3 rows and 4 columns
    rows = 3
    cols = 4
    left = Inches(1)
    top = Inches(1.5)
    width = Inches(6)
    height = Inches(3)

    table = slide.shapes.add_table(rows, cols, left, top, width, height).table

    # Set column widths
    for col in range(cols):
        table.columns[col].width = Inches(1.5)

    # Populate the table with sample data
    data = [
        ["Header 1", "Header 2", "Header 3", "Header 4"],
        ["Row 1, Col 1", "Row 1, Col 2", "Row 1, Col 3", "Row 1, Col 4"],
        ["Row 2, Col 1", "Row 2, Col 2", "Row 2, Col 3", "Row 2, Col 4"]
    ]

    for r
