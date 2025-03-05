import streamlit as st
import requests
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.util import Inches
import os
import re

GLAMA_AI_URL = "https://glama.ai/pricing"

def fetch_pricing_data():
    """Fetch pricing details from Glama AI's website."""
    try:
        response = requests.get(GLAMA_AI_URL, timeout=10)
        if response.status_code == 200:
            text = response.text

            # Extract pricing details using regex (since we avoid BeautifulSoup)
            plans = re.findall(r'>(Starter|Pro|Business)<', text)
            prices = re.findall(r'\$(\d+)[^<]*', text)  # Extracts prices like $26, $80
            descriptions = re.findall(r'>(For .*?)<', text)  # Extract plan descriptions
            features = re.findall(r'<li>(.*?)</li>', text)  # Extract feature list items

            if len(plans) != len(prices) or len(plans) != len(descriptions):
                return None  # In case the format is different

            # Organize extracted details into a structured list
            pricing_data = []
            for i in range(len(plans)):
                plan = plans[i]
                price = f"${prices[i]}" if i < len(prices) else "Free"
                description = descriptions[i]
                feature_list = ", ".join(features[i*5:(i+1)*5])  # Assuming 5 features per plan
                pricing_data.append([plan, description, price, feature_list])

            return pricing_data
        else:
            return None
    except Exception as e:
        st.error(f"Error fetching pricing data: {e}")
        return None

def add_title_slide(prs, title, subtitle):
    """Add a title slide with a title and subtitle."""
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = title
    slide.placeholders[1].text = subtitle

def add_text_box_slide(prs, title, text):
    """Add a slide with a text box."""
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = title
    textbox = slide.shapes.add_textbox(Inches(1), Inches(1.5), Inches(6), Inches(4))
    textbox.text_frame.text = text

def add_table_slide(prs, title, data):
    """Add a table slide with provided data."""
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = title

    rows, cols = len(data) + 1, len(data[0])
    left, top, width, height = Inches(1), Inches(1.5), Inches(8), Inches(0.8 + 0.5 * rows)

    table = slide.shapes.add_table(rows, cols, left, top, width, height).table

    # Set column widths
    for i in range(cols):
        table.columns[i].width = Inches(2)

    # Set headers
    headers = ["Plan", "Description", "Price", "Features"]
    for col_idx, header in enumerate(headers):
        table.cell(0, col_idx).text = header

    # Add data to table
    for row_idx, row_data in enumerate(data, start=1):
        for col_idx, value in enumerate(row_data):
            table.cell(row_idx, col_idx).text = value

def add_chart_slide(prs, title, categories, values):
    """Add a bar chart slide."""
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = title

    chart_data = CategoryChartData()
    chart_data.categories = categories
    chart_data.add_series('Price (USD)', values)

    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(1), Inches(1.5), Inches(8), Inches(4.5), chart_data
    ).chart
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM

def add_image_slide(prs, title, img_path):
    """Add a slide with an image."""
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = title
    slide.shapes.add_picture(img_path, Inches(1), Inches(1.5), width=Inches(6), height=Inches(4))

def generate_presentation(topic, description, pricing_data, image_file=None):
    """Generate a PowerPoint presentation based on fetched data."""
    prs = Presentation()
    add_title_slide(prs, topic, "Generated using Streamlit & python-pptx")

    add_text_box_slide(prs, "Introduction", description)

    if pricing_data:
        add_table_slide(prs, "Glama AI Pricing Plans", pricing_data)

        # Bar Chart Data (exclude free plan)
        categories = [plan[0] for plan in pricing_data if plan[2] != "Free"]
        values = [int(plan[2][1:]) for plan in pricing_data if plan[2] != "Free"]
        add_chart_slide(prs, "Pricing Comparison", categories, values)

    if image_file:
        add_image_slide(prs, "Uploaded Image", image_file)

    filename = "glama_ai_presentation.pptx"
    prs.save(filename)
    return filename

# Streamlit UI
st.title("📊 Glama AI PowerPoint Generator (Live Data)")

topic = st.text_input("Enter Topic:", "Glama AI Overview")
description = st.text_area("Enter Description:", "An overview of Glama AI's latest pricing plans.")
uploaded_image = st.file_uploader("Upload an image (optional)", type=["png", "jpg", "jpeg"])

if st.button("Fetch & Generate PPT"):
    st.info("Fetching latest pricing details...")
    pricing_data = fetch_pricing_data()

    if pricing_data:
        st.success("Pricing data successfully retrieved!")
        img_path = None
        if uploaded_image:
            img_path = f"temp_{uploaded_image.name}"
            with open(img_path, "wb") as f:
                f.write(uploaded_image.getbuffer())

        pptx_file = generate_presentation(topic, description, pricing_data, img_path)

        with open(pptx_file, "rb") as file:
            st.download_button(
                label="📥 Download Presentation",
                data=file,
                file_name=pptx_file,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            )

        if img_path:
            os.remove(img_path)  # Cleanup temp file
    else:
        st.error("Failed to fetch pricing data. Please check the website or try again later.")
