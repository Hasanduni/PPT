import streamlit as st
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.util import Inches
import os

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

def add_bullet_points_slide(prs, title, text):
    """Add a bullet points slide based on the description."""
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = title

    if not slide.placeholders or len(slide.placeholders) < 2:
        st.error("Error: The selected slide layout does not have a content placeholder.")
        return

    content = slide.placeholders[1].text_frame
    if content is None:
        st.error("Error: No text frame found in the placeholder.")
        return

    points = text.split('. ')
    for point in points:
        if point.strip():
            content.add_paragraph(point.strip())


def add_table_slide(prs, title, data):
    """Add a sample table slide."""
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = title

    rows, cols = len(data) + 1, len(data[0])
    left, top, width, height = Inches(1), Inches(1.5), Inches(6), Inches(3)
    table = slide.shapes.add_table(rows, cols, left, top, width, height).table

    # Set column headers dynamically
    for j, header in enumerate(data[0].keys()):
        table.cell(0, j).text = header

    # Add data rows
    for i, row in enumerate(data, start=1):
        for j, (key, value) in enumerate(row.items()):
            table.cell(i, j).text = str(value)

def add_image_slide(prs, title, img_path):
    """Add a slide with an image."""
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = title
    slide.shapes.add_picture(img_path, Inches(1), Inches(1.5), width=Inches(6), height=Inches(4))

def generate_presentation(topic, description, image_file=None):
    """Generate a PowerPoint presentation based on user input."""
    prs = Presentation()
    add_title_slide(prs, topic, "Generated using Streamlit & python-pptx")

    add_text_box_slide(prs, "Introduction", description)
    add_bullet_points_slide(prs, f"Key Points about {topic}", description)
    add_chart_slide(prs, "Bar Chart Representation")
    add_pie_chart_slide(prs, "Pie Chart Breakdown")

    # Sample table data
    table_data = [
        {"Category": "A", "Value": 100, "Percentage": "25%"},
        {"Category": "B", "Value": 150, "Percentage": "37.5%"},
        {"Category": "C", "Value": 150, "Percentage": "37.5%"},
    ]
    add_table_slide(prs, "Data Table", table_data)

    if image_file:
        add_image_slide(prs, "Uploaded Image", image_file)

    filename = "generated_presentation.pptx"
    prs.save(filename)
    return filename

# Streamlit UI
st.title("📊 PowerPoint Generator")

topic = st.text_input("Enter Topic:")
description = st.text_area("Enter Description:")
uploaded_image = st.file_uploader("Upload an image (optional)", type=["png", "jpg", "jpeg"])

if st.button("Generate PPT"):
    if topic and description:
        img_path = None
        if uploaded_image:
            img_path = f"temp_{uploaded_image.name}"
            with open(img_path, "wb") as f:
                f.write(uploaded_image.getbuffer())

        pptx_file = generate_presentation(topic, description, img_path)

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
        st.warning("⚠️ Please enter both a topic and description.")
