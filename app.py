import streamlit as st
import requests
from pptx import Presentation
from pptx.util import Inches
from bs4 import BeautifulSoup

def fetch_pricing_data():
    """Fetch pricing data using ScraperAPI."""
    SCRAPER_API_KEY = "7b7d6359172aa8d26d022034260b0089"  # Use your actual ScraperAPI key
    GLAMA_AI_URL = "https://glama.ai/pricing"
    
    # ScraperAPI URL format
    api_url = f"https://api.scraperapi.com?api_key={SCRAPER_API_KEY}&url={GLAMA_AI_URL}&render=true"
    
    try:
        response = requests.get(api_url, timeout=10)
        print(f"Response Status Code: {response.status_code}")  # Check the status code
        print(f"Response Headers: {response.headers}")  # Print response headers
        print(f"Response Text: {response.text[:500]}")  # Print first 500 characters of response text for debugging

        if response.status_code == 200:
            return extract_data_from_html(response.text)  # Replace this with your actual data extraction function
        else:
            print(f"Failed with status code: {response.status_code}")
            return None
    except Exception as e:
        print(f"Error: {e}")
        return None

def extract_data_from_html(html):
    """Extract pricing data from HTML using simple parsing."""
    soup = BeautifulSoup(html, "html.parser")

    pricing_data = []
    plans = soup.find_all("h3")
    prices = soup.find_all("span", class_="price")
    descriptions = soup.find_all("p")
    
    for i in range(min(len(plans), len(prices), len(descriptions))):
        plan = plans[i].get_text(strip=True)
        price = prices[i].get_text(strip=True)
        description = descriptions[i].get_text(strip=True)
        pricing_data.append([plan, description, price])

    return pricing_data

def generate_presentation(pricing_data):
    """Generate PowerPoint presentation with extracted pricing data."""
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = "Glama AI Pricing Plans"
    slide.placeholders[1].text = "Generated using Streamlit & ScraperAPI"

    for plan, desc, price in pricing_data:
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        slide.shapes.title.text = plan
        textbox = slide.shapes.add_textbox(Inches(1), Inches(1.5), Inches(6), Inches(4))
        textbox.text_frame.text = f"{desc}\nPrice: {price}"

    filename = "Glama_AI_Pricing.pptx"
    prs.save(filename)
    return filename

# Streamlit UI
st.title("📊 Glama AI Pricing PPT Generator (ScraperAPI Version)")

if st.button("Fetch & Generate PPT"):
    pricing_data = fetch_pricing_data()
    
    if pricing_data:
        pptx_file = generate_presentation(pricing_data)

        with open(pptx_file, "rb") as file:
            st.download_button(
                label="📥 Download Presentation",
                data=file,
                file_name=pptx_file,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            )
    else:
        st.error("Failed to fetch pricing data. Please check the website or try again later.")
