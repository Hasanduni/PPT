import streamlit as st
import requests
from pptx import Presentation
from pptx.util import Inches
from bs4 import BeautifulSoup

# Function to fetch pricing data using ScraperAPI
def fetch_pricing_data():
    SCRAPER_API_KEY = "7b7d6359172aa8d26d022034260b0089"  # Replace with your actual ScraperAPI key
    GLAMA_AI_URL = "https://glama.ai/pricing"
    
    api_url = f"https://api.scraperapi.com?api_key={SCRAPER_API_KEY}&url={GLAMA_AI_URL}&render=true&premium=true"
    
    try:
        response = requests.get(api_url, timeout=10)
        
        # Print debugging details
        print("Status Code:", response.status_code)  # Debugging step to check status code
        print("Response Text (First 500 chars):", response.text[:500])  # Print first 500 characters
        
        if response.status_code == 200:
            # Check if the data is valid and contains expected HTML structure
            if "pricing" in response.text:
                return extract_data_from_html(response.text)  # Proceed with extraction if valid
            else:
                print("Error: Pricing data not found in response")
                return None
        else:
            print(f"Failed with status code: {response.status_code}")
            return None
    except Exception as e:
        print(f"Error during request: {e}")
        return None

# Function to extract pricing data from the HTML response using BeautifulSoup
def extract_data_from_html(html):
    soup = BeautifulSoup(html, "html.parser")
    pricing_data = []
    
    # Check if we can find the expected elements
    plans = soup.find_all("h3")
    prices = soup.find_all("span", class_="price")
    descriptions = soup.find_all("p")
    
    if not plans or not prices or not descriptions:
        print("Error: Couldn't find the expected HTML elements for pricing")
        return None
    
    for i in range(min(len(plans), len(prices), len(descriptions))):
        plan = plans[i].get_text(strip=True)
        price = prices[i].get_text(strip=True)
        description = descriptions[i].get_text(strip=True)
        pricing_data.append([plan, description, price])

    return pricing_data

# Function to generate a PowerPoint presentation with extracted data
def generate_presentation(pricing_data):
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
st.title("📊 Glama AI Pricing PPT Generator")

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
