import streamlit as st
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from pptx import Presentation
from pptx.util import Inches
import chromedriver_autoinstaller
import time

def fetch_pricing_data_with_selenium():
    """Fetch pricing data from Glama AI using Selenium."""
    # Automatically download and install ChromeDriver
    chromedriver_autoinstaller.install()

    # Set up Chrome options to run in headless mode (no UI)
    chrome_options = Options()
    chrome_options.add_argument("--headless")  # Ensure the browser window doesn't pop up
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")

    # Start the WebDriver
    driver = webdriver.Chrome(options=chrome_options)

    # Open the Glama AI pricing page
    url = "https://glama.ai/pricing"
    driver.get(url)

    # Wait for the page to fully load
    time.sleep(5)

    # Now we extract the pricing data
    pricing_data = []
    try:
        plans = driver.find_elements(By.TAG_NAME, "h3")
        prices = driver.find_elements(By.CLASS_NAME, "price")
        descriptions = driver.find_elements(By.TAG_NAME, "p")

        # Make sure to loop through all the extracted elements and store them in a list
        for plan, price, description in zip(plans, prices, descriptions):
            pricing_data.append([plan.text.strip(), description.text.strip(), price.text.strip()])
    except Exception as e:
        print(f"Error during data extraction: {e}")
    finally:
        driver.quit()  # Close the browser

    return pricing_data

def generate_presentation(pricing_data):
    """Generate PowerPoint presentation with extracted pricing data."""
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = "Glama AI Pricing Plans"
    slide.placeholders[1].text = "Generated using Streamlit & Selenium"

    for plan, desc, price in pricing_data:
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        slide.shapes.title.text = plan
        textbox = slide.shapes.add_textbox(Inches(1), Inches(1.5), Inches(6), Inches(4))
        textbox.text_frame.text = f"{desc}\nPrice: {price}"

    filename = "Glama_AI_Pricing.pptx"
    prs.save(filename)
    return filename

# Streamlit UI
st.title("📊 Glama AI Pricing PPT Generator (Selenium Version)")

if st.button("Fetch & Generate PPT"):
    pricing_data = fetch_pricing_data_with_selenium()
    
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
