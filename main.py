'''Ensure to install the following libraries:
- requests
- pandas
- selenium
Download a web driver and set up the path to the driver.
The attached file should be named 'python-assignment.docx' '''

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import pandas as pd
import time
from docx import Document


# Path to your ChromeDriver (change to match your system)
chromedriver_path = "C:\\WebDriver\\chromedriver.exe"

# Setup Selenium WebDriver with Chrome
chrome_options = Options()
chrome_options.add_argument("--headless")  
service = Service(chromedriver_path)
driver = webdriver.Chrome(service=service, options=chrome_options)

# Read the DOCX file to extract URLs
doc = Document("python-assignment.docx")
urls = []
for paragraph in doc.paragraphs:
    text = paragraph.text.strip()
    if text.startswith("http"):
        urls.append(text)

# Data structure to hold all the scraped data
scraped_data = []

# Scrape each URL
for url in urls:
    try:
        driver.get(url)  # Open the URL
        time.sleep(4)  # Wait for the page to load

        #Extract title, all text, and all links
        title = driver.title
        body_text = driver.find_element(By.TAG_NAME, "body").text
        links = [element.get_attribute("href") for element in driver.find_elements(By.TAG_NAME, "a")]

        # Store the data
        scraped_data.append({
            "URL": url,
            "Title": title,
            "Body Text": body_text[:400],  
            "Links": links
        })
    except Exception as e:
        print(f"Error scraping {url}: {e}")

# Convert scraped data to DataFrame
df = pd.DataFrame(scraped_data)

# Export to Excel
output_path = "/path/to/scraped_data.xlsx"
df.to_excel(output_path, index=False)

# Close the Selenium driver
driver.quit()

print(f"Data saved to {output_path}")
