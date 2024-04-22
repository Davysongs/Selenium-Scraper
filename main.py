import pandas as pd
from docx import Document

# Load the DOCX file
file_path = 'python-assignment.docx'
doc = Document(file_path)

# Prepare data structures to store the text and links
data = []  # List of dictionaries to create a DataFrame

# Parse the document
current_title = None

for para in doc.paragraphs:
    text = para.text.strip()
    if not text:
        continue  # Skip empty paragraphs

    if not current_title and "http" not in text:
        # If the paragraph doesn't contain a URL and we don't have a title yet, set it as the current title
        current_title = text

    elif "http" in text:
        # If the paragraph contains URLs, extract them
        urls = [word for word in text.split() if "http" in word]
        for url in urls:
            # Ensure each URL has an associated title
            data.append({"Title": current_title, "Link": url})

# Create a DataFrame to store the data
df = pd.DataFrame(data)

# Save to an Excel file
excel_file_path = 'scraped_data.xlsx'
df.to_excel(excel_file_path, index=False)

print(f"Data has been extracted and stored in {excel_file_path}")
