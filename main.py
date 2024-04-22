import pandas as pd
from docx import Document

# Load the DOCX file
file_path = 'python-assignment.docx'
doc = Document(file_path)

# Prepare data structures to store the text and links
titles = []
links = []

# Parse the document
for para in doc.paragraphs:
    text = para.text.strip()
    if text:
        # If the paragraph contains a URL, extract it
        if "http" in text:
            # Find all URLs in the text (simple heuristic)
            urls = [word for word in text.split() if "http" in word]
            links.extend(urls)
            # If it's a title followed by links, store them in separate lists
            title = text.split()[0]  # Assuming the first word is the title
            titles.extend([title] * len(urls))
        else:
            # If the paragraph is not a link, consider it as a title
            titles.append(text)

# Create a DataFrame to store the data
df = pd.DataFrame({
    'Title': titles,
    'Link': links
})

# Save to an Excel file
excel_file_path = 'scraped_data.xlsx'
df.to_excel(excel_file_path, index=False)

print(f"Data has been extracted and stored in {excel_file_path}")
