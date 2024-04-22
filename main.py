import docx
from openpyxl import Workbook

def scrape_docx_to_excel(docx_file, excel_file):
  """
  Extracts text, formatting, and embedded objects from a DOCX file
  and stores them in a structured Excel spreadsheet.

  Args:
      docx_file (str): Path to the DOCX file.
      excel_file (str): Path to the output Excel file.
  """

  # Open DOCX file
  doc = docx.Document(docx_file)

  # Create Excel workbook and worksheet
  wb = Workbook()
  ws = wb.active

  # Header row for data fields (customize as needed)
  ws.append(["Text", "Bold", "Italic", "Underline", "Image Path (if embedded)"])

  # Iterate through paragraphs and extract relevant information
  for paragraph in doc.paragraphs:
    text = paragraph.text.strip()

    # Check for potential changes in font attribute access
    if hasattr(paragraph, 'style'):  # Check if style attribute exists
      font = paragraph.style.font  # Access font from style if available
    else:
      font = paragraph.runs[0].font  # Access font from first run if no style

    is_bold = font.bold
    is_italic = font.italic
    is_underline = font.underline
    image_path = None  # Initialize image path

    # Check for inline objects (images) using parent element
    for inline in paragraph._element.inline_objects:  # Access inline objects from parent element
      if inline.type == docx.inlineobject.INLINEOBJ_TYPE.PICTURE:
        image_path = inline.properties.content.image_data.filename  # Extract image path

    # Append data to Excel sheet
    ws.append([text, is_bold, is_italic, is_underline, image_path])

  # Save Excel file
  wb.save(excel_file)

# Example usage
docx_file = "python-assignment.docx"
excel_file = "extracted_data.xlsx"
scrape_docx_to_excel(docx_file, excel_file)

print("Data extracted from DOCX and saved to Excel file.")
