import json
import sys
from docx import Document
from docx2pdf import convert  # Import the convert function from python-docx2pdf

# Read the JSON file
with open('data.json', 'r') as json_file:
    data = json.load(json_file)

# Open the Word document
data_received = sys.argv[1]
doc = Document(data_received)
max_length = max(len(data[key]) for key in data)
for i in range(max_length):
    for paragraph in doc.paragraphs:
        for key, value in data.items():
            if key in paragraph.text:
                start = paragraph.text.index(key)
                paragraph.text = paragraph.text[:start] + \
                    key + "- " + str(value[i])
    filename = 'output_word_document'+str((i+1))+'.docx'
    doc.save(filename)
    convert(filename)
    # Convert the Word document to PDF
    # Convert the first generated Word document to PDF
