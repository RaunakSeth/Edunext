import json
from docx import Document

# Read the JSON file
with open('data.json', 'r') as json_file:
    data = json.load(json_file)

# Open the Word document
doc = Document('ds_project_Synopsis[1].docx')
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
