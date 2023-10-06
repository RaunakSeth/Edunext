import docx
import json
# Define a function to extract formatting information from a run


def extract_formatting(run):
    if run is None:
        return {}

    formatting_info = {
        'bold': run.bold,
        'italic': run.italic,
        'underline': run.underline,
        'font_size': run.font.size,
        'font_name': run.font.name,
        'color': run.font.color.rgb if run.font.color is not None else None,
    }
    return formatting_info


filename = "data.json"

with open(filename, "r") as json_file:
    data = json.load(json_file)


# Open the Word document
doc = docx.Document('ds_project_Synopsis[1].docx')

# Initialize variables to store current run's text and formatting
current_run_text = ""
current_run_formatting = None

# Iterate through paragraphs in the document
for paragraph in doc.paragraphs:
    for run in paragraph.runs:
        # Check if run.text is non-empty (ignoring whitespace)
        if run.text.strip():
            # Check if formatting changes, start a new run
            if current_run_formatting is not None and current_run_formatting != extract_formatting(run):
                # print("Text:", current_run_text)
                # print("Formatting:", current_run_formatting)
                # print("-------------------------")

                current_run_text = ""
            current_run_text += run.text
            current_run_formatting = extract_formatting(run)

    # Print the last run within the paragraph
    # if current_run_text.strip():
    #     print("Text:", current_run_text)
    #     print("Formatting:", current_run_formatting)
    #     print("-------------------------")

    # Reset variables for the next paragraph
    current_run_text = ""
    current_run_formatting = None
