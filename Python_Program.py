import tkinter as tk
from tkinter import filedialog
import shutil
import subprocess
import sys
import os
import openpyxl
from docx import Document

# Function to browse for Word and Excel files


# Function to browse for Word and Excel files
def browse_files():
    word_file_path = filedialog.askopenfilename(
        title="Select a Word file", filetypes=[("Word Files", "*.docx")])
    excel_file_path = filedialog.askopenfilename(
        title="Select an Excel file", filetypes=[("Excel Files", "*.xlsx")])

    if word_file_path and excel_file_path:
        # Move the selected files to the script's directory
        script_directory = os.path.dirname(os.path.realpath(__file__))
        word_file_name = os.path.basename(word_file_path)
        excel_file_name = os.path.basename(excel_file_path)

        # Check if the files already exist in the destination directory
        if not os.path.exists(os.path.join(script_directory, word_file_name)):
            shutil.copy(word_file_path, os.path.join(
                script_directory, word_file_name))
        if not os.path.exists(os.path.join(script_directory, excel_file_name)):
            shutil.copy(excel_file_path, os.path.join(
                script_directory, excel_file_name))

        # Perform the three Python scripts in order
        process_excel_file(os.path.join(script_directory, excel_file_name))
        third_script()
        process_word_file(os.path.join(script_directory, word_file_name))


def process_word_file(word_file_path):
    doc = Document(word_file_path)
    python_script = "edunextattempt3.py"  # Replace with the actual filename
    command = [sys.executable, python_script, word_file_path]
    subprocess.run(command)
    try:
        completed_process = subprocess.run(
            ["python", python_script],    # Command to run the script
            stdout=subprocess.PIPE,      # Capture standard output
            stderr=subprocess.PIPE,      # Capture standard error
            text=True                    # Decode output to text
        )
        if completed_process.returncode == 0:
            print("Script executed successfully.")
            print("Standard Output:")
            print(completed_process.stdout)
        else:
            print("Script encountered an error. Standard Error:")
            print(completed_process.stderr)
    except FileNotFoundError:
        print(f"Python executable not found. Make sure Python is installed and in your system's PATH.")
    except Exception as e:
        print(f"An error occurred: {str(e)}")


# Function to process the Excel file


def process_excel_file(excel_file_path):
    wb = openpyxl.load_workbook(excel_file_path)
    python_script = "Excel work.py"  # Replace with the actual filename
    command = [sys.executable, python_script, excel_file_path]
    subprocess.run(command)
    try:
        completed_process = subprocess.run(
            ["python", python_script],    # Command to run the script
            stdout=subprocess.PIPE,      # Capture standard output
            stderr=subprocess.PIPE,      # Capture standard error
            text=True                    # Decode output to text
        )
        if completed_process.returncode == 0:
            print("Script executed successfully.")
            print("Standard Output:")
            print(completed_process.stdout)
        else:
            print("Script encountered an error. Standard Error:")
            print(completed_process.stderr)
    except FileNotFoundError:
        print(f"Python executable not found. Make sure Python is installed and in your system's PATH.")
    except Exception as e:
        print(f"An error occurred: {str(e)}")


def third_script():
    python_script = "deletenulljson.py"  # Replace with the actual filename
    try:
        completed_process = subprocess.run(
            ["python", python_script],    # Command to run the script
            stdout=subprocess.PIPE,      # Capture standard output
            stderr=subprocess.PIPE,      # Capture standard error
            text=True                    # Decode output to text
        )
        if completed_process.returncode == 0:
            print("Script executed successfully.")
            print("Standard Output:")
            print(completed_process.stdout)
        else:
            print("Script encountered an error. Standard Error:")
            print(completed_process.stderr)
    except FileNotFoundError:
        print(f"Python executable not found. Make sure Python is installed and in your system's PATH.")
    except Exception as e:
        print(f"An error occurred: {str(e)}")


# Create the GUI window
window = tk.Tk()
window.title("File Selection and Processing")

# Create a button to browse for files
browse_button = tk.Button(
    window, text="Browse for Word and Excel Files", command=browse_files)
browse_button.pack()

# Start the GUI application
window.mainloop()
