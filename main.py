import os
import win32com.client as win32 # pip install pywin32
import sys
import comtypes.client # pip install comtypes

def ensure_dir_recursive(dir_path):
    if not os.path.exists(dir_path):
        os.makedirs(dir_path)

def convertFromDocx(input, output):
    if not os.path.exists(input):
        print("Input file does not exist")
        sys.exit(1)
    ensure_dir_recursive(os.path.dirname(output))
    wdFormatPDF = 17

    input = os.path.abspath(input)
    output = os.path.abspath(output)

    word = win32.Dispatch('Word.Application')
    doc = word.Documents.Open(input)
    doc.SaveAs(output, FileFormat=wdFormatPDF)
    doc.Close()
    word.Quit()

def convertFromPptx(input_file_path, output_file_path):
    if not os.path.exists(input_file_path):
        print("Input file does not exist")
        sys.exit(1)
    ensure_dir_recursive(os.path.dirname(output_file_path))
    input_file_path = os.path.abspath(input_file_path)
    output_file_path = os.path.abspath(output_file_path)
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = 1
    slides = powerpoint.Presentations.Open(input_file_path)
    slides.SaveAs(output_file_path, 32)
    slides.Close()
    powerpoint.Quit()

try:
    if len(sys.argv) != 4:
        print("Usage: main.exe <input_file_path> <output_file_path> <type>")
        sys.exit(1)
    input_file_path = sys.argv[1]
    output_file_path = sys.argv[2]
    file_type = sys.argv[3]
    if file_type == "docx":
        convertFromDocx(input_file_path, output_file_path)
    elif file_type == "pptx":
        convertFromPptx(input_file_path, output_file_path)
    else:
        print("Type not supported. Supported types: docx, pptx")
        sys.exit(1)
except Exception as e:
    print("Error: " + str(e))
    sys.exit(1)