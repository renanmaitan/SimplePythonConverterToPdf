import os
import win32com.client as win32 # pip install pywin32
import sys
import comtypes.client # pip install comtypes

def ensure_dir_recursive(dir_path):
    if not os.path.exists(dir_path):
        os.makedirs(dir_path)

def convertFromDocx(input, output):
    ensure_dir_recursive(os.path.dirname(output))
    ensure_dir_recursive(os.path.dirname(input))
    wdFormatPDF = 17

    input = os.path.abspath(input)
    output = os.path.abspath(output)

    word = win32.Dispatch('Word.Application')
    doc = word.Documents.Open(input)
    doc.SaveAs(output, FileFormat=wdFormatPDF)
    doc.Close()
    word.Quit()

def convertFromPptx(input_file_path, output_file_path):
    ensure_dir_recursive(os.path.dirname(output_file_path))
    ensure_dir_recursive(os.path.dirname(input_file_path))
    input_file_path = os.path.abspath(input_file_path)
    output_file_path = os.path.abspath(output_file_path)
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = 1
    slides = powerpoint.Presentations.Open(input_file_path)
    slides.SaveAs(output_file_path, 32)
    slides.Close()
    powerpoint.Quit()

if __name__ == '__main__':
    if len(sys.argv) < 4:
        print("Usage: python convert.py <input-file> <output-file> <type>")
        sys.exit(1)
    
    input_file_path = sys.argv[1]
    output_file_path = sys.argv[2]
    type = sys.argv[3]

    if type == "docx":
        convertFromDocx(input_file_path, output_file_path)
    elif type == "pptx":
        convertFromPptx(input_file_path, output_file_path)
    else:
        print("Type not supported. Supported types: docx, pptx")
        sys.exit(1)