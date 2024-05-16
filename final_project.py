import tkinter as tk
from tkinter import filedialog
from pdf2image import convert_from_path
import cv2
import fitz 
import os
import pytesseract
import re
import csv
import json
import pandas as pd
from PIL import Image
from tqdm import tqdm

def extract_text_from_image(image_path):  
    if not os.path.exists(image_path):
        print(f"Error: Image file '{image_path}' not found.")
        return ""
    myconfig = r"--psm 11 --oem 3"
    img = cv2.imread(image_path)
    if img is None:
        print(f"Error: Unable to read image '{image_path}'.")
        return ""
    gray_im = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    text = pytesseract.image_to_string(gray_im, config=myconfig, lang="eng")
    return str(text).strip()

def parse_text(text):
    patterns = {
        'Date': r'((\d{2})(\d{2})(\d{4}))',
        'IFC code': r'(\w+\w+\w+\w+\d{6,7})',
        'Account No.': r'\n(\d{12,16})',
        'Cheque No.': r'"(\d+\s\d+)',
        'Bank Name': r'\n\n([A-Za-z\s]+ Bank)',
        'paying to': r'pay (\w+\s+\w+)',
        'Amount': r'< (\d+)',
        'currency': r'Rupees ([A-Za-z\s]+)\s+Only'
    }

    parsed_data = {}
    for key, pattern in patterns.items():
        match = re.search(pattern, text)
        parsed_data[key] = match.group(1) if match else None

    return parsed_data

def write_to_csv(filename, data):
    with open(filename, mode='w', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(['Date', 'IFC code', 'Account No', 'Cheque No', 'Bank name', 'pay to', 'Amount', 'currency'])
        writer.writerow(data)

def write_to_json(filename, data):
    with open(filename, 'w') as file:
        json.dump(data, file)

def write_to_excel(filename, data):
    df = pd.DataFrame(data)  
    df.to_excel(filename, index=False)

def extract_data(pdf_path, output_folder):
    try:
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)
   
        pdf_document = convert_from_path(pdf_path, poppler_path=r'C:\Program Files\poppler-24.02.0\Library\bin')
        extracted_data = []
                
        for i, image in enumerate(pdf_document):
            image_path = os.path.join(output_folder, f"page_{i+1}.png")
            image.save(image_path, "PNG")

            text = extract_text_from_image(image_path)
            parsed_data = parse_text(text)
            extracted_data.append(parsed_data)  # Append each page's data to the list
        
        return extracted_data
    except Exception as e:
        print(f"Error during extraction: {e}")
        return []

def select_pdf_file():
    file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    pdf_path_entry.delete(0, tk.END)
    pdf_path_entry.insert(0, file_path)

def select_output_folder():
    folder_path = filedialog.askdirectory()
    output_folder_entry.delete(0, tk.END)
    output_folder_entry.insert(0, folder_path)

def extract_and_export_data():
    try:
        pdf_path = pdf_path_entry.get()
        output_folder = output_folder_entry.get()
        output_format = format_var.get()  # Get the selected output format
        
        extracted_data = extract_data(pdf_path, output_folder)
        
        if not extracted_data:
            result_label.config(text="No data extracted.", fg="red")
            return
        
        output_filename = os.path.join(output_folder, f'extracted_data.{output_format.lower()}')
        
        if output_format == "CSV":
            write_to_csv(output_filename, extracted_data[0].values())  # Pass first page's data values
        elif output_format == "JSON":
            write_to_json(output_filename, extracted_data)
        elif output_format == "Excel":
            output_filename = os.path.join(output_folder, f'extracted_data.xlsx')
            write_to_excel(output_filename, extracted_data)  # Pass first page's data
        
        result_label.config(text="Extraction and export successful!", fg="green")
    except Exception as e:
        result_label.config(text=f"An error occurred: {e}", fg="red")

# GUI
root = tk.Tk()
root.title("PDF Data Extractor")

pdf_path_label = tk.Label(root, text="PDF Path:")
pdf_path_label.grid(row=0, column=0)

pdf_path_entry = tk.Entry(root, width=50)
pdf_path_entry.grid(row=0, column=1)

pdf_path_button = tk.Button(root, text="Select PDF", command=select_pdf_file)
pdf_path_button.grid(row=0, column=2)

output_folder_label = tk.Label(root, text="Output Folder:")
output_folder_label.grid(row=1, column=0)

output_folder_entry = tk.Entry(root, width=50)
output_folder_entry.grid(row=1, column=1)

output_folder_button = tk.Button(root, text="Select Folder", command=select_output_folder)
output_folder_button.grid(row=1, column=2)

format_label = tk.Label(root, text="Output Format:")
format_label.grid(row=2, column=0)

format_var = tk.StringVar(root)
format_var.set("CSV")  # Default selection
format_optionmenu = tk.OptionMenu(root, format_var, "CSV", "JSON", "Excel")
format_optionmenu.grid(row=2, column=1)

extract_button = tk.Button(root, text="Extract and Export Data", command=extract_and_export_data)
extract_button.grid(row=3, column=1)

result_label = tk.Label(root, text="", fg="black")
result_label.grid(row=4, column=1)

root.mainloop()
