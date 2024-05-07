import tkinter as tk
from tkinter import filedialog
import os
import pandas as pd
import PyPDF2
import threading


def pdf_to_excel(pdf_file):
    pdf_reader = PyPDF2.PdfReader(pdf_file)
    num_pages = len(pdf_reader.pages)
    text_data = ''
    for page in range(num_pages):
        page_obj = pdf_reader.pages[page]
        text_data += page_obj.extract_text()
    text_data = text_data.split('\n')
    df = pd.DataFrame(text_data)
    excel_file = os.path.splitext(pdf_file)[0] + '.xlsx'
    df.to_excel(excel_file, index=False)
    return excel_file


def select_files():
    file_paths = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
    if file_paths:
        loading_label.config(text="Converting files...")
        thread = threading.Thread(target=convert_files, args=(file_paths,))
        thread.start()


def convert_files(file_paths):
    for file_path in file_paths:
        excel_file = pdf_to_excel(file_path)
        print(f"{file_path} converted to {excel_file}")
    loading_label.config(text="Conversion completed.")


def main():
    global loading_label
    root = tk.Tk()
    root.title("PDF to Excel Converter")

    frame = tk.Frame(root)
    frame.pack(padx=10, pady=10)

    select_button = tk.Button(frame, text="Select PDF Files", command=select_files)
    select_button.pack()

    loading_label = tk.Label(frame, text="")
    loading_label.pack()

    root.mainloop()


if __name__ == "__main__":
    main()
