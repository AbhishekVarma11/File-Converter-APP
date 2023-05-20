import tkinter as tk
from tkinter import filedialog
import docx2pdf
import PyPDF2
from docx import Document
import os
import shutil

def word_to_pdf():
    filepath = filedialog.askopenfilename(title="Select Word file", filetypes=(("Word files", "*.docx"),))
    if filepath:
        try:
            output_filename = filepath.replace(".docx", ".pdf")
            docx2pdf.convert(filepath, output_filename)
            status_label.config(text="Conversion successful! Click below to download the PDF.")
            download_button.config(state=tk.NORMAL, command=lambda: download_file(output_filename))
        except Exception as e:
            status_label.config(text="Error converting the file. Please select a valid Word file.")
    else:
        status_label.config(text="Please select a Word file.")

def pdf_to_word():
    filepath = filedialog.askopenfilename(title="Select PDF file", filetypes=(("PDF files", "*.pdf"),))
    if filepath:
        try:
            pdf = open(filepath, 'rb')
            reader = PyPDF2.PdfReader(pdf)
            output_filename = filepath.replace(".pdf", ".docx")
            document = Document()
            for page in reader.pages:
                text = page.extract_text()
                document.add_paragraph(text)
            document.save(output_filename)
            pdf.close()
            status_label.config(text="Conversion successful! Click below to download the Word document.")
            download_button.config(state=tk.NORMAL, command=lambda: download_file(output_filename))
        except Exception as e:
            status_label.config(text="Error converting the file. Please select a valid PDF file.")
    else:
        status_label.config(text="Please select a PDF file.")

def download_file(filepath):
    try:
        download_dir = os.path.expanduser("~") + "/Downloads/"
        shutil.copy(filepath, download_dir)
        status_label.config(text="Download successful!", fg="green")
    except Exception as e:
        status_label.config(text="Download failed.", fg="red")

root = tk.Tk()
root.title("File Converter")

# Configure app window size and position
window_width = 400
window_height = 250
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
x = (screen_width // 2) - (window_width // 2)
y = (screen_height // 2) - (window_height // 2)
root.geometry(f"{window_width}x{window_height}+{x}+{y}")

# Configure app window background color
root.configure(bg="#f0f0f0")

# Word to PDF button
word_to_pdf_button = tk.Button(root, text="Word to PDF", command=word_to_pdf, font=("Arial", 14), bg="#4CAF50", fg="white", padx=10, pady=5)
word_to_pdf_button.pack(pady=10)

# PDF to Word button
pdf_to_word_button = tk.Button(root, text="PDF to Word", command=pdf_to_word, font=("Arial", 14), bg="#2196F3", fg="white", padx=10, pady=5)
pdf_to_word_button.pack(pady=10)

# Status label
status_label = tk.Label(root, text="", font=("Arial", 12), fg="black", bg="#f0f0f0")
status_label.pack(pady=10)

# Download button
download_button = tk.Button(root, text="Download", state=tk.DISABLED, font=("Arial", 14), bg="#FF5722", fg="white", padx=10, pady=5)
download_button.pack(pady=10)

root.mainloop()
