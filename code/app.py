import tkinter as tk
import xlwings as xw
from tkinter import filedialog, messagebox

from imagepdf_to_text import extract_text_from_pdf
from text_to_excel import text_to_excel

class PDFConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF to Excel Converter")
        
        # PDF file picker
        self.pdf_label = tk.Label(root, text="PDF to Convert:")
        self.pdf_label.grid(row=0, column=0, padx=10, pady=5, sticky="w")
        self.pdf_entry = tk.Entry(root, width=50)
        self.pdf_entry.grid(row=0, column=1, padx=10, pady=5)
        self.pdf_button = tk.Button(root, text="Browse", command=self.select_pdf)
        self.pdf_button.grid(row=0, column=2, padx=10, pady=5)

        # Excel file picker
        self.excel_label = tk.Label(root, text="Output Excel Document:")
        self.excel_label.grid(row=1, column=0, padx=10, pady=5, sticky="w")
        self.excel_entry = tk.Entry(root, width=50)
        self.excel_entry.grid(row=1, column=1, padx=10, pady=5)
        self.excel_button = tk.Button(root, text="Browse", command=self.select_excel)
        self.excel_button.grid(row=1, column=2, padx=10, pady=5)

        # Convert Button
        self.convert_button = tk.Button(root, text="Convert", command=self.convert_files)
        self.convert_button.grid(row=2, column=1, pady=20)
        
        
    def select_pdf(self):
        file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if file_path:
            self.pdf_entry.delete(0, tk.END)
            self.pdf_entry.insert(0, file_path)


    def select_excel(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsm",
                                                 filetypes=[("Excel Files", "*.xlsm")])
        if file_path:
            self.excel_entry.delete(0, tk.END)
            self.excel_entry.insert(0, file_path)


    def convert_files(self):
        pdf_path = self.pdf_entry.get()
        excel_path = self.excel_entry.get()

        if not pdf_path or not excel_path:
            messagebox.showerror("Error", "Please select both a PDF file and an output Excel file.")
            return
        
        try:
            self.process_conversion(pdf_path, excel_path)
            messagebox.showinfo("Success", "Conversion completed successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Conversion failed:\n{e}")


    def append_text_to_workbook(self, text, excel_path):    
        with xw.App(visible=True) as app:
            workbook = app.books.open(excel_path)
            text_to_excel(workbook, text)
            workbook.save()


    def process_conversion(self, pdf_path, excel_path):
        text = extract_text_from_pdf(pdf_path)    
        self.append_text_to_workbook(text, excel_path)
        

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFConverterApp(root)
    root.mainloop()