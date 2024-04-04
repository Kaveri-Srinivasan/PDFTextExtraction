import os
from tkinter import Tk, messagebox, Button, Frame, Label, Text, Scrollbar
from tkinter.filedialog import askdirectory
from PyPDF2 import PdfReader
from openpyxl import Workbook

class ExcelDialog:
    def __init__(self, parent, excel_path):
        self.parent = parent
        self.excel_path = excel_path

        self.window = Tk()
        self.window.title("Extracted Excel File")
        self.window.configure(bg="lightgray")
        
        self.frame = Frame(self.window, bg="lightgray")
        self.frame.pack(padx=20, pady=20)

        self.label = Label(self.frame, text="Extracted Excel file Location:", bg="lightgray", font=("Arial", 12, "bold"), fg="black",anchor="w")
        self.label.pack(fill='x')

        self.text = Text(self.frame, height=2, width=50, font=("Arial", 10), bg="white", fg="black")
        self.text.insert("1.0", self.excel_path)
        self.text.pack(pady=15)

        self.scrollbar = Scrollbar(self.frame, command=self.text.yview)
        self.text.config(yscrollcommand=self.scrollbar.set)
        self.scrollbar.pack(side="right", fill="y")

        self.button = Button(self.window, text="Open Excel File", command=self.open_excel, font=("Arial", 12, "bold"), bg="Blue", fg="white", highlightbackground="green", padx=10, pady=5)
        self.button.pack(pady=5)

    def open_excel(self):
        try:
            os.startfile(self.excel_path)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open Excel file: {e}")

    def show(self):
        self.window.mainloop()

def select_folder():
    Tk().withdraw()  # We don't want a full GUI, so keep the root window from appearing
    folder = askdirectory()  # Show a dialog to select a directory
    return folder

def extract_text_from_pdf(pdf_path):
    text = ""
    try:
        reader = PdfReader(pdf_path)
        for page in reader.pages:
            text += page.extract_text() + "\n"
    except Exception as e:
        print(f"Error reading {pdf_path}: {e}")
    return text

def save_to_excel(pdf_texts, folder_path):
    excel_path = os.path.join(folder_path, "ExtractedTexts.xlsx")
    workbook = Workbook()
    sheet = workbook.active

    for i, text in enumerate(pdf_texts, start=1):
        sheet[f"A{i}"] = text

    workbook.save(excel_path)

    dialog = ExcelDialog(None, excel_path)
    dialog.show()

def main():
    folder_path = select_folder()
    if folder_path:
        pdf_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.pdf')]
        pdf_texts = []

        for pdf_file in pdf_files:
            full_path = os.path.join(folder_path, pdf_file)
            print(f"Extracting text from: {pdf_file}")
            pdf_texts.append(extract_text_from_pdf(full_path))

        if pdf_texts:
            save_to_excel(pdf_texts, folder_path)
        else:
            print("No PDF files found in the folder.")
    else:
        print("Folder selection was cancelled.")

if __name__ == "__main__":
    main()
