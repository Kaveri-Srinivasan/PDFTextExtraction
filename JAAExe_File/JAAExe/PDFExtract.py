import PyPDF2

def extract_text_from_pdf_with_pypdf2(pdf_file_path):
    text = ''
    with open(pdf_file_path, 'rb') as file:
        pdf_reader = PyPDF2.PdfReader(file)
        for page_num in range(len(pdf_reader.pages)):
            page = pdf_reader.pages[page_num]
            print(page_num)
            text += page.extract_text() + "\n"  # Extract text from each page
    return text

# Example usage
pdf_path = '/content/sample_data/PDFFiles/your_pdf_file.pdf'
extracted_text = extract_text_from_pdf_with_pypdf2(pdf_path)

print(extracted_text)
