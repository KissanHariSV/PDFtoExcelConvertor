import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Font

def pdf_to_excel(pdf_path, excel_path):
    wb = Workbook()
    ws = wb.active
    ws.title = 'PDF Data'
    
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            for table in page.extract_tables():
                for row in table:
                    ws.append(row)
                    
    # Apply basic formatting (bold)
    for row in ws.iter_rows():
        for cell in row:
            cell.font = Font(bold=True)
    
    wb.save(excel_path)

if __name__ == "__main__":
    pdf_path = 'your_pdf_file.pdf'  # Replace 'your_pdf_file.pdf' with the path to your PDF file
    excel_path = 'output.xlsx'       # Output Excel file path

    pdf_to_excel(pdf_path, excel_path)

    print(f"PDF has been converted to Excel. Output file saved as '{excel_path}'")
