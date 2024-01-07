import pdfplumber
import xlwt
import re
import os


def extract_price(text):
    match = re.search(r'小写.*?(\d+(\.\d+)?)', text)
    if match:
        return match.group(1)


def clean_text(text):
    return text.replace(' ', '').replace('　', '').replace('）', '').replace(')', '').replace('：', ':').replace('\n', '')


def get_pdf_files(directory):
    pdf_files = []
    for root, _, filenames in os.walk(directory):
        for filename in filenames:
            if filename.endswith('.pdf'):
                filepath = os.path.join(root, filename)
                pdf_files.append(filepath)
    return pdf_files


def read_and_write_to_excel(directory, output_file):
    data_to_write = []

    for pdf_file in get_pdf_files(directory):
        with pdfplumber.open(pdf_file) as pdf:
            first_page = pdf.pages[0]
            pdf_text = first_page.extract_text()

            if '发票' in pdf_text:
                price = extract_price(pdf_text)
                if price:
                    cleaned_price = clean_text(price)
                    data_to_write.append([cleaned_price])

    # Write data to Excel file
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet('Sheet1')

    for i, row in enumerate(data_to_write):
        for j, value in enumerate(row):
            sheet.write(i, j, value)

    workbook.save(output_file)
    print(f'Data has been written to {output_file}')


# Example Usage:
input_directory = r'C:\Users\long2244\Desktop\2023'
output_excel_file = 'output_prices.xls'
read_and_write_to_excel(input_directory, output_excel_file)