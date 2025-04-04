import requests
from lxml import html
import PyPDF2
import pdfplumber
import pandas as pd
import os  

def download_pdf():
    url = 'https://cioins.co.in/AnnualReports'
    response = requests.get(url)
    tree = html.fromstring(response.content)

    # Extract the link for Annual Report 2023-24
    relative_link = ''.join(tree.xpath('//a[contains(.,"Annual Report for 2023-24")]/@href'))
    
    if not relative_link:
        print("Annual Report link not found.")
        return None

    # Build full URL
    if not relative_link.startswith("http"):
        base_url = "https://cioins.co.in"
        pdf_url = base_url + relative_link
    else:
        pdf_url = relative_link

    print(f"Found report link: {pdf_url}")

    # Download and save PDF
    pdf_response = requests.get(pdf_url)
    pdf_path = "Insurance_Original.pdf"
    with open(pdf_path, "wb") as f:
        f.write(pdf_response.content)

    print(f"PDF downloaded and saved as: {pdf_path}")
    return pdf_path

def extract_rotate_and_save_pdf(input_pdf, output_pdf, page_number):
    with open(input_pdf, 'rb') as infile:
        reader = PyPDF2.PdfReader(infile)
        writer = PyPDF2.PdfWriter()

        page = reader.pages[page_number - 1]
        page.rotate(90)
        writer.add_page(page)

        with open(output_pdf, 'wb') as outfile:
            writer.write(outfile)
    print(f"Rotated page saved as: {output_pdf}")

def extract_tables_to_excel(pdf_path, excel_path):
    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[0]
        tables = page.extract_tables()

        if not tables:
            print("No tables found on the page.")
            return

        with pd.ExcelWriter(excel_path) as writer:
            for i, table in enumerate(tables):
                df = pd.DataFrame(table[1:], columns=table[0])
                df.to_excel(writer, sheet_name=f"Table_{i+1}", index=False)

        print(f"Tables saved to Excel: {excel_path}")

downloaded_pdf = download_pdf()

if downloaded_pdf:
    rotated_pdf = "INSURANCE OMBUDSMENT.pdf"
    excel_output = "INSURANCE OMBUDSMENT.xlsx"
    page_to_rotate = 52

    extract_rotate_and_save_pdf(downloaded_pdf, rotated_pdf, page_to_rotate)
    extract_tables_to_excel(rotated_pdf, excel_output)

    # Delete the original full PDF
    os.remove(downloaded_pdf)
    print(f"Removed original PDF: {downloaded_pdf}")
