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
    import pdfplumber
    import pandas as pd

    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[0]
        tables = page.extract_tables()

        if not tables:
            print("No tables found.")
            return

        table = tables[0]

        headers = [
    "Name of Company",
    "Complaints O/s at the beginning of the year",
    "Complaints Received during the period",
    "Complaints Total",
    "Disposed by way of - Recommendations",
    "Disposed by way of - Awards fvg complainant",
    "Disposed by way of - Awards fvg ins. Co.",
    "Disposed by way of - Withdrawal",
    "Disposed by way of - Non-Entertainable",
    "Disposed by way of - Total Disposed",
    "Disposal Duration - Within 3 months",
    "Disposal Duration - 3 months to 1 year",
    "Disposal Duration - Above 1 year",
    "Disposal Duration - Total Disposed",
    "Outstanding Duration - Within 3 months",
    "Outstanding Duration - 3 months to 1 year",
    "Outstanding Duration - Above 1 year",
    "Outstanding Duration - Total Outstanding"
]


        df = pd.DataFrame(table[2:], columns=headers)
        df = pd.DataFrame(table[2:], columns=headers)

        # Static premium data (₹ Cr) [Static values from https://www.screener.in/]
        premium_data = {
            "Aditya Birla Sun Life Insurance Co. Ltd.": 16000,
            "Aegon Life Ins.Co.Ltd.": 900,
            "Ageas Federal Life Ins.Co.Ltd.": 1200,
            "Aviva Life Ins. Co. India Pvt. Ltd.": 1100,
            "Bajaj Allianz Life Insurance Co. Ltd.": 21500,
            "Bharti AXA Life Ins. Co. Ltd.": 3200,
            "Canara HSBC Oriental Bank of Commerce Life Ins. Co. Ltd.": 3300,
            "Edelweiss Tokio Life Ins. Co. Ltd.": 1400,
            "Exide Life Insurance Company Ltd.": 2800,
            "Future Generali India Life Ins. Co. Ltd.": 2100,
            "HDFC Life Insurance Co. Ltd.": 60500,
            "ICICI Prudential Life Insurance Co. Ltd.": 52800,
            "IndiaFirst Life Insurance Co. Ltd.,": 5600,
            "Kotak Mahindra Life Insurance Company": 9400,
            "LIC of India": 240000,
            "Max Life insurance Co. Ltd.": 24000,
            "PNB Metlife India Ins. Co. P. Ltd.": 8900,
            "Pramerica Life Ins.Co.Ltd.": 500,
            "Reliance Nippon Life Insurance Co. Ltd.": 5800,
            "Sahara India Life Ins. Co. Ltd": 90,
            "SBI Life Insurance Co. Ltd.": 62000,
            "Shriram Life Ins. Co. Ltd.": 3500,
            "Star Union Dai-ichi-Life Ins. Co.": 4200,
            "Tata AIA Life Insurance Co. Ltd.": 14000
        }

        # Add sales and complaints per ₹100 Cr premium
        df["Total Complaints"] = pd.to_numeric(df["Complaints Total"], errors='coerce')
        df["Total Premium (₹ Cr)"] = df["Name of Company"].map(premium_data)
        df["Complaints per ₹100 Cr Premium"] = (df["Total Complaints"] / df["Total Premium (₹ Cr)"] * 100).round(2)

        # Save final output
        df.to_excel(excel_path, index=False)
        print(f"Cleaned data saved to: {excel_path}")


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
