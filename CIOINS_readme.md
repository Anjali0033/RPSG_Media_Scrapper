Insurance Ombudsman Complaints Extractor

This script:

- Downloads the **Annual Report 2023-24** from https://cioins.co.in/AnnualReports  
- Extracts TABLE OF Complaints Disposal statement for the period 01.04.2023 to 31.03.2024 
- Saves the page as a PDF and extracts tables into an Excel file


Requirements

- Python 3.9.12
- Modules: 'requests', 'lxml', 'pandas', 'openpyxl', 'pdfplumber', 'PyPDF2'

Install with:

bash
pip install requests ,lxml ,pandas ,openpyxl ,pdfplumber ,PyPDF2

How to Run
bash
python insurance_scraper.py


Output
- insurance_scraper.py (main script)
- 'INSURANCE OMBUDSMENT.pdf': Page 52, rotated - (Complaints Disposal statement for the period 01.04.2023 to 31.03.2024)
- 'INSURANCE OMBUDSMENT.xlsx': Tables from that page  
- Deletes original downloaded PDF after processing 
