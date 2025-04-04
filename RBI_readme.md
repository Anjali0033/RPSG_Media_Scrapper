**Website Navigation**
Home > Publications > Annual > Annual Report on Banking Ombudsman Scheme

- Home: https://www.rbi.org.in/Home.aspx  
- Publications: https://www.rbi.org.in/Scripts/publications.aspx  
- Annual: https://www.rbi.org.in/Scripts/publications.aspx  
- Annual Report on Banking Ombudsman Scheme: https://www.rbi.org.in/Scripts/AnnualPublications.aspx?head=Annual%20Report%20on%20Banking%20Ombudsman%20Scheme  


Requirements

- Python 3.9.12  
- Modules: 'requests', 'lxml', 'pandas', 'openpyxl'  

Install dependencies:  
pip install requests, lxml, pandas, openpyxl

Run
1. Set the product URL in the script:  
   'url = 'https://www.rbi.org.in/Scripts/PublicationsView.aspx?id=22432''

2. Run the script:  
   'python scrape_product.py'

Output  
Returns an Excel file: 'RBI_BANK_COMPLAINT.xlsx'  
Includes:
- Title merged and centered in the first row  
- Proper column headers  
- Data sorted by “Total Maintainable Complaints disposed during the year 2022-23”  
- Auto-adjusted column widths for readability

Output Columns:
- Name of the Bank  
- Total Maintainable Complaints disposed during the year 2022-23  
- Of (2), Complaints resolved through conciliation/ mediation/ issuance of advisories  
 and so on......
