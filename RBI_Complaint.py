import requests
from lxml import html
import pandas as pd
from openpyxl import load_workbook

def scrape_disposal_table(url):
    response = requests.get(url)
    tree = html.fromstring(response.content)

    title_xpath = "//td[@class='head' and contains(text(),'Mode of disposal of Maintainable Complaints against Scheduled Commercial Banks')]"
    title_element = tree.xpath(title_xpath)
    title = title_element[0].text_content().strip() if title_element else "Title not found"

    base_xpath = "//td[@class='head' and contains(text(),'Mode of disposal of Maintainable Complaints against Scheduled Commercial Banks')]/ancestor::table[1]"

    rows = tree.xpath(f"{base_xpath}//tr")

    data = []
    for row in rows:
        cells = row.xpath('./td')
        row_data = [cell.text_content().strip() for cell in cells]
        if row_data:
            data.append(row_data)

    columns = data[1]
    data = data[2:]  
    data = [row for row in data if len(row) == len(columns)]

    return title, columns, data

def save_to_excel(title, columns, data, filename='RBI_BANK_COMPLAINT.xlsx'):
    import openpyxl
    from openpyxl.utils import get_column_letter
    from openpyxl.styles import Font, Alignment

    aligned_data = []
    for row in data:
        if len(row) > len(columns):
            aligned_data.append(row[:len(columns)])
        else:
            aligned_data.append(row + [''] * (len(columns) - len(row)))

    df = pd.DataFrame(aligned_data, columns=columns)

    second_col = columns[1]
    df[second_col] = df[second_col].astype(str).str.replace(',', '', regex=False)
    df[second_col] = pd.to_numeric(df[second_col], errors='coerce')
    df = df.dropna(subset=[second_col])
    df_sorted = df.sort_values(by=second_col, ascending=False)

    title_row = pd.DataFrame([[title] + [''] * (len(columns) - 1)], columns=columns)
    header_row = pd.DataFrame([columns], columns=columns)

    final_df = pd.concat([title_row, header_row, df_sorted], ignore_index=True)
    final_df.to_excel(filename, index=False, header=False)

    wb = load_workbook(filename)
    ws = wb.active
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(columns))
    cell = ws.cell(row=1, column=1)
    cell.font = Font(bold=True, size=12)
    cell.alignment = Alignment(horizontal="center", vertical="center")

    for col_idx in range(1, len(columns) + 1):
        max_length = max(len(str(ws.cell(row=row, column=col_idx).value or "")) for row in range(1, ws.max_row + 1))
        ws.column_dimensions[get_column_letter(col_idx)].width = max_length + 2

    wb.save(filename)


url = 'https://www.rbi.org.in/Scripts/PublicationsView.aspx?id=22432'
title, columns, data = scrape_disposal_table(url)
save_to_excel(title, columns, data)
