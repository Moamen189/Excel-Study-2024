import pandas as pd
from datetime import datetime, timedelta
from openpyxl import Workbook
from openpyxl.styles import Alignment
from openpyxl.worksheet.table import Table, TableStyleInfo


start_date = datetime(2024, 1, 1)
end_date = datetime(2024, 12, 31)
date_range = pd.date_range(start_date, end_date)


data = {'Date': date_range, 'Work': '', 'TECH': ''}  
df = pd.DataFrame(data)


excel_file_path = '2024_Description_Schedule.xlsx'
with pd.ExcelWriter(excel_file_path, engine='openpyxl') as writer:
  
    df.to_excel(writer, sheet_name='Sheet1', index=False, startcol=8, startrow=0)


    workbook = writer.book
    worksheet = writer.sheets['Sheet1']


    header_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    for col_letter in ['I', 'J', 'K', 'L']:
        header_cell = worksheet[f'{col_letter}1']
        header_cell.alignment = header_alignment
        header_cell.font = header_cell.font.copy(bold=True)
        worksheet.column_dimensions[col_letter].width = 25
        worksheet.row_dimensions[1].height = 20

    date_column = worksheet['I']
    for cell in date_column:
        cell.number_format = 'yyyy-mm-dd'
        cell.alignment = Alignment(horizontal='center')
        worksheet.row_dimensions[cell.row].height = 25

    work_column = worksheet['J']
    for cell in work_column:
        cell.alignment = Alignment(wrap_text=True)
        worksheet.row_dimensions[cell.row].height = 25

    tech_column = worksheet['L']
    for cell in tech_column:
        cell.alignment = Alignment(wrap_text=True)
        worksheet.row_dimensions[cell.row].height = 25


    table_range = f'I1:L{len(df) + 1}' 
    table = Table(displayName="DataTable", ref=table_range)
    style = TableStyleInfo(
        name="TableStyleMedium9", showFirstColumn=False, showLastColumn=False,
        showRowStripes=True, showColumnStripes=True)
    table.tableStyleInfo = style
    worksheet.add_table(table)

print(f"Excel file '{excel_file_path}' created successfully.")
