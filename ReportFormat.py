import dataclasses
import os.path
import time

import openpyxl

from dataclasses import dataclass
from openpyxl.styles import Font, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.dimensions import ColumnDimension
from openpyxl.worksheet.table import Table, TableStyleInfo

import Constants
import Constants as cons

wb = openpyxl.Workbook()
sh = wb.active


# Verifier Information
@dataclass()
class VerifierInfo:
    name: str
    date: str
    location: str
    model: str


# Product Information Data From User Typing
product_info: [{'name': str, 'fw': str}] = []


# Check Whether The Report File Is Or Not
def whether_check_report():
    if os.path.isfile(cons.report_path):
        print('Report file is exist')
        wb = openpyxl.load_workbook(cons.report_path)
        sh = wb.active
    else:
        print('Report file is not exist')
        wb = openpyxl.Workbook()
        sh = wb.active


# Get The Verifier Information From User
def get_user_info():
    # Type The Verifier information
    print('Please type the verifier info\n')
    veri_name = input('verifier name?\n')
    veri_date = cons.today_time_format
    veri_location = input('Verifier Location?\n')
    veri_model = input('Verify Model?')
    user_info = VerifierInfo(name=veri_name, date=veri_date,
                             location=veri_location, model=veri_model)

    return user_info


# Get The Product Information
def get_product_info():
    # Type The Verifier information
    print('Please type the product info\n')
    modules = ['Model', 'IR', 'EO', 'OP1', 'OP2', 'OP3']
    for module in modules:
        name = input(f' If there are no additional fields to fill in, type None\n {module} Name?')
        fw = input(f'{module} FW?')
        if name.lower() == 'none':
            return
        elif fw.lower() == 'none':
            return
        else:
            item = {'name': name, 'fw': fw}
            product_info.append(item)


# Make A Table With Column Title
def make_table_with_title(titles: [str], title_pos: {'column': int, 'row': int},
                          align: {'ver': str, 'hor': str}, data_area: {'start': str, 'end': str}):
    whether_check_report()
    for i, title in enumerate(titles):
        title_col = sh.cell(column=title_pos['column'] + i, row=title_pos['row'], value=title)
        title_col.alignment = Alignment(vertical=align['ver'], horizontal=align['hor'])
    add_table = Table(displayName=f'{titles[0]}', ref=f'{data_area["start"]}:{data_area["end"]}')
    style = TableStyleInfo(name='TableStyleMedium9', showFirstColumn=False,
                           showLastColumn=False, showRowStripes=True, showColumnStripes=True)
    add_table.tableStyleInfo = style
    sh.add_table(add_table)


# Fill In Information Data in Excel
def fill_user_product_info(user):
    user_info_fields = [
        ['Verifier', f'{user.name}'],
        ['Date', f'{user.date}'],
        ['Location', f'{user.location}'],
        ['Model', f'{user.model}']
    ]
    # Create A User Information Table
    make_table_with_title(cons.user_info_title, cons.user_info_title_pos,
                          {'ver': 'center', 'hor': 'center'}, cons.user_data_area)
    # Create A Product Information Table
    make_table_with_title(cons.product_info_title, cons.product_info_title_pos,
                          {'ver': 'center', 'hor': 'center'}, cons.product_data_area)
    # Fill In The User Field
    for row, info in enumerate(user_info_fields):
        sh.cell(column=2, row=row + 7, value=info[0])
        sh.cell(column=3, row=row + 7, value=info[1])
    # Fill In The Product Field
    for i, product in enumerate(product_info, start=cons.product_info_title_pos['row']):
        sh.cell(column=cons.product_info_title_pos['column'],
                row=i + 1, value=f'{product["name"]}')
        sh.cell(column=cons.product_info_title_pos['column'] + 1,
                row=i + 1, value=f'{product["fw"]}')

    # Alignment The Table
    for column_cells in sh.columns:
        length = max(len(str(cell.value)) * 1.3 for cell in column_cells)
        sh.column_dimensions[column_cells[0].column_letter].width = length

    wb.save(cons.report_path)


# Set The Test File Format
# Title
def set_report_format(date, model):
    sh.title = f'{model}'
    font_clear_go = Font(
        name='맑은 고딕',
        bold=True,
        size=24
    )
    border_title = Border(
        left=Side(border_style='thin'),
        right=Side(border_style='thin'),
        top=Side(border_style='thin'),
        bottom=Side(border_style='thin')
    )

    title = sh.merge_cells(start_column=2, start_row=2, end_column=7, end_row=4)
    title_name = sh.cell(row=2, column=2, value='Test Report')
    title_name.alignment = Alignment(horizontal='center', vertical='center')
    title_name.font = font_clear_go
    border_around_merged_cell(sh, 2, 2,
                              7, 4, 'thin')

    # Create Excel File By Time
    wb.save(cons.report_path)


# Set Only A Specified Border Of Merged Cell
def border_merged_cell_with_part(sheet, start_column, start_row, end_column, end_row, style, part):
    if part == 'top':
        for i in range(start_column, end_column + 1):
            # Set The Top Border
            cell_top = sheet.cell(column=start_column, row=start_row)
            cell_top.border = Border(top=Side(border_style=f'{style}'))
            start_column += 1
    elif part == 'bottom':
        for i in range(start_column, end_column + 1):
            # Set The Bottom Border
            cell_bottom = sheet.cell(column=start_column, row=end_row)
            cell_bottom.border = Border(bottom=Side(border_style=f'{style}'))
            start_column += 1
    elif part == 'left':
        for i in range(start_row, end_row + 1):
            # Set The Left Border
            cell_left = sheet.cell(column=start_column, row=start_row)
            cell_left.border = Border(left=Side(border_style=f'{style}'))
            start_row += 1
    elif part == 'right':
        # Set The Right Border
        cell_right = sheet.cell(column=end_column, row=start_row)
        cell_right.border = Border(right=Side(border_style=f'{style}'))
        start_row += 1


# Set A Whole Border Of Merged Cell
def border_around_merged_cell(sheet, start_column, start_row, end_column, end_row, style):
    loop_count_fir = 0
    # Set The Top and Bottom Border
    for i in range(start_column, end_column + 1):
        loop_count_fir = i
        # Set The Top Border
        cell_top = sheet.cell(column=start_column, row=start_row)
        cell_top.border = Border(top=Side(border_style=f'{style}'))
        # Set The Bottom Border
        cell_bottom = sheet.cell(column=start_column, row=end_row)
        cell_bottom.border = Border(bottom=Side(border_style=f'{style}'))
        start_column += 1

    start_column = start_column - loop_count_fir + 1
    loop_count_sec = 0
    # Set The Left and Right Border
    for i in range(start_row, end_row + 1):
        loop_count_sec = i
        # Set The Left Border
        cell_left = sheet.cell(column=start_column, row=start_row)
        cell_left.border = Border(left=Side(border_style=f'{style}'))
        # Set The Right Border
        cell_right = sheet.cell(column=end_column, row=start_row)
        cell_right.border = Border(right=Side(border_style=f'{style}'))
        start_row += 1

    # Retry border Set The LeftTop and RightBottom
    # Because When Setting The Merge, Only The Final Setting Seem to be Applied
    start_row = start_row - loop_count_sec + 1
    sheet.cell(column=start_column, row=start_row).border = Border(
        left=Side(border_style=f'{style}'),
        top=Side(border_style=f'{style}'),
    )
    sheet.cell(column=end_column, row=end_row).border = Border(
        right=Side(border_style=f'{style}'),
        bottom=Side(border_style=f'{style}'),
    )
    sheet.cell(column=end_column, row=start_row).border = Border(
        right=Side(border_style=f'{style}'),
        top=Side(border_style=f'{style}'),
    )
    sheet.cell(column=start_column, row=end_row).border = Border(
        left=Side(border_style=f'{style}'),
        bottom=Side(border_style=f'{style}'),
    )
