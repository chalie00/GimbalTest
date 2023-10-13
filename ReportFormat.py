import dataclasses
import time

import openpyxl

from dataclasses import dataclass
from openpyxl.styles import Font, Border, Side, Alignment
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


# TODO Set The Test File Format
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
    title_name = sh.cell(row=2, column=2, value="Test Report")
    title_name.alignment = Alignment(horizontal='center', vertical='center')
    title_name.font = font_clear_go
    border_around_merged_cell(sh, 2, 2,
                              7, 4, 'thin')

    # TODO Create Excel File By Time
    wb.save(rf'TestResult\Test Report_{date}.xlsx')


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
