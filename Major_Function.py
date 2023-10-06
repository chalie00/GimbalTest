import time
import numpy as np

import pytesseract
import xlsxwriter
import pyocr
import cv2

import Constants

from pywinauto import mouse
from PIL import Image


def login_correct_WithOnlyPW(dia, pw):
    # dia['passwordEdit'].type_keys(f'{pw}')
    # dia['OK'].click()
    dia.passwordEdit.type_keys(pw)
    time.sleep(1)
    dia['OK'].click()


def capture_Menu(dia, menus):
    workbook = xlsxwriter.Workbook(r'TestResult\Menu Capture.xlsx')
    worksheet = workbook.add_worksheet()
    row_position = 0

    for row_num, menu in enumerate(menus):
        row_num = row_position
        dia.child_window(title=f'{menu}', control_type='CheckBox').click_input()
        dia.capture_as_image().save(f'Capture/{menu}.png')
        worksheet.write(row_num, 0, menu)
        worksheet.insert_image(row_num + 1, 0, f'Capture/{menu}.png')

        time.sleep(1)
        row_position += 38

    workbook.close()


def cal_element_coordinate(x, y, w, h):
    ele_coord = {'x': 0,
                 'y': 0,
                 'w': 0,
                 'h': 0}
    coordinate_array = Constants.main_Dig_Coordinate
    ele_coord['x'] = x - coordinate_array[0]
    ele_coord['y'] = y - coordinate_array[1]
    ele_coord['w'] = w - coordinate_array[0]
    ele_coord['h'] = h - coordinate_array[1]

    return ele_coord

