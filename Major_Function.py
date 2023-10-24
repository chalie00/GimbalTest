import time
from typing import Any

import numpy as np
import os.path

import pyautogui
import pytesseract
import xlsxwriter
import pyocr
import cv2
import openpyxl

import Constants as cons

from pywinauto import mouse, WindowSpecification
from PIL import Image
from pynput.mouse import Listener, Button


def login_correct_WithOnlyPW(dia, pw):
    # dia['passwordEdit'].type_keys(f'{pw}')
    # dia['OK'].click()
    dia.passwordEdit.type_keys(pw)
    time.sleep(1)
    dia['OK'].click()


# Upload To The Excel File With Menus(can select object) Array
def upload_excel_with_menus(dia, menus):
    if os.path.isfile(cons.report_path):
        print('Report file is exist')
        wb = openpyxl.load_workbook(cons.report_path)
        sh = wb.active
        capture_after_check_menus_count(wb, sh, dia, menus)
    else:
        print('Report file is not exist')
        wb = openpyxl.Workbook()
        sh = wb.active
        capture_after_check_menus_count(wb, sh, dia, menus)


# Upload To The Excel File With Coordinate(can't select object) Array
# After Check The Excel File
def upload_excel_after_check_Report(dia: WindowSpecification, coordinates: [{str: str, str: [int]}]):
    if os.path.isfile(cons.report_path):
        wb = openpyxl.load_workbook(cons.report_path)
        sh = wb.active
        capture_upload_excel_with_coordinate(dia, wb, sh, coordinates)
    else:
        wb = openpyxl.Workbook()
        sh = wb.active
        capture_upload_excel_with_coordinate(dia, wb, sh, coordinates)


# Upload To The Excel File With Coordinate(can't select object) Array
def capture_upload_excel_with_coordinate(dia: WindowSpecification, wb: None, sh: None,
                                         coordinates: [{str: str, str: [int]}]):
    for coordinate in coordinates:
        area_capture_with_coordinate(dia, coordinate['name'], coordinate['coord'])
        sh.cell(row=cons.excel_No, column=1, value=f'{coordinate["name"]} Crop')
        img = openpyxl.drawing.image.Image(rf'Capture\{coordinate["name"]}_crop.png')
        sh.add_image(img, f'A{cons.excel_No + 1}')
        time.sleep(1)
        cons.excel_No += 3
    wb.save(cons.report_path)


# Check Menu Array Count And Capture
def capture_after_check_menus_count(wb, sh, dia, menus):
    if len(menus) > 1:
        for i, menu in enumerate(menus):
            dia.child_window(title=f'{menu}', control_type='CheckBox').click_input()
            dia.capture_as_image().save(f'Capture/{menu}.png')
            sh.cell(row=cons.excel_No, column=1, value=rf'{menus[i]}')
            img = openpyxl.drawing.image.Image(rf'Capture\{menus[i]}.png')
            sh.add_image(img, f'A{cons.excel_No + 1}')
            time.sleep(1)
            cons.excel_No += 35
        wb.save(cons.report_path)
    else:
        dia.capture_as_image().save(rf'Capture\{menus[0]}.png')
        time.sleep(1)
        sh.cell(row=cons.excel_No, column=1, value=f'{menus[0]}')
        img = openpyxl.drawing.image.Image(rf'Capture\{menus[0]}.png')
        sh.add_image(img, f'A{cons.excel_No + 1}')
        time.sleep(1)
        cons.excel_No += 35
        wb.save(cons.report_path)


#  Calculate coordinates from the zero position in the main frame
#  to the selected element
def cal_element_coordinate(left, top, right, bottom):
    ele_coord = {'l': 0,
                 't': 0,
                 'r': 0,
                 'b': 0}
    coordinate_array = cons.Dig_zero_Coordinate
    ele_coord['l'] = left - coordinate_array[0]
    ele_coord['t'] = top - coordinate_array[1]
    ele_coord['r'] = right - coordinate_array[0]
    ele_coord['b'] = bottom - coordinate_array[1]

    return ele_coord


# Capture selected Area or Element
def area_capture_with_coordinate(dia: WindowSpecification, title: str, coordinate: [int]):
    coord = cal_element_coordinate(coordinate[0], coordinate[1],
                                   coordinate[2], coordinate[3])
    img = dia.capture_as_image()
    (img.crop([coord['l'], coord['t'], coord['r'], coord['b']])
     .save(fr'Capture\{title}_crop.png'))


# Set the save directory of capture image
def set_save_image_dir(group, code):
    current_time = cons.current_time_bdYHMS
    save_path = rf'Capture\{group}\{code}_{current_time}.png'
    return save_path


# TODO (Com 10/24) - Virtual JoyStick Control By Mouse Drag
def mouse_drag_drop(start_coord, direction, hold_time, degree, end_coord):
    pyautogui.moveTo(start_coord[0], start_coord[1])
    pyautogui.mouseDown()
    cal_degree = [(abs(end_coord[0] - start_coord[0])) * degree,
                  (abs(start_coord[1] - end_coord[1])) * degree]
    print(cal_degree)
    if direction == 'up':
        pyautogui.moveTo(start_coord[0], start_coord[1] - cal_degree[1])
    elif direction == 'down':
        pyautogui.moveTo(start_coord[0], start_coord[1] + cal_degree[1])
    elif direction == 'left':
        pyautogui.moveTo(start_coord[0] - cal_degree[0], start_coord[1])
    elif direction == 'right':
        pyautogui.moveTo(start_coord[0] + cal_degree[0], start_coord[1])
    time.sleep(hold_time)
    pyautogui.mouseUp()

# Function called on a mouse click
# def on_click(x, y, button, pressed):
#     # Check if the left button was pressed
#     if pressed and button == Button.left:
#         # Print the click coordinates
#         print(f'x={x} and y={y}')
#
#
# # Initialize the Listener to monitor mouse clicks
# with Listener(on_click=on_click) as listener:
#     listener.join()
