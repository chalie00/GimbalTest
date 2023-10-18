import Constants as cons
import Major_Function as mf
import CHARACTER_TRANSFORMATION as ct
import ReportFormat as rf

import os
import time
import xlsxwriter
import pandas as pd
import pytesseract
import pyautogui

from pywinauto import Application, keyboard, mouse
from pywinauto_recorder.player import *
from PIL import Image, ImageGrab

# Get The Verifier Information
user_info = rf.get_user_info()

# Get The Product Information
rf.get_product_info()

rf.fill_user_product_info(user_info)

# Today Implement
rf.set_report_format(user_info.date, user_info.model)

# Start with Application Object
app = Application(backend='uia').start(r'C:\GimbalTBX\GimbalTBX.exe')
mainDlg = app['Gimbal TBX Series']

# Login -> 닫기 -> 취소
# mf.login_correct_WithOnlyPW(mainDlg, 'qwer')
# time.sleep(1)
# mainDlg.child_window(title='닫기', control_type='Button').click()
# time.sleep(1)

mainDlg.child_window(title="No", control_type='Button').click()
time.sleep(1)

# Left Menu Hidden/Show
mouse.click(coords=(cons.left_menu_hide[0], cons.left_menu_hide[1]))
mouse.click(coords=(cons.left_menu_show[0], cons.left_menu_show[1]))

# Left Menu Select
time.sleep(1)
mainDlg.CheckBox2.click_input()

time.sleep(1)
mainDlg.CheckBox3.click_input()

time.sleep(1)
mainDlg.CheckBox4.click_input()
mainDlg.capture_as_image().save(r'Capture/Gimbal.png')

time.sleep(1)
mainDlg.CheckBox1.click_input()

time.sleep(1)
# Screen Shot
mainDlg.capture_as_image().save(r'Capture/Power.png')
time.sleep(1)

# Menu Screen Capture
menus = ['EO', 'IR', 'GIMBAL', 'POWER']
mf.upload_excel_with_menus(mainDlg, menus)

# OCR
# Yaw, Pitch Angle OCR
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
img = Image.open(r'Capture\EO.png')

yaw_coord = mf.cal_element_coordinate(cons.yaw_lt_rb[0],
                                      cons.yaw_lt_rb[1],
                                      cons.yaw_lt_rb[2],
                                      cons.yaw_lt_rb[3])
pitch_coord = mf.cal_element_coordinate(cons.pitch_lt_rb[0],
                                        cons.pitch_lt_rb[1],
                                        cons.pitch_lt_rb[2],
                                        cons.pitch_lt_rb[3])

(img.crop([yaw_coord['l'],
           yaw_coord['t'],
           yaw_coord['r'],
           yaw_coord['b']])
 .save(r'Capture\Yaw.png'))

(img.crop([pitch_coord['l'],
           pitch_coord['t'],
           pitch_coord['r'],
           pitch_coord['b']])
 .save(r'Capture\Pitch.png'))

img_yaw: str = 'Capture\Yaw.png'
img_pitch: str = 'Capture\Pitch.png'

txt_msk_yaw = ct.image_to_string_with_hsk(img_yaw)
txt_msk_pitch = ct.image_to_string_with_hsk(img_pitch)

result_yaw = ct.generate_txt_array_with_img(txt_msk_yaw, 'Yaw')
result_pitch = ct.generate_txt_array_with_img(txt_msk_pitch, 'Pitch')

# CLick Aspect Ratio
mouse.click(coords=(cons.aspect_ratio[0], cons.aspect_ratio[1]))
time.sleep(1)
mf.upload_excel_with_menus(mainDlg, ['ASPECT'])

# Function Test from Power to Gimbal
mainDlg.child_window(title="POWER", control_type='CheckBox').click_input()
# EO Power Off -> On
mouse.click(coords=(cons.eo_power_off[0], cons.eo_power_off[1]))
time.sleep(1)
mouse.click(coords=(cons.gimbal_power_ui[0], cons.gimbal_power_ui[1]))

area_capture_dic = [{'name': 'Gimbal Power Off', "coord": cons.gimbal_power_ui},
                    {'name': 'EO Power Off', "coord": cons.eo_power_ui},
                    {'name': 'IR Power Off', "coord": cons.ir_power_ui},
                    {'name': 'FAN Power Off', "coord": cons.fan_power_ui},
                    {'name': 'HEATER Power Off', "coord": cons.heater_power_ui},
                    {'name': 'LRF Power Off', "coord": cons.lrf_power_ui}
                    ]
mf.upload_excel_after_check_Report(mainDlg, area_capture_dic)

# TODO (Com) - Apply the information in Excel with multiple table
# TODO Fill In The Test Step and Test Code, Result In Excel
# TODO (Com) - Modify a Save Path
