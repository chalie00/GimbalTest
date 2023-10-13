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

# Type The Verifier information
print('Please type the verifier info\n')
verifier_name = input('verifier name?\n')
verification_date = cons.today_time_format
verification_location = input('Verifier Location?\n')
verification_model = input('Verify Model?')

# TODO Today Implement
verifier = rf.VerifierInfo(name=verifier_name, date=verification_date,
                           location=verification_location, model=verification_model)
rf.set_report_format(verification_date, verification_model)

# Start with Application Object
app = Application(backend='uia').start(r'C:\GimbalTBX\GimbalTBX.exe')
mainDlg = app['Gimbal TBX Series']
# sys_tray = app.window(class_name = 'Shell_TrayWnd')
# sys_tray = app.ShellTrayWnd.NotificationAreaToolbar

# sys_tray.child_window(title='Gimbal TBX Series').click()
# sys_tray.ClickSystemTrayIcon('Gimbal TBX SeriesDialog')
# print(taskbar.SystemTrayIcons.texts())
# time.sleep(3)

# app_Connect = Application().connect(path=r'C:\GimbalTBX\GimbalTBX.exe')
# systray_Icon = app.ShellTrayWnd.NotificationAreaToolbar
# taskbar.ClickSystemTrayIcon('GimbalTBX')

# Login -> 닫기 -> 취소
mf.login_correct_WithOnlyPW(mainDlg, 'qwer')
time.sleep(1)
mainDlg.child_window(title='닫기', control_type='Button').click()
time.sleep(1)

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

# # attach an image to excel with xlsxwriter
# workbook = xlsxwriter.Workbook('image.xlsx')
# worksheet = workbook.add_worksheet()
# worksheet.insert_image('B2', r'Capture\Gimbal.png')
# worksheet.insert_image('B38', r'Capture\Power.png')
# workbook.close()
#
# # Control Excel with Pandas
# df = pd.DataFrame({'Data': [10, 20, 30]})
# writer = pd.ExcelWriter('Result.xlsx', engine='xlsxwriter')
# df.to_excel(writer, sheet_name='Capture')
# workbook = writer.book
# worksheet = writer.sheets['Capture']
# worksheet.insert_image('B2', r'Capture\Gimbal.png')
# writer.close()

# Menu Screen Capture
menus = ['EO', 'IR', 'GIMBAL', 'POWER']
mf.upload_excel_with_menus(mainDlg, menus)

# # Next implementation Get Static Text
# static_ele = mainDlg.child_window(title='Static5', class_name="Static")
# if static_ele.exists(timeout=10):
#     print(static_ele)
# label = mainDlg.child_window(class_name="Static", found_index=1).wait('exists')
# print(label.get_value())
# static_ele = mainDlg['Static'].get_value()
# print(static_ele)

# # OCR
# pyocr.tesseract.TESSERACT_CMD = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
# tools = pyocr.get_available_tools()
# tool = tools[0]

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

# img2 = Image.open(r'Capture\IR.png')
#
# yaw_coord2 = mf.cal_element_coordinate(2561, 170, 2682, 186)
# pitch_coord2 = mf.cal_element_coordinate(2692, 170, 2812, 186)
#
# (img2.crop([yaw_coord2['x'], yaw_coord2['y'], yaw_coord2['w'], yaw_coord2['h']])
#  .save(r'Capture\Yaw2.png')
# (img2.crop([pitch_coord2['x'], pitch_coord2['y'], pitch_coord2['w'], pitch_coord2['h']])
#  .save(r'Capture\Pitch2.png'))

# img_yaw2: str = 'Capture\Yaw2.png'
# img_pitch2: str = 'Capture\Pitch2.png'
#
# txt_gray_yaw = ct.image_to_string_with_lim(img_yaw2)
# txt_gray_pitch = ct.image_to_string_with_lim(img_pitch2)
#
# txt_blur_yaw = ct.image_to_string_with_blur(img_yaw2)
# txt_blur_pitch = ct.image_to_string_with_blur(img_pitch2)

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

# (ImageGrab.grab(bbox=(cons.eo_power_ui[0], cons.eo_power_ui[1],cons.eo_power_ui[2], cons.eo_power_ui[3]))
#  .save(r'Capture\PowerOff.png'))

# eo_ui_power_coord = mf.cal_element_coordinate(cons.eo_power_ui[0], cons.eo_power_ui[1],
#                                               cons.eo_power_ui[2], cons.eo_power_ui[3])
# eo_power_off_img = Image.open(r'Capture\POWER.png')
# (eo_power_off_img.crop([eo_ui_power_coord['l'], eo_ui_power_coord['t'],
#                       eo_ui_power_coord['r'], eo_ui_power_coord['b']])
#  .save(r'Capture\PowerOff.png'))

area_capture_dic = [{'name': 'Gimbal Power Off', "coord": cons.gimbal_power_ui},
                    {'name': 'EO Power Off', "coord": cons.eo_power_ui},
                    {'name': 'IR Power Off', "coord": cons.ir_power_ui},
                    {'name': 'FAN Power Off', "coord": cons.fan_power_ui},
                    {'name': 'HEATER Power Off', "coord": cons.heater_power_ui},
                    {'name': 'LRF Power Off', "coord": cons.lrf_power_ui}
                    ]
mf.upload_excel_after_check_Report(mainDlg, area_capture_dic)

