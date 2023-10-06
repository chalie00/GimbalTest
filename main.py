import Login_MainFrame

import os
import time
import xlsxwriter
import pandas as pd
import pytesseract

from pywinauto import Application, keyboard, mouse
from pywinauto_recorder.player import *
from PIL import Image

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

mainDlg.child_window(title='닫기', control_type='Button').click()
time.sleep(1)

mainDlg.child_window(title="No", control_type='Button').click()
time.sleep(1)

# Left Menu Hidden/Show
mouse.click(coords=(2264, 179))
mouse.click(coords=(2284, 179))

# Left Menu Select
time.sleep(1)
mainDlg.CheckBox2.click_input()

time.sleep(1)
mainDlg.CheckBox3.click_input()

time.sleep(1)
mainDlg.CheckBox4.click_input()
mainDlg.capture_as_image().save('/Capture/Gimbal.png')

time.sleep(1)
mainDlg.POWER.click_input()

time.sleep(1)

# Screen Shot
mainDlg.capture_as_image().save('/Capture/POWER.png')

# attach an image to excel with xlsxwriter
workbook = xlsxwriter.Workbook('image.xlsx')
worksheet = workbook.add_worksheet()
worksheet.insert_image('B2', 'GIMBAL.png')
worksheet.insert_image('B38', 'POWER.png')
workbook.close()

# Control Excel with Pandas
df = pd.DataFrame({'Data': [10, 20, 30]})
writer = pd.ExcelWriter('Result.xlsx', engine='xlsxwriter')
df.to_excel(writer, sheet_name='Capture')
workbook = writer.book
worksheet = writer.sheets['Capture']
worksheet.insert_image('B2', 'Gimbal.png')
writer.close()
menus = ['EO', 'IR', 'GIMBAL', 'POWER']
Login_MainFrame.capture_Menu(mainDlg, menus)

# Next implementation Get Static Text

static_ele = mainDlg.child_window(title='Static5', class_name="Static")
if static_ele.exists(timeout=10):
    print(static_ele)
label = mainDlg.child_window(class_name="Static", found_index=1).wait('exists')
print(label.get_value())
static_ele = mainDlg['Static'].get_value()
print(static_ele)

# OCR
# pyocr.tesseract.TESSERACT_CMD = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
# tools = pyocr.get_available_tools()
# tool = tools[0]

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
img = Image.open(r'Capture\EO.png')

yaw_coord = Login_MainFrame.cal_element_coordinate(2561, 170, 2682, 186)
pitch_coord = Login_MainFrame.cal_element_coordinate(2692, 170, 2812, 186)

(img.crop([yaw_coord['x'], yaw_coord['y'], yaw_coord['w'], yaw_coord['h']])
 .save(r'Capture\Yaw.png'))
(img.crop([pitch_coord['x'], pitch_coord['y'], pitch_coord['w'], pitch_coord['h']])
 .save(r'Capture\Pitch.png'))

img_yaw: str = 'Capture\Yaw.png'
img_pitch: str = 'Capture\Pitch.png'

txt_msk_yaw = Login_MainFrame.image_to_string_with_hsk(img_yaw)
txt_msk_pitch = Login_MainFrame.image_to_string_with_hsk(img_pitch)

result_yaw = Login_MainFrame.generate_txt_array_with_img(txt_msk_yaw, 'Yaw')
result_pitch = Login_MainFrame.generate_txt_array_with_img(txt_msk_pitch, 'Pitch')

print(result_yaw[0])
print('--------------------------------')
print(result_pitch[0])

img2 = Image.open(r'Capture\IR.png')

yaw_coord2 = Login_MainFrame.cal_element_coordinate(2561, 170, 2682, 186)
pitch_coord2 = Login_MainFrame.cal_element_coordinate(2692, 170, 2812, 186)

(img2.crop([yaw_coord2['x'], yaw_coord2['y'], yaw_coord2['w'], yaw_coord2['h']])
 .save(r'Capture\Yaw2.png'))
(img2.crop([pitch_coord2['x'], pitch_coord2['y'], pitch_coord2['w'], pitch_coord2['h']])
 .save(r'Capture\Pitch2.png'))

img_yaw2: str = 'Capture\Yaw2.png'
img_pitch2: str = 'Capture\Pitch2.png'

txt_gray_yaw = Login_MainFrame.image_to_string_with_lim(img_yaw2)
txt_gray_pitch = Login_MainFrame.image_to_string_with_lim(img_pitch2)

txt_blur_yaw = Login_MainFrame.image_to_string_with_blur(img_yaw2)
txt_blur_pitch = Login_MainFrame.image_to_string_with_blur(img_pitch2)

