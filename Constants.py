import pywinauto
from datetime import date, datetime

from openpyxl.styles import Font

excel_No = 1

today = date.today()
today_time_now = datetime.now()
current_time_bdYHMS = today_time_now.strftime("%b-%d-%Y-%H-%M-%S")
current_only_time = today_time_now.strftime("%H-%M-%S")
today_format = today.strftime("%b-%d-%Y")

# Report File Constants
# TODO (Com) - Modify a Save Path
report_path = rf'TestResult\Test Report_{current_time_bdYHMS}.xlsx'
report_path_vba = rf'TestResult\vba_macro.xlsm'
report_title_font = Font(
    name='맑은 고딕',
    bold=True,
    size=24
)
report_title_info = {'title': 'TEST REPORT',
                     'start_col': 2, 'start_row': 2,
                     'end_col': 9, 'end_row': 4,
                     'align': 'center',
                     'border': 'thin'
                     }

user_info_title = ['Category', 'Information']
user_info_title_pos = {'column': 3, 'row': 6}
user_data_area = {'start': 'C6', 'end': 'D10'}

product_info_title = ['Product_Name', 'Firmware']
product_info_title_pos = {'column': 3, 'row': 12}
product_data_area = {'start': 'C12', 'end': 'D18'}

test_title_font = Font(
    name='맑은 고딕',
    bold=True,
)
test_pic_pos = {'col': 2, 'row': 20, 'value': 'PICTURE'}
test_code_pos = {'col': test_pic_pos['col'] + 1,
                 'row': 20, 'value': 'CODE'}
test_dep_pos = {'col': test_code_pos['col'] + 1,
                'row': test_code_pos['row'], 'value': 'DEPTH'}
test_pre_pos = {'col': test_code_pos['col'] + 2,
                'row': test_code_pos['row'], 'value': 'PRE STATE'}
test_step_pos = {'col': test_code_pos['col'] + 3,
                 'row': test_code_pos['row'], 'value': 'TEST STEP'}
test_expect_pos = {'col': test_code_pos['col'] + 5,
                   'row': test_code_pos['row'], 'value': 'EXPECT RESULT'}
test_result_pos = {'col': test_code_pos['col'] + 7,
                   'row': test_code_pos['row'], 'value': 'RESULT'}
test_time_pos = {'col': test_code_pos['col'] + 8,
                 'row': test_code_pos['row'], 'value': 'TIME(H-M-S)'}

# Element Coordinate
Dig_zero_Coordinate = [2240, 165, 3520, 885]
left_menu_hide = [2264, 179]
left_menu_show = [2284, 179]
aspect_ratio = [3150, 177]
yaw_lt_rb = [2561, 170, 2682, 186]
pitch_lt_rb = [2692, 170, 2812, 186]

# POWER Tab (Power Control Setting)
eo_power_on = [2404, 328]
eo_power_off = [2479, 328]

# Display Power UI
gimbal_power_ui = [3293, 757, 3358, 795]
eo_power_ui = [3365, 755, 3428, 794]
ir_power_ui = [3435, 754, 3504, 794]
fan_power_ui = [3293, 818, 3355, 852]
heater_power_ui = [3365, 817, 3429, 851]
lrf_power_ui = [3438, 815, 3502, 852]

# Virtual Controller UI
vj_center = [2316, 805]
vj_up = [2316, 768]
vj_down = [2316, 842]
vj_left = [2281, 803]
vj_right = [2352, 803]