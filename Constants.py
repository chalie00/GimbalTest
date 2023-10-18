import pywinauto
from datetime import date, datetime

excel_No = 1

today = date.today()
today_time_now = datetime.now()
today_time_format = today_time_now.strftime("%b-%d-%Y-%H-%M-%S")
today_format = today.strftime("%b-%d-%Y")

# Report File Constants
report_path = rf'TestResult\Test Report_{today_time_format}.xlsx'
user_info_title = ['Category', 'Information']
user_info_title_pos = {'column': 2, 'row': 6}
user_data_area = {'start': 'B6', 'end': 'C10'}

product_info_title = ['Product_Name', 'Firmware']
product_info_title_pos = {'column': 5, 'row': 6}
product_data_area = {'start': 'E6', 'end': 'F12'}

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
