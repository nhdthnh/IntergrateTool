from tkinter import *
import tkinter as tk
from tkinter import font,ttk,Label
import os
from Support_Function import *
import datetime
import sys
import pandas as pd
import threading
from openpyxl import load_workbook
# Path Config
USS_CENTRALIZATION_PATH         = r'\\bosch.com\dfsRB\DfsVN\LOC\Hc\RBVH\20_EDA\10_EDA9\01_Internal\10_Tool\00_ToolChain_Template'
    # Path Logo
LOG_PATH                        = os.path.join(USS_CENTRALIZATION_PATH, 'Log')
LOGO_PATH                       = os.path.join(USS_CENTRALIZATION_PATH, 'Icon')

# DBC_GENERATION_PATH             = os.path.join(USS_CENTRALIZATION_PATH,'1.DBC_Generation')

# IMPORT_DBC_PATH                 = os.path.join(USS_CENTRALIZATION_PATH,'2.Import_DBC')
# IMPORT_DBC_FOLDER_PATH          = os.path.join(IMPORT_DBC_PATH, 'Import DBC')
# IMPORT_DBC_LINK_PATH            = os.path.join(IMPORT_DBC_PATH,'README.txt')

# VSM_GENERATION_PATH             = os.path.join(USS_CENTRALIZATION_PATH,'3.VSM_Generation')
# VSM_GENERATION_LINK_PATH        = os.path.join(VSM_GENERATION_PATH,'README.txt')

# TX_OFFSETTUNNING_PATH           = os.path.join(USS_CENTRALIZATION_PATH,'4.Tx_OffsetTunning')
# TX_OFFSETTUNNING_LINK_PATH      = os.path.join(TX_OFFSETTUNNING_PATH,'README.txt')

# LAB_MONITORING_PATH             = os.path.join(USS_CENTRALIZATION_PATH,'5.Lab_Remote')
# VSM_CONVERSION_ASW_PATH         = os.path.join(USS_CENTRALIZATION_PATH,'6.VSM_ConversionASW')
# OPL_TEMPLATE_PATH               = os.path.join(USS_CENTRALIZATION_PATH,'7.OPL_Template')
# REVIEW_CHECKLIST_PATH           = os.path.join(USS_CENTRALIZATION_PATH,'8.Review_Checklist')
# FH_DESIGN_TEMPLATE_PATH         = os.path.join(USS_CENTRALIZATION_PATH,'9.FH_Template')
# RESOURCE_MEASUREMENT_PATH       = os.path.join(USS_CENTRALIZATION_PATH,'10.ResourceManagement_Template')
# DIAG_COMMON_UNDERSTANDING_PATH  = os.path.join(USS_CENTRALIZATION_PATH,'11.DiagCommonUnderstanding')
# FIM_DESIGN_TEMPLATE_PATH        = os.path.join(USS_CENTRALIZATION_PATH,'12.FiM_Design_Template')

# RUNTIME_CHECKLIST_PATH          = os.path.join(USS_CENTRALIZATION_PATH,'13.Runtime_Checklist')
# RUNTIME_CHECKLIST_LINK_PATH     = os.path.join(RUNTIME_CHECKLIST_PATH,'README.txt')

# DOORS_UPDATE_PATH               = os.path.join(USS_CENTRALIZATION_PATH,'14.Doors_update_requirement_linking')

# FAQ_PATH                        = os.path.join(USS_CENTRALIZATION_PATH,'FAQ')
# FAQ_LINK_PATH                   = os.path.join(FAQ_PATH,'README.txt')

TOOLS_VERSION_PATH              = os.path.join(USS_CENTRALIZATION_PATH,'Tools_Version')
USS_KIT_VERSION_PATH            = os.path.join(TOOLS_VERSION_PATH,'USS_KIT_Version.txt')

USS_KIT                         = os.path.join(USS_CENTRALIZATION_PATH,'USS_KIT')

CONFIG_PATH                     = os.path.join(USS_CENTRALIZATION_PATH,'Config')
BUTTON_CONFIG_XLSX_PATH         = os.path.join(os.getcwd(),'Button_Config.xlsx')
STATE_CONFIG_PATH               = os.path.join(os.getcwd(),'State.txt')
SW_VER = 'SW_1.0'
def PersonalGUI():
    def button_click(button_text):
        # Xử lý sự kiện khi một nút được nhấn
        print(f"Nút {button_text} đã được nhấn")

    # Đọc dữ liệu từ tệp Excel
    file_path = BUTTON_CONFIG_XLSX_PATH  # Thay 'ten_file.xlsx' bằng đường dẫn thực tế của bạn
    try:
        df = pd.read_excel(file_path, sheet_name='Open_Folder',usecols=['button_name'])
    except pd.errors.EmptyDataError:
        # Trường hợp tệp Excel rỗng, tạo DataFrame trống
        df = pd.DataFrame()

    # Tạo cửa sổ
    window = tk.Tk()
    window.title("GUI với Nút 4x3 (Tự động thêm nút từ Excel)")

    # Tạo lưới nút từ dữ liệu trong DataFrame
    buttons = df.values.tolist() if not df.empty else []

    def update_gui():
        for i in range(len(buttons)):
            # Tính toán vị trí dòng và cột dựa trên số cột tối đa (max_columns)
            row = i // max_columns
            col = i % max_columns

        
            button_value = buttons[i][0] if isinstance(buttons[i], list) else buttons[i]

            button_text = str(button_value)  # Chuyển đổi giá trị thành chuỗi
            btn = tk.Button(window, text=button_text, padx=20, pady=20, command=lambda text=button_text: button_click(text))
            btn.grid(row=row, column=col)
    max_columns = 4


    # Cập nhật giao diện ban đầu
    update_gui()
    window.mainloop()
PersonalGUI()