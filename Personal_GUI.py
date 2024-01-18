import tkinter as tk
import pandas as pd
import os

from tkinter import *
from tkinter import colorchooser

BUTTON_CONFIG_XLSX_PATH = os.path.join(os.getcwd(), "Button_Config.xlsx")

def PersonalGUI():
    def add_button_popup():
        def colorpicker_click():
            c_code = colorchooser.askcolor()[1]  # Lấy mã màu hex từ tuple trả về
            color_var.set(c_code)

        popup = tk.Toplevel(window)
        popup.title("Thêm Nút")

        def save_button():
            new_button_name = entry1.get()
            new_button_path = entry2.get()
            button_color = color_var.get()

            if new_button_name and new_button_path and button_color:
                df.loc[len(df)] = [new_button_name, new_button_path, button_color]
                df.to_excel(BUTTON_CONFIG_XLSX_PATH, sheet_name='Open_Folder', index=False)
                update_gui()
                popup.destroy()
            else:
                tk.messagebox.showerror("ERROR", "Please enter information.")

        label1 = tk.Label(popup, text="button_name:")
        label1.grid(row=0, column=0, padx=5, pady=5)
        entry1 = tk.Entry(popup)
        entry1.grid(row=0, column=1, padx=5, pady=5)

        label2 = tk.Label(popup, text="Path:")
        label2.grid(row=1, column=0, padx=5, pady=5)
        entry2 = tk.Entry(popup)
        entry2.grid(row=1, column=1, padx=5, pady=5)

        color_label = tk.Label(popup, text="Màu:")
        color_label.grid(row=2, column=0, padx=5, pady=5)
        color_var = tk.StringVar()
        color_var.set("#000000")  # Màu mặc định là đen
        color_entry = tk.Entry(popup, textvariable=color_var, state='readonly')
        color_entry.grid(row=2, column=1, padx=5, pady=5)
        color_button = tk.Button(popup, text="Color", command=colorpicker_click)
        color_button.grid(row=2, column=2, padx=5, pady=5)

        save_button_button = tk.Button(popup, text="SAVE", command=save_button)
        save_button_button.grid(row=3, column=0, columnspan=3, pady=10)

    def button_click(button_text):
        file_path = df.loc[df['button_name'] == button_text, 'path'].values[0]
        os.startfile(file_path)
       

    file_path = BUTTON_CONFIG_XLSX_PATH
    try:
        df = pd.read_excel(file_path, sheet_name='Open_Folder', usecols=['button_name', 'path', 'color'])
    except pd.errors.EmptyDataError:
        df = pd.DataFrame()

    window = tk.Tk()
    window.title("Personal GUI")

    buttons = df.values.tolist() if not df.empty else []

    def update_gui():
        for i in range(len(buttons)):
            row = (i // max_columns)
            col = i % max_columns
            button_value = buttons[i][0] if isinstance(buttons[i], list) else buttons[i]
            button_text = str(button_value)
            button_color = buttons[i][2] if len(buttons[i]) > 2 else "#000000"  # Màu mặc định là đen
            btn = tk.Button(window, text=button_text, width=15, height=5, command=lambda text=button_text: button_click(text), bg=button_color)
            btn.grid(row=row, column=col)

        add_button = tk.Button(window, text="ADD BUTTON", command=add_button_popup)
        add_button.grid(row=(len(buttons) // max_columns) + 1, column=len(buttons) % max_columns, padx=10, pady=10)

    max_columns = 4
    update_gui()
    window.mainloop()

PersonalGUI()
