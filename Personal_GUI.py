import tkinter as tk
import pandas as pd
import os

from tkinter import font
from tkinter import *
from tkinter import ttk
from tkinter import colorchooser
from tkinter import messagebox
BUTTON_CONFIG_XLSX_PATH = os.path.join(os.getcwd(), "Button_Config.xlsx")

def PersonalGUI():
    def add_button_popup():
        def colorpicker_click():
            c_code = colorchooser.askcolor()[1]  # Lấy mã màu hex từ tuple trả về
            color_var.set(c_code)

        popup = tk.Toplevel(window)
        popup.title("ADD BUTTON")

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

        label1 = tk.Label(popup, text="Button Name:")
        label1.grid(row=0, column=0, padx=5, pady=5)
        entry1 = tk.Entry(popup)
        entry1.grid(row=0, column=1, padx=5, pady=5)

        label2 = tk.Label(popup, text="Path:")
        label2.grid(row=1, column=0, padx=5, pady=5)
        entry2 = tk.Entry(popup)
        entry2.grid(row=1, column=1, padx=5, pady=5)

        color_label = tk.Label(popup, text="Color:")
        color_label.grid(row=2, column=0, padx=5, pady=5)
        color_var = tk.StringVar()
        color_var.set("#000000")  
        color_entry = tk.Entry(popup, textvariable=color_var, state='readonly')
        color_entry.grid(row=2, column=1, padx=5, pady=5)
        color_button = tk.Button(popup, text="Color", command=colorpicker_click)
        color_button.grid(row=2, column=2, padx=5, pady=5)

        save_button_button = tk.Button(popup, text="SAVE", command=save_button)
        save_button_button.grid(row=3, column=0, columnspan=3, pady=10)
   
  
    def remove_selected_button():
        selected_button = combo_var.get()
        if not selected_button:
            messagebox.showerror("Error", "Please select a button to remove.")
            return

        confirmation = messagebox.askquestion("Remove Button", f"Are you sure you want to remove the button '{selected_button}'?", icon='warning')
        if confirmation == 'yes':
            df.drop(df[df['button_name'] == selected_button].index, inplace=True)
            df.to_excel(BUTTON_CONFIG_XLSX_PATH, sheet_name='Open_Folder', index=False)
            update_gui()
            remove_window.destroy()

        remove_window = tk.Toplevel()  # Tạo cửa sổ riêng biệt
        remove_window.title("Remove Button")

        combo_var = tk.StringVar()
        combo_box = ttk.Combobox(remove_window, textvariable=combo_var, values=buttons, state="readonly")
        combo_box.grid(row=0, column=0, padx=10, pady=10)
        combo_box.set("Select Button")

        remove_button_btn = tk.Button(remove_window, text="Remove", command=remove_selected_button)
        remove_button_btn.grid(row=0, column=1, padx=10, pady=10)
    def button_click(button_text):
        file_path = df.loc[df['button_name'] == button_text, 'path'].values[0]
        os.startfile(file_path) 
        execution= file_path.split("\\")[-1]
        last_file_label = tk.Label(window, text=f"Open: {execution}", font=custom_font)
        last_file_label.grid(row=(len(buttons) // max_columns) + 2, column=0, columnspan=max_columns, pady=10)
        window.after(2000, lambda: last_file_label.config(text=""))
    file_path = BUTTON_CONFIG_XLSX_PATH
    try:
        df = pd.read_excel(file_path, sheet_name='Open_Folder', usecols=['button_name', 'path', 'color'])
    except pd.errors.EmptyDataError:
        df = pd.DataFrame()

    window = tk.Tk()
    window.title("Personal GUI")

    buttons = df.values.tolist() if not df.empty else []
    custom_font = font.Font(family="Arial", size=15, weight="bold")
    def update_gui():
        for i in range(len(buttons)):
            row = (i // max_columns)
            col = i % max_columns
            button_value = buttons[i][0] if isinstance(buttons[i], list) else buttons[i]
            button_text = str(button_value)
            button_color = buttons[i][2] if len(buttons[i]) > 2 else "#000000"  
            btn = tk.Button(window, text=button_text, width=15, height=5, command=lambda text=button_text: button_click(text), bg=button_color, font = custom_font)
            btn.grid(row=row, column=col)
            remove_button_btn = tk.Button(window, text="Remove", command=remove_selected_button)
            remove_button_btn.grid(row=(len(buttons) // max_columns) + 1, column=len(buttons) % max_columns + 3, padx=10, pady=10)
        add_button = tk.Button(window, text="ADD BUTTON", command=add_button_popup)
        add_button.grid(row=(len(buttons) // max_columns) + 1, column=len(buttons) % max_columns, padx=10, pady=10)
       
        
    max_columns = 4
    update_gui()
    window.mainloop()

PersonalGUI()
