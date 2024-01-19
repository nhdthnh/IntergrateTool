
import pandas as pd
import os
from tkinter import font
from tkinter import *
import tkinter as tk
from tkinter import ttk
from tkinter import colorchooser
from tkinter import messagebox
BUTTON_CONFIG_XLSX_PATH_PERSONAL = os.path.join(os.getcwd(), "share/Button_Config.xlsx")

def PersonalGUI():
    file_path = BUTTON_CONFIG_XLSX_PATH_PERSONAL
    df = pd.read_excel(file_path)
    max_rows = int(df['max_row'].values[0])
    max_columns = int(df['max_column'].values[0])
    if df['max_column'].values[0] == None or df['max_row'].values[0] == None:
        df.at[0, 'max_row'] = 5
        df.at[0, 'max_column'] = 5
        df.to_excel(file_path, sheet_name='Sheet1', index=False)
    ##ADD#######################
    def add_button_popup():
        def colorpicker_click():
            c_code = colorchooser.askcolor()[1] 
            color_var.set(c_code)
        popup = tk.Toplevel(window)
        popup.title("ADD BUTTON")
        def save_button():
            new_button_name = entry1.get()
            new_button_path = entry2.get()
            button_color = color_var.get()
            if new_button_name and new_button_path and button_color:
                if button_color == "#000000" or button_color == "#ffffff":
                    tk.messagebox.showerror("ERROR", "Color is not supported")
                    popup.mainloop()
                else:
                    df.loc[len(df)] = [new_button_name, new_button_path, button_color]
                    df.to_excel(file_path, sheet_name='Sheet1', index=False)
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
    #EDIT################################
    def edit():
        def colorpicker_click():
            c_code = colorchooser.askcolor()[1] 
            color_var.set(c_code)
        def confirm_edit():
            selected_button = combo.get()
            if selected_button == "Choose Button":
                messagebox.showwarning("Warning", "Please choose a button.")
            else:
                df = pd.read_excel(file_path)
                button_info = df[df['button_name'] == selected_button].iloc[0]

                # Cập nhật nội dung của Label để hiển thị thông tin
                label_info.config(text=f"Button Name:    {button_info['button_name']}\nPath:    {button_info['path']}\nColor:    {button_info['color']}")
                button_color = entry_color.get()            
                new_button_name = entry_button_name.get()
                new_button_path = entry_path.get()
                df.loc[df['button_name'] == selected_button, 'button_name'] = new_button_name
                df.loc[df['button_name'] == selected_button, 'path'] = new_button_path
                df.loc[df['button_name'] == selected_button, 'color'] = button_color
                print(df.loc[df['button_name'] == selected_button, 'color'])
                df.to_excel(file_path, sheet_name='Sheet1', index=False)
            
        color_var = tk.StringVar()
        color_var.set("#000000")  
        edit_window = tk.Toplevel()
        edit_window.title("Remove Button")
        edit_window.geometry("350x350")
        combo = ttk.Combobox(edit_window, values=df['button_name'].tolist())
        combo.set("Choose Button")
        combo.grid(row=0, column=0, pady=20)
        def on_combobox_selected(event):
            confirm_edit()
        combo = ttk.Combobox(edit_window, values=df['button_name'].tolist())
        combo.set("Choose Button")
        combo.grid(row=0, column=0, pady=20)
        combo.bind("<<ComboboxSelected>>", on_combobox_selected)
        label_info = tk.Label(edit_window, text="")
        label_info.grid(row=1, column=0, pady=10, columnspan=3)
        
        # Tạo Entry cho Button Name
        label_button_name = tk.Label(edit_window, text="Button Name:")
        label_button_name.grid(row=2, column=0, pady=10)
        entry_button_name = tk.Entry(edit_window)
        entry_button_name.grid(row=2, column=4, pady=10)

        # Tạo Entry cho Path
        label_path = tk.Label(edit_window, text="Path:")
        label_path.grid(row=3, column=0, pady=10)
        entry_path = tk.Entry(edit_window)
        entry_path.grid(row=3, column=4, pady=10)

        # Tạo Entry cho Color
        label_color = tk.Label(edit_window, text="Color:")
        label_color.grid(row=4, column=0, pady=10)
        entry_color = tk.Entry(edit_window, textvariable=color_var)
        entry_color.grid(row=4, column=4, pady=10)
        label_color_text = tk.Button(edit_window, text="COLOR",command = colorpicker_click)
        label_color_text.grid(row=4, column=6, pady=10)
        # Thêm hàm confirm_edit vào sự kiện của nút "Confirm Edit"
        confirm_edit_button = tk.Button(edit_window, text="Edit", command=confirm_edit)
        confirm_edit_button.grid(row=5, column=0, pady=20, columnspan=3)
    #CONFIGURE##########################
    def configure():
        def save():
            max_column_value = entryColumn.get()
            max_row_value = entryRow.get()
            df.at[0, 'max_row'] = max_row_value
            df.at[0, 'max_column'] = max_column_value
            df.to_excel(file_path, sheet_name='Sheet1', index=False)
            configure.destroy()
        configure = tk.Tk()
        configure.geometry("300x300")
        configure.title("Configure")
        MaxColumnLabel = tk.Label(configure, text="Max column: ",font= custom_font)
        MaxColumnLabel.pack()
        entryColumn = tk.Entry(configure,font="Arial 13")
        entryColumn.pack()
        MaxRowLabel = tk.Label(configure, text="Max column: ",font= custom_font)
        MaxRowLabel.pack()
        entryRow = tk.Entry(configure, font="Arial 13")
        entryRow.pack()
        SaveButton = tk.Button(configure, text="SAVE", command=save,font= "Arial 15")
        SaveButton.pack()
        configure.mainloop()
    #REMOVE##################
    def remove_selected_button():
        def confirm_removal():
            selected_button = combo.get()
            messagebox.askquestion("Warning",f"Confirm delete button {selected_button}")
            df = pd.read_excel(file_path, index_col='button_name')  
            df = df.drop(index=selected_button)
            df.to_excel(file_path)
            remove_window.destroy()
        remove_window = tk.Toplevel()
        remove_window.title("Remove Button")
        remove_window.geometry("150x150")
        combo = ttk.Combobox(remove_window, values=df['button_name'].tolist())
        combo.set("Choose Button")
        combo.grid(row=0, column=0, pady=20)
        confirm_button = tk.Button(remove_window, text="Confirm", command=confirm_removal)
        confirm_button.grid(row=1, column=0, pady=10)
    def refesh():
        window.destroy()
        PersonalGUI()
    def button_click(button_text):
        file_path = df.loc[df['button_name'] == button_text, 'path'].values[0]
        if os.path.exists(file_path):
            os.startfile(file_path) 
            execution= file_path.split("\\")[-1]
            last_file_label = tk.Label(left_column_frame, text=f"Open: {execution}", font=custom_font)
            last_file_label.grid(row=(len(buttons) // max_columns) + 2, column=0, columnspan=max_columns, pady=10)
            window.after(2000, lambda: last_file_label.config(text=""))
        else:
            messagebox.showerror("Error", "Cannot find the file")
            window.mainloop()
    
    try:
        df = pd.read_excel(file_path, sheet_name='Sheet1', usecols=['button_name', 'path', 'color'])
    except pd.errors.EmptyDataError:
        df = pd.DataFrame()

    window = tk.Tk()
    window.title("Personal GUI")
    window.resizable(False, False)  
    
    buttons = df.values.tolist() if not df.empty else []
    left_column_frame = tk.Frame(window)
    left_column_frame.grid(row=0, column=0, rowspan=max_rows)
    custom_font = font.Font(family="Arial", size=15, weight="bold")
    def update_gui():
        for i in range(len(buttons)):
            if i // max_columns >= max_rows:
                break
            row = i // max_columns
            col = i % max_columns
            button_value = buttons[i][0] if isinstance(buttons[i], list) else buttons[i]
            button_text = str(button_value)
            button_color = buttons[i][2] if len(buttons[i]) > 2 else "#000000" 
            def on_enter(event):
                event.widget.config(bg="white")  
                event.widget.config(cursor="hand2")
            def on_leave(event):
                event.widget['background'] = event.widget.default_bg 
            btn = tk.Button(window, text=button_text, width=15, height=5, command=lambda text=button_text: button_click(text), bg=button_color, font="Arial 15")
            btn.default_bg = button_color
            btn.bind("<Enter>", on_enter)
            btn.bind("<Leave>", on_leave)
            btn.grid(row=row, column=col + 1) 
    
    add_button_left = tk.Button(left_column_frame, text="ADD BUTTON", command=add_button_popup)
    add_button_left.grid(row=0, column=0, padx=10, pady=10)
    remove_button_left = tk.Button(left_column_frame, text="REMOVE BUTTON", command=remove_selected_button)
    remove_button_left.grid(row=1, column=0, padx=10, pady=10)
    refresh_button_left = tk.Button(left_column_frame, text="REFRESH", command=refesh)
    refresh_button_left.grid(row=2, column=0, padx=10, pady=10)
    config = tk.Button(left_column_frame, text="CONFIGURE", command=configure)
    config.grid(row=3, column=0, padx=10, pady=10)
    edit = tk.Button(left_column_frame, text="EDIT BUTTON",command = edit)
    edit.grid(row=4, column=0, padx=10, pady=10)
    Row = tk.Label(left_column_frame, text=f"Max rows: {max_rows}")
    Row.grid(row=5, column=0, padx=10, pady=10)
    Column = tk.Label(left_column_frame, text=f"Max columns: {max_columns}")
    Column.grid(row=6, column=0, padx=10, pady=10)
    update_gui()
      
    window.mainloop()
PersonalGUI()
