
import pandas as pd
import os
from tkinter import font,ttk,Label
from tkinter import *
import tkinter as tk
from tkinter import ttk
from tkinter import colorchooser
from tkinter import messagebox
import webbrowser
from urllib.parse import urlparse

BUTTON_CONFIG_XLSX_PATH_PERSONAL = os.path.join(os.getcwd(), "share/Button_Config.xlsx")
CONFIGURE                        = os.path.join(os.getcwd(),"share/configure.txt")
def PersonalGUI():

    
    
    def on_enter(e):
    # Get current color
        bg = e.widget['background']
        # Convert to RGB
        r, g, b = window.winfo_rgb(bg)
        # Improve brightness, rgb 16 bit
        r, g, b = min(r * 1.5, 65535), min(g * 1.5, 65535), min(b * 1.5, 65535)
        # Scale 0-255,rgb 8 bit
        e.widget['background'] = '#%02x%02x%02x' % (int(r // 256), int(g // 256), int(b // 256))
        e.widget.config(cursor="hand2")
    # Function not hover
    def on_leave(e):
        e.widget['background'] = e.widget.default_bg 
    
    def on_enter_left(e):
    # Get current color
        bg = e.widget['background']
        e.widget['background'] = '#f5fabb'
        e.widget.config(cursor="hand2")
    def on_leave_left(e):
        e.widget['background'] = '#ffd166'
    file_path = BUTTON_CONFIG_XLSX_PATH_PERSONAL
    file_txt = CONFIGURE
    def is_url(text):
        try:
            result = urlparse(text)
            return all([result.scheme, result.netloc])
        except ValueError:
            return False
    df = pd.read_excel(file_path)
    with open(file_txt, "r", encoding="utf-8") as file:
        lines = file.readlines()
        if len(lines) >= 2:
            max_rows = int(lines[0].strip())
            max_columns = int(lines[1].strip()) 
    # if df['max_column'].values[0] == None or df['max_row'].values[0] == None:
    #     df.at[0, 'max_row'] = 5
    #     df.at[0, 'max_column'] = 5
    #     df.to_excel(file_path, sheet_name='Sheet1', index=False)
    ##ADD#######################
    def add_button_popup():
        def colorpicker_click():
            c_code = colorchooser.askcolor()[1] 
            color_var.set(c_code) 
            color_entry.delete(0, tk.END)
            color_entry.insert(0,c_code)
            popup.lift()
            popup.attributes('-topmost', True)
            popup.attributes('-topmost', False)
        popup = tk.Toplevel(window)
        popup.title("ADD BUTTON")
        def save_button():
            new_button_name = entry1.get()
            new_button_path = entry2.get()
            button_color = color_entry.get()
            columns = int(entry_colums.get())
            rows = int(entry_rows.get())
            print(type(columns))
            print(type(rows))
            if new_button_name and new_button_path and button_color and columns and rows:
                if button_color == "#000000" or button_color == "#ffffff":
                    tk.messagebox.showerror("ERROR", "Color is not supported")
                    popup.mainloop()
                else:
                    df.loc[len(df)] = [str(new_button_name), str(new_button_path), str(button_color), columns, rows]
                    df.to_excel(file_path, sheet_name='Sheet1', index=False)
                    update_gui()
                    popup.destroy()
                    messagebox.showinfo("","Add button successfully")
                    window.destroy()
                    PersonalGUI()
            else:
                tk.messagebox.showerror("ERROR", "Please enter information.")
        
        label1 = tk.Label(popup, text="Button Name:")
        label1.grid(row=0, column=0, padx=5, pady=5)
        entry1 = tk.Entry(popup)
        entry1.grid(row=0, column=1, padx=5, pady=5)

        label2 = tk.Label(popup, text="Path/URL:")
        label2.grid(row=1, column=0, padx=5, pady=5)
        entry2 = tk.Entry(popup)
        entry2.grid(row=1, column=1, padx=5, pady=5)

        color_label = tk.Label(popup, text="Color:")
        color_label.grid(row=2, column=0, padx=5, pady=5)
        color_var = tk.StringVar()
        color_var.set('')
        color_entry = tk.Entry(popup)
        color_entry.grid(row=2, column=1, padx=5, pady=5)
        color_button = tk.Button(popup, text="Color", command=colorpicker_click)
        color_button.grid(row=2, column=2, padx=5, pady=5)
        label_colums = tk.Label(popup, text="Columns:")
        label_colums.grid(row=3, column=0, padx=5, pady=5)
        entry_colums = tk.Entry(popup)
        entry_colums.grid(row=3, column=1, padx=5, pady=5)
        label_rows = tk.Label(popup, text="Rows:")
        label_rows.grid(row=4, column=0, padx=5, pady=5)
        entry_rows = tk.Entry(popup)
        entry_rows.grid(row=4, column=1, padx=5, pady=5)
        color_entry = tk.Entry(popup)
        color_entry.grid(row=2, column=1, padx=5, pady=5)
        save_button_button = tk.Button(popup, text="SAVE", command=save_button)
        save_button_button.grid(row=5, column=0, columnspan=3, pady=10)
    #EDIT################################
    def edit():
        def colorpicker_click():
            c_code = colorchooser.askcolor()[1] 
            color_var.set(c_code)
            edit_window.lift()
            edit_window.attributes('-topmost', True)
            edit_window.attributes('-topmost', False)
        def show():
            selected_button = combo.get()
            if selected_button == "Choose Button":
                messagebox.showwarning("Warning", "Please select button")
            else:
                df = pd.read_excel(file_path)
                button_info = df[df['button_name'] == selected_button].iloc[0]
                #label_info.config(text=f"Button Name: {button_info['button_name']}\n Path: {button_info['path']}\nColor: {button_info['color']}")
                entry_button_name.delete(0, tk.END)
                entry_path.delete(0, tk.END)
                entry_color.delete(0, tk.END)
                entry_button_name.insert(0, button_info["button_name"])
                entry_path.insert(0, button_info["path"])
                entry_color.insert(0, button_info["color"])
        def save():
            selected_button = combo.get()
            if selected_button == "Choose Button":
                messagebox.showwarning("Warning", "Please select button")
            else:
                df = pd.read_excel(file_path)
                new_button_name = entry_button_name.get()
                new_button_path = entry_path.get()
                new_button_color = entry_color.get()
                df.loc[df['button_name'] == selected_button, 'color'] = new_button_color
                df.loc[df['button_name'] == selected_button, 'button_name'] = new_button_name
                df.loc[df['button_name'] == selected_button, 'path'] = new_button_path
                df.to_excel(file_path)
                messagebox.showinfo("", "Update button successfully")
                edit_window.destroy()
                window.destroy()
                PersonalGUI()
        color_var = tk.StringVar()
        color_var.set("")  
        edit_window = tk.Toplevel()
        edit_window.title("Edit Button")
        edit_window.resizable(False, False)
        def on_combobox_selected(event):
            show()
        combo = ttk.Combobox(edit_window, values=df['button_name'].tolist())
        combo.set("Choose Button")
        combo.grid(row=0, column=0, columnspan=5,pady=20)
        combo.bind("<<ComboboxSelected>>", on_combobox_selected)
        label_button_name = tk.Label(edit_window, text="Button Name:")
        label_button_name.grid(row=1, column=0, pady=10)
        entry_button_name = tk.Entry(edit_window)
        entry_button_name.grid(row=1, column=2, pady=10)
        label_path = tk.Label(edit_window, text="Path/URL:")
        label_path.grid(row=2, column=0, pady=10)
        entry_path = tk.Entry(edit_window)
        entry_path.grid(row=2, column=2, pady=10)
        label_color = tk.Label(edit_window, text="Color:")
        label_color.grid(row=3, column=0, pady=10)
        entry_color = tk.Entry(edit_window, textvariable=color_var)
        entry_color.grid(row=3, column=2, pady=10)
        button_color_text = tk.Button(edit_window, text="COLOR",command = colorpicker_click)
        button_color_text.grid(row=3, column=4, pady=10)
        save_button = tk.Button(edit_window,text = "SAVE",command = save)
        save_button.grid(row=4, column=0, pady=10)
    #CONFIGURE##########################
    def configure():
        def save():
            max_column_value = entryColumn.get()
            max_row_value = entryRow.get()
            with open(file_txt,'w') as file:
                file.write(f"{max_row_value}\n{max_column_value}")
            configure.destroy()
            messagebox.showinfo("","Edit window successfully")
            window.destroy()
            PersonalGUI()
        configure = tk.Tk()
        configure.title("Configure")
        MaxColumnLabel = tk.Label(configure, text="Max column: ",font= custom_font)
        MaxColumnLabel.pack()
        entryColumn = tk.Entry(configure,font="Arial 13")
        entryColumn.pack()
        MaxRowLabel = tk.Label(configure, text="Max row: ",font= custom_font)
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
            print(df)
            remove_window.destroy()
            messagebox.showinfo("","Remove button successfully")
            window.destroy()
            PersonalGUI()
        remove_window = tk.Toplevel()
        remove_window.title("Remove Button")
        # remove_window.geometry("250x250")
        combo = ttk.Combobox(remove_window, values=df['button_name'].tolist())
        combo.set("Choose Button")
        combo.grid(row=0, column=0, pady=20)
        confirm_button = tk.Button(remove_window, text="Confirm", command=confirm_removal)
        confirm_button.grid(row=1, column=0, pady=10)
    # def refesh():
    #     window.destroy()
    #     PersonalGUI()
    def button_click(button_text):
        file_path = df.loc[df['button_name'] == button_text, 'path'].values[0]
        if os.path.exists(file_path):                
            os.startfile(file_path) 
            execution= file_path.split("\\")[-1]
            last_file_label = tk.Label(left_column_frame, text=f"Open: {execution}", font="Arial 10", bg = '#0d6759',fg='white')
            last_file_label.grid(row=8, column=0, pady=10)
            window.after(2000, lambda: last_file_label.config(text=""))
        elif is_url(file_path):
            webbrowser.open(file_path, new=0, autoraise=True)
            url = file_path.strip().split("/")[2]
            last_file_label = tk.Label(left_column_frame, text=f"Open: {url}", font="Arial 10", bg = '#0d6759',fg='white')
            last_file_label.grid(row=8, column=0, pady=10)
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
    window.lift()
    window.attributes('-topmost', True)
    window.attributes('-topmost', False)
    buttons = df.values.tolist() if not df.empty else []
    left_column_frame = Canvas(window, width=60,bg="#0d6759",highlightthickness=0,relief='raised', borderwidth=5)
    left_column_frame.grid(row=1, column=0, rowspan=max_rows+1, sticky='nsew')
    custom_font = font.Font(family="Arial", size=15, weight="bold")
    title = Label(window, text="My Tools", bg="#0d6759", fg="white", font = font.Font(size=13, weight= 'bold'),height= 2, relief='raised', borderwidth=5)
    title.grid(row=0, column=0, columnspan=max_columns+1,sticky='ew')
    def update_gui():
        for i in range(len(buttons)):
            if i // max_columns >= max_rows:
                break
            row = i // max_columns
            col = i % max_columns
            button_value = buttons[i][0] if isinstance(buttons[i], list) else buttons[i]
            button_text = str(button_value)
            button_color = buttons[i][2] if len(buttons[i]) > 2 else "#000000" 
            # def on_enter(event):
            #     event.widget.config(bg="white")  
            #     event.widget.config(cursor="hand2")
            # def on_leave(event):
            #     event.widget['background'] = event.widget.default_bg 
            btn = tk.Button(window, text=button_text, width=15, height=5, command=lambda text=button_text: button_click(text), bg=button_color, font=font.Font(size=13, weight='bold'), relief='raised', borderwidth=5)
            btn.default_bg = button_color
            btn.bind("<Enter>", on_enter)
            btn.bind("<Leave>", on_leave)
            btn.grid(row= row + 1, column= col + 1) 
    
    add_button_left = tk.Button(left_column_frame, text="ADD ", command=add_button_popup, relief='raised', borderwidth=5,bg='#ffd166',activebackground='#ffd166', fg='#000',width=10)
    add_button_left.grid(row=0, column=0, padx=10, pady=10)
    remove_button_left = tk.Button(left_column_frame, text="REMOVE", command=remove_selected_button, relief='raised', borderwidth=5,bg='#ffd166',activebackground='#ffd166', fg='#000',width=10)
    remove_button_left.grid(row=1, column=0, padx=10, pady=10)
    #refresh_button_left = tk.Button(left_column_frame, text="REFRESH", command=refesh)
    #refresh_button_left.grid(row=2, column=0, padx=10, pady=10)
    config = tk.Button(left_column_frame, text="CONFIGURE", command=configure, relief='raised', borderwidth=5,bg='#ffd166',activebackground='#ffd166', fg='#000',width=10)
    config.grid(row=3, column=0, padx=10, pady=10)
    edit_button = tk.Button(left_column_frame, text="EDIT BUTTON",command = edit, relief='raised', borderwidth=5,bg='#ffd166',activebackground='#ffd166', fg='#000',width=10)
    edit_button.grid(row=4, column=0, padx=10, pady=10)
    Row = tk.Label(left_column_frame, text=f"Max rows: {max_rows}", bg = '#0d6759',fg='white')
    Row.grid(row=5, column=0, padx=10, pady=10)
    Column = tk.Label(left_column_frame, text=f"Max columns: {max_columns}", bg = '#0d6759',fg='white')
    Column.grid(row=6, column=0, padx=10, pady=10)
    Max_tool = tk.Label(left_column_frame, text=f"Can currently accommodate\n up to {max_columns*max_rows} tools", bg = '#0d6759',fg='white')
    Max_tool.grid(row=7, column=0, padx=10, pady=10)
    add_button_left.bind("<Enter>", on_enter_left)
    add_button_left.bind("<Leave>", on_leave_left)
    remove_button_left.bind("<Enter>", on_enter_left)
    remove_button_left.bind("<Leave>", on_leave_left)
    edit_button.bind("<Enter>", on_enter_left)
    edit_button.bind("<Leave>", on_leave_left)
    add_button_left.bind("<Enter>", on_enter_left)
    add_button_left.bind("<Leave>", on_leave_left)
    config.bind("<Enter>", on_enter_left)
    config.bind("<Leave>", on_leave_left)
    update_gui()   
    window.mainloop()