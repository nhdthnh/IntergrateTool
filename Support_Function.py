import webbrowser
import os
import shutil
import tkinter.messagebox as messagebox
import time 
from pywinauto import Application

# Function opening link

def open_link(PATH, progress_bar, message_label):
    button_name = PATH.split('\\')[-2].split('.')[-1]
    progress_bar.start()
    message_label.config(text=f"[{button_name}]: Opening links...")
    with open(PATH, 'r') as file:
        data = file.readlines()
    urls = [line.split('Please follow wiki link:')[-1].strip() for line in data if 'Please follow wiki link:' in line]
    for url in urls:
        print(f'[{button_name}]: Open URL:    {url}')
        webbrowser.open(url)
        time.sleep(3)
    print(f'[{button_name}]: Opened {len(urls)} links successfully!')
    if len(urls) ==1:
        message_label.config(text=f"[{button_name}]: Opened link successfully!")
    else:
        message_label.config(text=f"[{button_name}]: Opened {len(urls)} links successfully!")
    progress_bar.stop()
    time.sleep(1)
    message_label.config(text="")
    
# Function opening folder
def open_folder(folder_path, progress_bar, message_label):
    button_name = folder_path.split('\\')[-1].split('.')[-1]
    progress_bar.start()
    message_label.config(text=f"[{button_name}]: Opening folder...")
    # Check if the folder is empty
    if not os.listdir(folder_path):  
        print(f'[{button_name}]: Folder is empty. Cannot open folder: {folder_path}')
        messagebox.showinfo("Maintenance", "This folder is under maintenance.")
        message_label.config(text=f"[{button_name}]: Cannot open this folder.")
        progress_bar.stop()
        time.sleep(1)
        message_label.config(text="")
    else:
        os.startfile(folder_path)
        print(f'[{button_name}]: Open folder: {folder_path}')
        time.sleep(1)
        message_label.config(text=f"[{button_name}]: Opened folder successfully!")
        progress_bar.stop()
        time.sleep(1)
        message_label.config(text="")
        
# Function opening folder and link    
def open_folder_and_link(folder_path, link_path, progress_bar, message_label):
    open_folder(folder_path, progress_bar, message_label)
    open_link(link_path, progress_bar, message_label)
    
# Function opening file
'''
file_name format: 'file_name.extension'
get computer name: ex: username of computer *\ABCDEF => computer_name = ABCDEF

check if folder named computer_name exists => open exe file in this folder 
if folder_name does not exist => create folder named computer_name and copy exe file to this folder then open exe file in this folder

check update by read update_status.txt file => check every time click on button to open file
if status=='update' => 2 options: update now by click OK or update later by click Cancel
If click OK => copy new file from folder_path and replace file in folder named computer_name (Maybe error by copying file (Bug)=> should read sys.txt to check if you can't open file after copy)
If click Cancel => open current file
'''
def open_file(folder_path, file_name,progress_bar,message_label):
    progress_bar.start()
    button_name = folder_path.split('\\')[-1].split('.')[-1]
    message_label.config(text=f"[{button_name}]: Opening {file_name}...")
    if not os.listdir(folder_path):  
        print(f'[{button_name}]: Folder is empty. Cannot open file: {folder_path}')
        messagebox.showinfo("Maintenance", "This file is under maintenance.")
        message_label.config(text=f"[{button_name}]: Cannot open {file_name}.")
        progress_bar.stop()
        time.sleep(1)
        message_label.config(text="")
    else:
        time.sleep(2)
        message_label.config(text=f'[{button_name}]: Please wait a few minutes for opening {file_name}...')
        computer_name = get_myComputer()
        print('User name:   '+computer_name)
        file_name_not_extension = file_name.split('.')[0]
        new_folder_path = os.path.join(folder_path,'User', computer_name)
        status_file = os.path.join(new_folder_path, 'update_status.txt')
        with open(status_file, 'r') as f:
            status = f.read().strip()
        if status == 'Update':
            result = messagebox.askokcancel("Notification", f"A new version of {file_name_not_extension} is available.\nClick OK to update now or CANCEL to update later.")
            if result:
                os.remove(os.path.join(new_folder_path, file_name))
                shutil.copy(os.path.join(folder_path,file_name), new_folder_path)
                messagebox.showinfo("Notification", f"{file_name_not_extension} has been successfully updated!")
                with open(status_file, 'w') as f:
                    f.write('No Update')
        print(f'[{button_name}]: Open file:   '+ os.path.join(new_folder_path, file_name))
        os.startfile(os.path.join(new_folder_path, file_name))
        while True:
            try:
                app = Application().connect(path=os.path.join(new_folder_path, file_name))
                if app.windows():
                    message_label.config(text=f"[{button_name}]: Opened {file_name} successfully!")
                    break
            except Exception as e:
                print(f"[{button_name}]: Waiting for the {file_name} window to open: {e}")
                time.sleep(1)
        progress_bar.stop()
        time.sleep(1)
        message_label.config(text="")
        
# Function getting computername
def get_myComputer():
    computer_name = os.getlogin().upper()
    return computer_name

# Function compare 2 files by reading binary
def compare_files(file1, file2):
    with open(file1, 'rb') as f1, open(file2, 'rb') as f2:
        file1_data = f1.read()
        file2_data = f2.read()
    return file1_data == file2_data

# Function checking update
'''
If it has a difference between 2 files => write 'Update' to update_status.txt file.
Else: write 'No update'.
Function will be called once for each time open GUI.
'''
def check_update(folder_path, file_name,progress_bar,message_label):
    progress_bar.start()
    message_label.config(text="Checking for application updates...")
    time.sleep(2)
    message_label.config(text=f"Checking for {file_name} updates...")
    if not os.listdir(folder_path):  
        time.sleep(1)
        print(f'Folder is empty. Cannot open file in: {folder_path}')
        message_label.config(text=f"Cannot found {file_name}.")
        progress_bar.stop()
        time.sleep(2)
        message_label.config(text="")
    else:
        computer_name = get_myComputer()
        new_folder_path = os.path.join(folder_path,'User', computer_name)
        status_file = os.path.join(new_folder_path, 'update_status.txt')
        if not os.path.exists(new_folder_path):
            os.makedirs(new_folder_path)
            shutil.copy(os.path.join(folder_path, file_name), new_folder_path)
            with open(status_file, 'w') as f:
                f.write('No Update')
        elif os.path.exists(os.path.join(new_folder_path,file_name)):
            if not compare_files(os.path.join(folder_path,file_name), os.path.join(new_folder_path,file_name)):
                with open(status_file, 'w') as f:
                    f.write('Update')
                print(f"Application {file_name} needs to be updated.")
            else:
                with open(status_file, 'w') as f:
                    f.write('No Update')
                print(f"Application {file_name} does not need to be updated.")
        message_label.config(text="Checked for updates successfully!")
        progress_bar.stop()
        time.sleep(1)
        message_label.config(text="")

def check_version(SW_VER, version_path,sw_path):
    # Read the version from the file
    file_version = ''
    with open(version_path, 'r') as f:
        file_version = f.read().strip()
        f.close()
    # Check if the versions match
    if SW_VER == file_version:
        print("No need to update the application.")
    else:
        # If versions don't match, show a pop-up message
        MsgBox = messagebox.askquestion ('New version available','A new version of the application is available. Do you want to update now?',icon = 'warning')
        print("Update application")
        if MsgBox == 'yes':
            os.startfile(sw_path)  # Open the file path
        else:
            print("You can continue using the application.")