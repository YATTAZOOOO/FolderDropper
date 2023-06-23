import shutil
import pandas as pd
import os
from art import *
import time

time.sleep(0.3)
print('----------------------------------------------------------------------------------------------------------------')
time.sleep(0.3)

name = text2art("Folder_ Dropper  v.1") # Return ASCII text (default font) and default chr_ignore=True 
print(name)

time.sleep(0.3)
print('----------------------------------------------------------------------------------------------------------------')
time.sleep(0.3)

print('')
print("What do you want?")
print("1. Folder copy manually")
print("2. Folder copy from an Excel file")
user_choice = input('Write a number: ')

def folder_copy():

    source_dir_input = input('Enter a source folder path: ')
    destination_dir_input = input('Enter a destination folder path: ')
    folder_name_list = input('Enter folder names: ').split(', ')
    
    missing_folders = []

    for folder_name in folder_name_list:
        source_dir = os.path.join(source_dir_input, folder_name.strip())
        destination_dir = os.path.join(destination_dir_input, folder_name.strip())

        if not os.path.exists(source_dir):
            missing_folders.append(folder_name.strip())
            continue

        try:
            shutil.copytree(source_dir, destination_dir)
        except FileExistsError:
            print(f"The folder '{destination_dir}' already exists.")
        except Exception as e:
            print(f"An error occurred while copying '{source_dir}': {str(e)}")

    if len(missing_folders) > 0:
        print("Missing folders:")
        for folder_name in missing_folders:
            print(folder_name)
    else:
        print("Done")

def folder_copy_excel():

    path_input = input('Enter an Excel file path: ')
    column_input =  input('Enter a column name: ')

    df = pd.read_excel(path_input, sheet_name="Лист1")
    column_list = df[column_input].tolist()
    
    source_dir_input = input('Enter a source folder path: ')
    destination_dir_input = input('Enter a destination folder path: ')

    missing_files = []

    for folder_name in column_list:
        source_dir = source_dir_input + '\\' + str(folder_name)
        destination_dir = destination_dir_input + '\\' + str(folder_name)

        try:
            shutil.copytree(source_dir, destination_dir)
        except FileNotFoundError:
            missing_files.append(str(folder_name))
            continue

    if len(missing_files) > 0:
        print("Missing files:")
        for file_name in missing_files:
            print(file_name)
    else:
        print("Done")

if user_choice == '1':
    folder_copy()
elif user_choice == '2':
    folder_copy_excel()