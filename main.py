import shutil
import pandas as pd
import os
from art import *
import time

time.sleep(0.3)
print('----------------------------------------------------------------------------------------------------------------')
time.sleep(0.3)

name = text2art("Folder_ Dropper  v.1")  # Generate ASCII art from text
print(name)

time.sleep(0.3)
print('----------------------------------------------------------------------------------------------------------------')
time.sleep(0.3)

print('')
print("What do you want?")
print("1. Folder copy manually")
print("2. Folder copy from an Excel file")
user_choice = input('Write a number: ')

# Function to copy folders manually
def folder_copy():
    source_dir_input = input('Enter a source folder path: ')
    destination_dir_input = input('Enter a destination folder path: ')
    folder_name_list = input('Enter folder names: ').split(', ')
    
    missing_folders = []  # List to store missing folders

    for folder_name in folder_name_list:
        source_dir = os.path.join(source_dir_input, folder_name.strip())  # Build the source directory path
        destination_dir = os.path.join(destination_dir_input, folder_name.strip())  # Build the destination directory path

        if not os.path.exists(source_dir):  # Check if the source folder exists
            missing_folders.append(folder_name.strip())  # Add missing folder to the list
            continue

        try:
            shutil.copytree(source_dir, destination_dir)  # Copy the folder from source to destination
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

# Function to copy folders from an Excel file
def folder_copy_excel():
    path_input = input('Enter an Excel file path: ')
    column_input =  input('Enter a column name: ')

    df = pd.read_excel(path_input, sheet_name="Лист1")  # Read Excel file into a DataFrame
    column_list = df[column_input].tolist()  # Extract the column values as a list
    
    source_dir_input = input('Enter a source folder path: ')
    destination_dir_input = input('Enter a destination folder path: ')

    missing_files = []  # List to store missing files

    for folder_name in column_list:
        source_dir = source_dir_input + '\\' + str(folder_name)  # Build the source directory path
        destination_dir = destination_dir_input + '\\' + str(folder_name)  # Build the destination directory path

        try:
            shutil.copytree(source_dir, destination_dir)  # Copy the folder from source to destination
        except FileNotFoundError:
            missing_files.append(str(folder_name))  # Add missing file to the list
            continue

    if len(missing_files) > 0:
        print("Missing files:")
        for file_name in missing_files:
            print(file_name)
    else:
        print("Done")

if user_choice == '1':
    folder_copy()  # Perform manual folder copy
elif user_choice == '2':
    folder_copy_excel()  # Perform folder copy from an Excel file