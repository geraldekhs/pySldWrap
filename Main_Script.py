# Imports
# Standard library imports
import os
import time
import importlib
from pathlib import Path
import datetime

# Related third party imports
import tkinter as tk
from tkinter import filedialog
import pandas as pd
import win32com.client
import win32gui
import numpy as np
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

# Local application/library specific imports
import pySldWrap.sw_tools as sw_tools


def reload_and_connect():
    '''
    Reload library and connect to Solidworks
    Remember to change the Solidworks version to the current one being used
    '''
    importlib.reload(sw_tools)
    sw_tools.connect_sw("2024")

def prompt_user_for_path(path_name,path_type):
    """
    Opens a file dialog box to select a file or directory.

    Args:
        path_type (str): Either 'File' or 'Directory', indicating the type of path to select.
        path_name (str): The title of the dialog box.

    Returns:
        str: The selected file path or directory path.
    """

    # Create a hidden Tkinter root window
    root = tk.Tk()
    root.withdraw()

    # Get the current script's directory.
    script_dir = os.path.dirname(os.getcwd())

    # Open the appropriate file dialog based on path_type.
    if path_type == 'File':
        file_path = filedialog.askopenfilename(
            title=f"{path_name}",  # Set the dialog box title.
            initialdir=script_dir  # Set the initial directory to the script's directory.
        )
    elif path_type == 'Directory':
        file_path = filedialog.askdirectory(
            title=f"{path_name}",
            initialdir=script_dir
        )

    # Briefly show the root window to allow the dialog box to close properly.
    root.deiconify()
    root.focus_force()
    root.after(1, lambda: root.withdraw())

    return file_path

def store_state(bom_file):
    """
    Stores the initial state of the BOM file in a DataFrame.

    Args:
        bom_file (str): Path to the BOM Excel file.

    Returns:
        pandas.DataFrame: A DataFrame representing the initial BOM state.
    """
    bom_df = pd.read_excel(bom_file)
    return bom_df


class ExcelCloseHandler(FileSystemEventHandler):
    """
    This class monitors an Excel file for changes and sets a flag when the file is closed.

    Attributes:
        file_closed (bool): Flag indicating whether the file has been closed.
        path (str): Path to the Excel file being monitored.
    """
    def __init__(self, path):
        self.file_closed = False
        self.path = path

    def on_modified(self, event):
        if event.src_path == self.path:
            self.file_closed = True


def open_excel_file(bom_file):
    """
    Opens the BOM file in Excel.

    Args:
        bom_file (str): Path to the BOM Excel file.

    Returns:
        tuple: contains Excel application object and opened workbook object.
    """
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True  # Make Excel visible if needed
    workbook = excel.Workbooks.Open(bom_file)
    print('File opened sucessfully.')

    return excel, workbook


def wait_for_workbook_close(excel, workbook):
    '''
    Pauses further execution of script until excel file closure is detected 
    '''
    while True:
        try:
            if workbook not in excel.Workbooks:
                break
        except:
            time.sleep(0.5)  # Wait a bit before retrying
        time.sleep(0.1)

def close_excel_file(excel):
    '''
    Close the excel file using python.
    Opening an excel file using python prevents normal excel closure by user from closing excel in the background thus this is needed. 
    '''
    excel.DisplayAlerts = False
    excel.Quit()

def extract_final_state(bom_file):
    """
    Extracts the final state of the BOM file after it's closed.

    Args:
        bom_file (str): Path to the BOM Excel file.

    Returns:
        pandas.DataFrame: A DataFrame representing the final BOM state.
    """
    # Wait for Excel file to fully close before reading its state
    time.sleep(1)
    final_bom_df = pd.read_excel(bom_file)
    return final_bom_df


def compare_dataframes(initial_df, final_df):
    """
    Compares two DataFrames, identifying rows where changes have occurred.

    Args:
        initial_df (pandas.DataFrame): The original DataFrame.
        final_df (pandas.DataFrame): The DataFrame to compare against.

    Returns:
        pandas.DataFrame: A DataFrame containing the rows from `final_df` where changes were detected.
    """
    # Pre-process for comparison:
    # 1. Fill NaN values
    # 2. Strip whitespace from string columns
    initial_df = initial_df.fillna('')
    final_df = final_df.fillna('')
    
    str_cols_initial = initial_df.select_dtypes(include=['object']).columns
    str_cols_final = final_df.select_dtypes(include=['object']).columns

    initial_df[str_cols_initial] = initial_df[str_cols_initial].applymap(lambda x: x.strip() if isinstance(x, str) else x)
    final_df[str_cols_final] = final_df[str_cols_final].applymap(lambda x: x.strip() if isinstance(x, str) else x)

    # Compare the DataFrames and find changed rows
    differences = initial_df.ne(final_df)
    changed_rows = differences.any(axis=1)

    # Extract changed data from the final DataFrame
    changed_data = final_df.loc[changed_rows]

    return changed_data


def modify_sw_file_properties(directory, df_of_modified_files, sld_app):
    '''
    Modifies the file properties of SolidWorks parts and assemblies based on changes in a DataFrame.

    Args:
        directory (str): The directory containing the SolidWorks files.
        df_of_modified_files (pandas.DataFrame): A DataFrame containing names and properties of any files with modified properties.
        sld_app (win32com.client.Dispatch): The SolidWorks application object.

    '''

    # Generates log file names for error logging
    logfile_count = 1
    while os.path.exists(f"Solidworks_Log_File{logfile_count}.txt"):
        logfile_count += 1
    logfile_name = f"Solidworks_Log_File{logfile_count}.txt"

    # Get list of modified files, fix the names which have been modified by Solidworks BOM generation 
    names_of_modified_files = list(df_of_modified_files.loc[:,"Title"])
    new_names_of_modified_files = [string.replace('\n', '') for string in names_of_modified_files]

    # Drops any properties that we don't want to modify (either native to Solidworks or company does not need them in the BOM)
    excel_df_2 = df_of_modified_files.set_index(df_of_modified_files.columns[1])
    skip_columns = {'S/N','Enterprise Part No.', 'SurfaceFinish', 'Project', 'Title', 'V_Name', 'Revision', 'Creation Date', 'DrawnDate', 'Material', 'CheckedDate', 'EngAppDate', 'MfgAppDate', 'QAAppDate', 'Remarks', 'DrawnBy', 'CheckedBy', 'EngApproval', 'MfgApproval', 'QAApproval'}
    new_excel_df = excel_df_2.drop(columns=skip_columns, errors='ignore')

    #get seperate property_value_list for each modified file
    for i in range(df_of_modified_files.shape[0]):
        column_names = new_excel_df.columns
        property_value_list = list(new_excel_df.loc[names_of_modified_files[i], :])
        property_value_dict = dict(zip(column_names, property_value_list))

        #regenerate actual filenames
        if "Assembly" in str(new_names_of_modified_files[i]) or "Assem" in str(new_names_of_modified_files[i]):
            filename = new_names_of_modified_files[i] + '.SLDASM'
        else:
            filename = new_names_of_modified_files[i] + '.SLDPRT'

        # open part/assemblies and get custom file properties
        try:
            part_path = os.path.join(directory, filename)
            print(part_path)
            if "Assembly" in filename or "Assem" in filename:
                model = sw_tools.open_assembly(part_path)
            else:
                model = sw_tools.open_part(part_path)
            sw_tools.set_file_properties(model,property_value_dict,sld_app,part_path)
            sw_tools.close(part_path.split('\\')[-1])
            model = None    # release COM objects to prevent memory leaks
        
        # Create log file for error handling
        except Exception as e:
            error_timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            print(f'Error processing {filename}: {e}. See {logfile_name} for details.')
            with open(logfile_name, 'a') as f:
                f.write(f"{error_timestamp} - {part_path} - Error: {e}\n")  

def main():

    '''
    Script Operation:
        User is prompted to select part/assembly directory
        User selects directory
        User gets prompt to select BOM file
        User selects BOM file
        State and values of BOM file are stored in a dataframe.
        BOM file opens
        User makes changes to BOM file
        User saves changes or closes BOM excel file
        Script detects closure of excel file and extracts final BOM file state.
        Script compares changes between initial state and final state of dataframe and updates any rows that have been changed.
        Script begins opening solidworks files and updating file properties.
    '''

    importlib.reload(sw_tools)
    sld_app = sw_tools.connect_sw("2024")  # open connection and pass Solidworks version

    working_directory = prompt_user_for_path('Select your assembly/part file directory','Directory')
    bom_file = prompt_user_for_path('Select the BOM file','File')
    initial_bom_df = store_state(bom_file)

    excel, workbook = open_excel_file(bom_file) # BOM file is now open
    wait_for_workbook_close(excel, workbook)
    close_excel_file(excel)
    final_bom_df = extract_final_state(bom_file)

    changes = compare_dataframes(initial_bom_df, final_bom_df)

    modify_sw_file_properties(working_directory,changes,sld_app)


if __name__ == "__main__":
    main()
