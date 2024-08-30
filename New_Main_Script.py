'''
# * Script Description
* This script is aimed at helping users to quickly update file properties for Solidworks part files and assemblies based on a generated Solidworks BOM without needing to manually open each file and edit file properties.

* The script is based on a fork of an open source library pySldWrap: https://github.com/ThomasNeve/pySldWrap. The library uses the pywin32 project (win32com python library) to communicate with the COM interface of the Solidworks API. Python functions are then wrapped around a subset of the Solidworks API.

* The forked library includes several new functions:
    1. get_custom_file_properties
        * Retrieves all custom file properties of a single file (be it a part or assembly) and returns list with some arguments related to the property. 
    2. export_custom_file_properties
        * Exports filenames and their respective custom properties to an excel file.
        * To be used when there are only part files and no assemblies as this generates a BOM of sorts.
    3. set_custom_file_properties 
        * Writes the values in the excel file to the SOLIDWORKS part files and assemblies.

# * Script Usage
* Generate a Solidworks BOM file.
* Run the script/exe file.
* A folder selection box for the part/assembly file directory will appear.
* Select the directory.
* A file selection box for the excel BOM location will appear.
* Select the BOM file. This will open the excel BOM.
* Make changes to the BOM file. DO NOT CLOSE IT.
* Once changes are complete, close the BOM file. The script will detect the closure and begin writing changes from the excel file to the Solidworks part and assemblies.


# * Main Use Cases
Use Case 1
* Extract file properties from all parts, put them into an excel file
* Modify the excel file
* Extract data from modified excel file, use this to modify data in part files.
* Generate the modified BOM (done within Solidworks and not using python)

Use Case 2
* Changing property values in multiple part files (currently supports same values only)
* Eg. 10 part files need their project names changed

'''

#Imports
import pySldWrap.sw_tools as sw_tools
import importlib
import os
import time 
from pathlib import Path

import tkinter as tk
from tkinter import filedialog
import os
import pandas as pd
import win32com.client
import win32gui

from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

import datetime
import numpy as np



def reload_and_connect():
    '''
    Reload file and connect to Solidworks
    Remember to change the Solidworks version to the current one being used
    '''
    importlib.reload(sw_tools)
    sw_tools.connect_sw("2024")

def retrieve_file_properties_single_part(part_path):
    '''
    Retrieve custom file properties for a single part
    '''
    reload_and_connect()
    model = sw_tools.open_part(part_path)  # open the model, link is returned
    custom_properties = sw_tools.get_custom_file_properties(part_path)
    a = sw_tools.export_custom_file_properties(custom_properties)

def retrieve_file_properties_multiple_parts(directory):
    reload_and_connect()

def prompt_user_for_path(path_name,path_type):
    '''
    Gets filename for some processing, returns file_path for additional processing
    '''
    root = tk.Tk()
    root.withdraw()

    script_dir = os.path.dirname(os.getcwd())

    #get file path or directory
    if path_type == 'File':
        #get file path
        file_path = filedialog.askopenfilename(
            title=f"{path_name}", 
            initialdir=script_dir 
        )
    if path_type == 'Directory':
        file_path = filedialog.askdirectory(
            title=f"{path_name}", 
            initialdir=script_dir 
        )

    # Bring the file dialog window to the front
    root.deiconify()  # Make the root window visible (it's already hidden)
    root.focus_force()  # Force the root window to get focus
    root.after(1, lambda: root.withdraw())  # Hide the root window after a short delay

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



'''
Excel Functions

'''
class ExcelCloseHandler(FileSystemEventHandler):
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
    """
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True  # Make Excel visible if needed
    workbook = excel.Workbooks.Open(bom_file)
    print('File opened sucessfully.')

    return excel, workbook

def wait_for_workbook_close(excel, workbook):
    while True:
        try:
            if workbook not in excel.Workbooks:
                break
        except:
            time.sleep(0.5)  # Wait a bit before retrying
        time.sleep(0.1)

def wait_for_file_to_close(bom_file):
    event_handler = ExcelCloseHandler(bom_file)
    observer = Observer()
    observer.schedule(event_handler, path=bom_file, recursive=False)
    observer.start()

    try:
        while not event_handler.file_closed:
            time.sleep(0.1)
    finally:
        observer.stop()
        observer.join()

def close_excel_file(excel):
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
    # You'll likely need to wait for the Excel file to close before reading it again
    # Implement a way to detect the file closure (e.g., using file monitoring).
    # For simplicity, we'll assume the file is closed after a short delay.
    time.sleep(1)  # Adjust the delay as needed
    final_bom_df = pd.read_excel(bom_file)
    return final_bom_df


def compare_dataframes(initial_df, final_df):
    # Fill NaN values with a consistent placeholder (e.g., an empty string)
    initial_df = initial_df.fillna('')
    final_df = final_df.fillna('')

    # Apply strip only to object columns (string columns)
    str_cols_initial = initial_df.select_dtypes(include=['object']).columns
    str_cols_final = final_df.select_dtypes(include=['object']).columns

    initial_df[str_cols_initial] = initial_df[str_cols_initial].applymap(lambda x: x.strip() if isinstance(x, str) else x)
    final_df[str_cols_final] = final_df[str_cols_final].applymap(lambda x: x.strip() if isinstance(x, str) else x)

    # Compare the two DataFrames
    differences = initial_df.ne(final_df)
    
    # Find the rows where any column is different
    changed_rows = differences.any(axis=1)
    
    # Get the subset of rows from final_df where changes occurred
    changed_data = final_df.loc[changed_rows]
    
    return changed_data


def modify_sw_file_properties(directory, df_of_modified_files, sld_app):
    '''
    Reads from a dataframe containing names of partfiles/assemblies whose file properties have been modified
    Iterates through these files based on a user provided directory
    

    Reads names and values of file properties in excel file
    Checks these against names and values in part/assembly files
    Overwrites them
    '''
    # *create a for loop that does this for every name in the title column:
    # *read the title and associate it with a name of a part file or assembly; if title contains assembly then its an assembly; if not its not
    # *open the specified file or assembly
    # *get the file properties in that part file/assembly
    # *replace the file properties in the part file/assembly with those in the excel file

    # Generate a new log file name
    logfile_count = 1
    while os.path.exists(f"Solidworks_Log_File{logfile_count}.txt"):
        logfile_count += 1
    logfile_name = f"Solidworks_Log_File{logfile_count}.txt"


    

    names_of_modified_files = list(df_of_modified_files.loc[:,"Title"])
    new_names_of_modified_files = [string.replace('\n', '') for string in names_of_modified_files]

    # #accesses the first excel row to retrieve names of rows
    excel_df_2 = df_of_modified_files.set_index(df_of_modified_files.columns[1])

    skip_columns = {'S/N','Enterprise Part No.', 'SurfaceFinish', 'Project', 'Title', 'V_Name', 'Revision', 'Creation Date', 'DrawnDate', 'Material', 'CheckedDate', 'EngAppDate', 'MfgAppDate', 'QAAppDate', 'Remarks', 'DrawnBy', 'CheckedBy', 'EngApproval', 'MfgApproval', 'QAApproval'}

    new_excel_df = excel_df_2.drop(columns=skip_columns, errors='ignore')

    # print(new_excel_df)

    # new_excel_df = df_of_modified_files.drop(columns=skip_columns, errors='ignore')

    # print('excel df:',new_excel_df)


    #get seperate property_value_list for each modified file
    for i in range(df_of_modified_files.shape[0]):
        #checking if its a part file or assembly file

        if "Assembly" in str(new_names_of_modified_files[i]) or "Assem" in str(new_names_of_modified_files[i]):
            filename = new_names_of_modified_files[i] + '.SLDASM'
        else:
            filename = new_names_of_modified_files[i] + '.SLDPRT'

        # # Assuming 'new_excel_df' is your DataFrame
        # new_excel_df['Index_Column'] = new_excel_df.index

        # # Reorder the columns so that 'Index_Column' is the first column
        # new_excel_df = new_excel_df[['Index_Column'] + [col for col in new_excel_df.columns if col != 'Index_Column']]

        # Now you can extract the property_value_list as before

        column_names = new_excel_df.columns
        property_value_list = list(new_excel_df.loc[names_of_modified_files[i], :])

        property_value_dict = dict(zip(column_names, property_value_list))

        # print(property_value_dict)


        # open part/assembly and get custom file properties
        try:
            part_path = os.path.join(directory, filename)
            print(part_path)
            if "Assembly" in filename or "Assem" in filename:
                model = sw_tools.open_assembly(part_path)
            else:
                model = sw_tools.open_part(part_path)
            sw_tools.set_file_properties(model,property_value_dict,sld_app,part_path)
            sw_tools.close(part_path.split('\\')[-1])
            model = None
        
        # Create log file for error handling
        except Exception as e:
            error_timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            print(f'Error processing {filename}: {e}. See {logfile_name} for details.')
            with open(logfile_name, 'a') as f:
                f.write(f"{error_timestamp} - {part_path} - Error: {e}\n")  

def main():
    # Assume BOM is generated already
    # Run script

    # *User is prompted to select part/assembly directory
    # *User selects directory
    # *User gets prompt to select BOM file
    # *User selects BOM file
    # *State and values of BOM file are stored in a dataframe
    # BOM file opens
    # User makes changes to BOM file
    # User saves changes or closes BOM excel file
    # Script detects closure of excel file and extracts final BOM file state. #! ask GPT if this is possible
    # Script compares changes between initial state and final state of dataframe and updates any rows that have been changed.
    # Script begins opening solidworks files and updating file properties.

    importlib.reload(sw_tools)
    sld_app = sw_tools.connect_sw("2024")  # open connection and pass Solidworks version

    working_directory = prompt_user_for_path('Select your assembly/part file directory','Directory')
    bom_file = prompt_user_for_path('Select the BOM file','File')
    initial_bom_df = store_state(bom_file)

    # BOM file opens
    # excel = open_excel_file(bom_file)
    excel, workbook = open_excel_file(bom_file)
    print('excel open')

    # User makes changes to BOM file
    # User saves changes or closes BOM excel file
    # Wait for user to manually close the workbook
    wait_for_workbook_close(excel, workbook)

    # Close Excel application

    close_excel_file(excel)
    # workbook.Close(SaveChanges=False)
    # excel.Quit()

    final_bom_df = extract_final_state(bom_file)
    # changes = compare_dataframes(initial_bom_df, final_bom_df)

    # print('initial_bom:', initial_bom_df.iloc[:, 4:])

    # print('final bom:', final_bom_df.iloc[:, 4:])

    # changes = final_bom_df[~final_bom_df.eq(initial_bom_df).all(axis=1)]


    # # Create a Boolean mask where values differ between initial_df and final_df
    # mask = (initial_bom_df != final_bom_df).any(axis=1)

    # # Use the mask to filter initial_df for rows with changes
    # changes = final_bom_df[mask]

    # print('changes:',changes)


    # Compare DataFrames to find the changes
    changes = compare_dataframes(initial_bom_df, final_bom_df)

    modify_sw_file_properties(working_directory,changes,sld_app)


#     # Proceed with the rest of your logic
#     print('Changes detected:', changes)
#     if not changes.empty:
#         modify_sw_file_properties(working_directory, changes)
    

#     # shortened_working_directory = './' + working_directory.split('/')[-1]

#     # modify_sw_file_properties(shortened_working_directory,changes)



if __name__ == "__main__":
    main()
