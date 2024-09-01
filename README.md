# Script Overview
* This script is aimed at helping users to quickly update file properties for Solidworks part files and assemblies based on a generated Solidworks BOM without needing to manually open each file and edit file properties.

* The script is based on a fork of an open source library pySldWrap: https://github.com/ThomasNeve/pySldWrap. The library uses the pywin32 project (win32com python library) to communicate with the COM interface of the Solidworks API. Python functions are then wrapped around a subset of the Solidworks API.


<!-- TODO: update this part -->
* The forked library includes several new functions:
    1. get_custom_file_properties
        * Retrieves all custom file properties of a single file (be it a part or assembly) and returns list with some arguments related to the property. 
    2. export_custom_file_properties
        * Exports filenames and their respective custom properties to an excel file.
        * To be used when there are only part files and no assemblies as this generates a BOM of sorts.
    3. set_custom_file_properties 
        * Writes the values in the excel file to the SOLIDWORKS part files and assemblies.

## Installation
* Navigate to a directory of your choice.
* Create a virtual environment and activate it.
* Clone the repo to the directory.
* Run the following command to install the necessary files:
    ```sh
    pip install -r requirements.txt
    ```
<!-- TODO: generate a requirements.txt file -->

## Script Operation
* Before running the script, SolidWorks should be opened. This should be the default blank screen at start up.
* To run the script, run this command in the script directory:
    ```sh
    python New_Main_script.py
    ```
* Upon running the script, the following will happen:
    1. A directory selection dialog box will be opened to select the part/assembly file directory. If the dialog box is not visible, minimise your application windows until you see it.
    2. After selecting the directory, a file selection dialog box will be opened to select the excel BOM.
    3. After selecting the excel BOM, wait for the script to open the excel BOM.
    4. Now make changes to the BOM file. <span color="red">DO NOT CLOSE IT</span>
    

    <p style="color:red">DO NOT CLOSE IT</p>

     as this will start the next part of the script.
    5. Once changes are complete, close the BOM file. The script will detect the closure and begin writing changes from the excel file to the Solidworks part and assemblies.

**Stopping the script**
* The script can only be stopped by:
    * Closing the directory selection or file selection dialog boxes before choosing any file or directory
    * abc
* Currently, there is no way to stop the script once step 5 has started.


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