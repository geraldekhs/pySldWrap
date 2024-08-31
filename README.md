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