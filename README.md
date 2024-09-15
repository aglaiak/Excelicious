# Excelicious

![Logo](Logo-1.PNG)

Automating Excel Processes with a Comprehensive GUI for a better barcode future.

In many professional settings, particularly in academic and laboratory environments, Excel serves as the primary tool for documentation, tracking, and data management. However, manual input and management of data in Excel, especially for tasks like barcode generation and tracking, are susceptible to errors. Duplicate entries, inconsistent file naming, and manual data verification can lead to inaccuracies, duplicates and data loss.

This GUI-based application aims to provide a robust and user-friendly solution for individuals who rely on Excel to manage and track specific values—such as barcodes—across multiple spreadsheets. It automates common processes, including identifying duplicate values, verifying data entry consistency, and managing Excel file operations, all through an intuitive graphical interface.

# GUI

The basic GUI is shown below:

![GUI](image.png)

The user can select the function they are interested in. 

### Check for Duplicate Barcodes

In order for the application to run correctly, there are a few steps that should be followed:
- The spreadsheet's tab should be named as "Reference". Otherwise, the source code in line 47 should 
be changed accordingly.
- The barcode sheet should have the columns "barcode", "date", "name" and "status". An excel template is
provided. 
- "Status" is an excel Formula < IF(NOT(ISBLANK('SelectedCellGoesHere')); "edited"; "not edited") >
and results to "edited" if a barcode value is taken or "not edited" when a barcode value is not taken.

Upon clicking on the button, the user is prompted to select a folder with all the excel files. 
Each excel file's 100 values will be read from bottom to top and a concatenated version of barcode, date,
and name will be compared among the values of the other files. If there is a hit, the barcode is considered as duplicate

### Check for Naming Inconsistencies

Naming convention is very important. This function checks the date (which is usually used as prefix) and evaluates the number of digits. If the number of digits surpasses the value of 6 or is lower than 6, then the 
format is deemed inconsistent.

Further updates will be available upon request.

### Terminate all Excel Process

This function, serves as a tool that kills all active Excel instances. 



