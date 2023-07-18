# Python App that automates updating multiple Excel files

Excel Files Updater

Overview:
This application allows updating multiple Excel files in a folder by searching for a date and updating a corresponding percentage value. It provides a GUI to configure the date, percentage, and folder path to process.

Usage:

1. Install Required Packages

This application requires the following packages:
- tkinter 
- openpyxl
- os
- datetime

Make sure these are installed by running:

pip install tkinter openpyxl os-sys datetime

2. Provide Inputs

- Enter the date to search for in m/d/yyyy format 
- Enter the percentage to update matching cells to
- Click "Choose Directory" and select the folder with Excel files

3. Click "Update Files"

This will recursively search for .xlsx files in the selected folder and sub-folders. It will search for the entered date in Column A and update the corresponding value in Column B with the entered percentage.

The text box will display the status and any errors. Green messages indicate successes, red indicates failures.

4. Executing

To run the application, execute the excel_automationGUI.py script:

python excel_automationGUI.py

This will launch the GUI. Follow steps 2 and 3 above to run the automated update on a folder of Excel files.
