# HPLCSystemSuitability
VBA code for macro in System Suitability Spreadsheet for HPLC - Report

Code is for automatically extracting information from TXO datafiles produced by HPLC software and inputting it into the above excel report.

This code can be copy-pasted into a VBA editor in excel to create a macro, and then the macro assigned to a shape to act as a button.

![image](https://github.com/k-schisler/HPLCSystemSuitability/assets/162852955/0e6b1521-5f03-4626-8fc0-e4cd37491589)

*Note: this code DOES NOT FORMAT an excel file to match the report - the formatting must be done prior to use.*

Code needs to be adjusted per instrument, since file names are specific to the instrument
> This is a space for improvement, find a way to make these sections of code universal.

## Main functions of the macro include:
1. Clearing all data from relevant cells.
2. Pop-up message boxes to input relevant information like serial numbers, lot numbers, and standard solution information.
3. Setting a custom file path to open the file explorer to the loction the data files are in based on HPLC number, test method number, and date.
4. Allowing multiselect of the relevant data files and extracting the relevant data to input into the correct cells in the report.
5. Opening the print preview.



