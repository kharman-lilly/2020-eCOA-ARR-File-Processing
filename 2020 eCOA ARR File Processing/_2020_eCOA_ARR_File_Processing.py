# Author:  Ken Harman
# Version: 1  - 9/21/2020 - Original
# Check the response files from reviewers in connection with the 2020 eCOA annual account review.

# This script loops thru all the .xlsx file in a directory.  For each file it makes some validation checks and logs any discrepancies in
# an output file.  Errors are logged in a text file.  Both are located in the same directory as where the workbooks are found.
# There are two different vendors and each has a different report formatting.  The report format is detected and output tabs in the Excel workbook
# are unique to each vendor.  There is one tab (per vendor) for accounts to be revoked, one for account role changes, and one for comments from the reviewers.

import os
import pandas as pd
import pathlib
import openpyxl
import datetime
import win32com.client
from win32com.client import Dispatch
import time
import xlsxwriter

path = 'C:\\Users\\ga65937\\OneDrive - Eli Lilly and Company\\Documents\\eCOA\\ARR\\2020\\Complete Received Files'
#path = 'C:\\Users\\ga65937\\Desktop\\New folder'
OutputFileName = "2020_ARR_Summary.xlsx"
ErrorLog = "2020_ErrorLog.txt"
HeadersSignant = ['Study', 'Site Name', 'Site Number', 'Last Name', 'First Name', 'Email', 'Domain', 'Role', 'Reviewer Name', 'Retain/Revoke', 'New Role', 'Comments']
HeadersERT = ['Study Name', 'Email Address', 'Domain', 'First Name', 'Last Name', 'Site', 'Role', 'Reviewer Name', 'Retain/Revoke', 'New Role', 'Comments']
HeadersEPX = ['Study Name', 'System', 'Email Address', 'Domain', 'First Name', 'Last Name', 'Site', 'Role', 'Reviewer Name', 'Retain/Revoke', 'New Role', 'Comments']
dfRevokeOutSignant = pd.DataFrame()
dfNewRoleOutSignant = pd.DataFrame()
dfCommentOutSignant = pd.DataFrame()
dfRevokeOutERT = pd.DataFrame()
dfNewRoleOutERT = pd.DataFrame()
dfCommentOutERT = pd.DataFrame()
dfRevokeOutEPX = pd.DataFrame()
dfNewRoleOutEPX = pd.DataFrame()
dfCommentOutEPX = pd.DataFrame()

os.chdir(path) # Change to the working directory

# Remove previous ErrorLog and output file
if pathlib.Path(ErrorLog).exists():
    os.remove(ErrorLog)
if pathlib.Path(OutputFileName).exists():
    os.remove(OutputFileName)

# Process files in the directory
for filename in os.listdir(path):
    if filename.endswith(".xlsx") and filename != OutputFileName: #Only interested in Excel workbooks and skipping the output file.
        # Log which file we are working with
        with open(ErrorLog, "a") as text_file:
            text_file.write(f"{filename} {str(datetime.datetime.now())}\n")
        print(filename)

        dfXL = pd.read_excel(filename, dtype='string') # Force datatype as string.
        dfXL = dfXL.select_dtypes(['string']).apply(lambda x: x.str.strip()) # Remove leading/trailing spaces in the dataframe.

        # Detect the vendor by looking at the expected column headers.
        if list(dfXL.columns) == HeadersSignant:
            vendor = 'Signant'
        elif list(dfXL.columns) == HeadersERT:
            vendor = 'ERT'
        elif list(dfXL.columns) == HeadersEPX:
            vendor = 'EPX'
        else:
            Vendor = ''
            print(f"*** Fatal Error {filename} has unexpected column names. Skipping file.")
            with open (ErrorLog, "a") as text_file:
                text_file.write(f"*** FATAL ERROR *** {filename} has unexpected column names. Skipping file.\n")
            continue
        
        # Make sure every row in the Reviewer Name column has a value.  Log count of null values.
        null_count = pd.isnull(dfXL["Reviewer Name"]).sum()
        if null_count > 0:
            with open(ErrorLog, "a") as text_file:
                text_file.write(f"*** {filename} has {str(null_count)} missing Reviewers.\n")

        # Make sure every row in the Retian/REvoke column has a value. Log count of null values.
        null_count = pd.isnull(dfXL["Retain/Revoke"]).sum()
        if null_count > 0:
            with open (ErrorLog, "a") as text_file:
                text_file.write(f"*** {filename} has {str(null_count)} missing Retain/Revoke entires.\n")

        # Collect rows that have a Revoke in the Retain/Revoke column
        dfRevokeRows = dfXL[dfXL["Retain/Revoke"] == "Revoke"]
        if vendor == "Signant":
            dfRevokeOutSignant = dfRevokeOutSignant.append(dfRevokeRows)
        elif vendor == "EPX":
            dfRevokeOutEPX = dfRevokeOutEPX.append(dfRevokeRows)
        else:
            dfRevokeOutERT = dfRevokeOutERT.append(dfRevokeRows)

        # Collect rows that have a New Role
        dfNewRoles = dfXL[dfXL["New Role"].notnull()]
        if vendor == "Signant":
            dfNewRoleOutSignant = dfNewRoleOutSignant.append(dfNewRoles) 
        elif vendor == "EPX":
            dfNewRoleOutEPX = dfNewRoleOutEPX.append(dfNewRoles)
        else:
            dfNewRoleOutERT = dfNewRoleOutERT.append(dfNewRoles)

        # Collect rows that have a Commnet
        dfComment = dfXL[dfXL["Comments"].notnull()]
        if vendor == "Signant":
            dfCommentOutSignant = dfCommentOutSignant.append(dfComment)
        elif vendor == "EPX":
            dfCommentOutEPX = dfCommentOutEPX.append(dfComment)
        else:
            dfCommentOutERT = dfCommentOutERT.append(dfComment)

    continue

# Drop colums that aren't needed for output
dfRevokeOutSignant = dfRevokeOutSignant.drop(['Domain'], axis=1)
dfNewRoleOutSignant = dfNewRoleOutSignant.drop(['Domain'], axis=1)
dfCommentOutSignant = dfCommentOutSignant.drop(['Domain'], axis=1)
dfRevokeOutERT = dfRevokeOutERT.drop(['Domain'], axis=1)
dfNewRoleOutERT = dfNewRoleOutERT.drop(['Domain'], axis=1)
dfCommentOutERT = dfCommentOutERT.drop(['Domain'], axis=1)
dfRevokeOutEPX = dfRevokeOutEPX.drop(['Domain'], axis=1)
dfNewRoleOutEPX = dfNewRoleOutEPX.drop(['Domain'], axis=1)
dfCommentOutEPX = dfCommentOutEPX.drop(['Domain'], axis=1)

# Write output dataframes to output file (Excel)
with open (ErrorLog, "a") as text_file:
    text_file.write(f"Output File Begin: {str(datetime.datetime.now())}\n")
with pd.ExcelWriter(OutputFileName, engine='xlsxwriter') as writer:
    dfRevokeOutSignant.to_excel(writer, sheet_name='Revoke-Signant', index=False)
    dfNewRoleOutSignant.to_excel(writer, sheet_name='New Role-Signant', index=False)
    dfCommentOutSignant.to_excel(writer, sheet_name='Comments-Signant', index=False)
    dfRevokeOutERT.to_excel(writer, sheet_name='Revoke-ERT', index=False)
    dfNewRoleOutERT.to_excel(writer, sheet_name='New Role-ERT', index=False)
    dfCommentOutERT.to_excel(writer, sheet_name='Comments-ERT', index=False)
    dfRevokeOutEPX.to_excel(writer, sheet_name='Revoke-EPX', index=False)
    dfNewRoleOutEPX.to_excel(writer, sheet_name='New Role-EPX', index=False)
    dfCommentOutEPX.to_excel(writer, sheet_name='Comments-EPX', index=False)
with open (ErrorLog, "a") as text_file:
    text_file.write(f"Output File End: {str(datetime.datetime.now())}\n")

# Open Excel and use AutoFit to change column widths (ExcelWriter doesn't have this ability)
time.sleep(30) # Short delay to allow Windows time to clean up before repoening the file.
excel = Dispatch('Excel.Application')
wb = excel.Workbooks.Open(path + "\\" + OutputFileName) # This wants the full path to the file.
for sheet in range(1,10):
    excel.Worksheets(sheet).Activate()
    excel.ActiveSheet.Columns.AutoFit()
    excel.ActiveWindow.SplitColumn = 0
    excel.ActiveWindow.SplitRow = 1
    excel.ActiveWindow.FreezePanes = True
    continue
excel.Worksheets(1).Activate() # Put focus back on the first worksheet
wb.Save()
wb.Close()
excel.Application.quit()
