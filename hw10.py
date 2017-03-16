#!/Library/Frameworks/Python.framework/Versions/3.6/bin/python3

"""Usage - 1. ensure that the excel file 'BPM_Automation_Roadmap2 _1612_16Nov.xlsx'
              is in the current directory.
           2. Ensure that the file MasterScriptTests.txt is in the current directory
           3. Ensure hw10.py is in the current directory
           4. Ensure file_operations.py is in the current directory
           5. Run on command line hw10.py 'BPM_Automation_Roadmap2 _1612_16Nov.xlsx' 'MasterScriptTests.txt'
           6. The ouput file will be 'BPM_Automation_Roadmap2.xlsx' in the current directory
           7. The script populates column R of the Async sheet of the excel sheet with x's if the test names in
            'MasterScriptTests.txt' are in 'BPM_Automation_Roadmap2 _1612_16Nov.xlsx'
           8. The 'InMasterScriptFolderNotInRoadMa' sheet is populated with tests that are in the 'MasterScriptTests.txt'
              but not in the Async sheet
 """

# MasterScript folder is the repository for all automated test cases

# import openpyxl.reader.excel as py
from file_operations import FileOperations
import sys

roadmap_ops = FileOperations()
# Road map file name
excel_file = sys.argv[1] #"BPM_Automation_Roadmap2 _1612_16Nov.xlsx"
text_file = sys.argv[2] #'./MasterScriptTests.txt'

# Open the road map, (testers will update this file manually before submitting their test case to the MasterScript
#  folder
list_wb_mfolder_list = roadmap_ops.open_the_files(excel_file, text_file)
# wb = py.load_workbook(roadmap_filename)

# Open file containing all the automated test cases from the MasterScript folder
# master_script_file = open("./MasterScriptTests.txt", 'r')

# Put the MasterScript files into a list then close the file
mfolder_list = list_wb_mfolder_list[1]
# master_script_file.close()

# the workbook
wb = list_wb_mfolder_list[0]

# Clear out the SVNMasterScriptExport sheet, if it exists, otherwise create it... this is where the test cases from
#  the MasterScript folder will go
cleared_sheet_SVN = roadmap_ops.get_clean_excel_sheet('SVNMasterScriptExport', wb)
cleared_sheet_InM = roadmap_ops.get_clean_excel_sheet('InMasterScriptFolderNotInRoadMa', wb)

# if 'SVNMasterScriptExport' in wb.sheetnames:
#   remove_sheet = wb.get_sheet_by_name('SVNMasterScriptExport')
#  wb.remove(remove_sheet)
# wb.create_sheet('SVNMasterScriptExport')
# else:
# wb.create_sheet('SVNMasterScriptExport')

# Clear out the InMasterScriptFolderNotInRoadMa sheet, if it exists, otherwise create it...
#  this is where the test cases from the MasterScript folder that are not in the roadmap will go
# if 'InMasterScriptFolderNotInRoadMa' in wb.sheetnames:
#   remove_sheet = wb.get_sheet_by_name('InMasterScriptFolderNotInRoadMa')
#  wb.remove(remove_sheet)
# wb.create_sheet('InMasterScriptFolderNotInRoadMa')
# else:
#   wb.create_sheet('InMasterScriptFolderNotInRoadMa')

# create a reference to the 'InMasterScriptFolderNotInRoadMa' sheet
not_in_rm_ws = wb.get_sheet_by_name('InMasterScriptFolderNotInRoadMa')

# Copy the MasterScript tests to the SVNMasterScriptExport, column E in the road map spreadsheet.
# Get the SVN sheet
svn_ws = wb.get_sheet_by_name('SVNMasterScriptExport')
# Must create the cells, 1000 rows in column E before assigning them values
for i in range(1, 10000):
    svn_ws.cell(row=i, column=5)

# Do the copy from MasterScript to the roadmap sheet SVNMasterScriptExport, strip the \n's
i = 1
for x in mfolder_list:
    x = x.strip()
    svn_ws.cell(row=i, column=5, value=x)
    i += 1

# Prepare to compare what's on the SVNMasterScriptExport sheet with what is on the Async sheet, column E
# Get the Async sheet
async_ws = wb.get_sheet_by_name('Async')
# Clear column R starting with row 15, don't clear previous rows
for i in range(15, 10000):
    async_ws.cell(row=i, column=18, value="")

# Compare each MasterScript test with the road map test case name in sheet Async,(column E, starting with row 15)

# If the MasterScript test is on the road map, put an x in column R to signify that the test is in the MasterScript
# folder
for i in range(1,10000):
    test_found = 0
    if svn_ws.cell(row=i, column=5).value is None:
        break
    for j in range(15,10000):
        if async_ws.cell(row=j, column=5).value is None:
            break
        if (svn_ws.cell(row=i, column=5).value).strip() == (async_ws.cell(row=j, column=5).value).strip():
            async_ws.cell(row=j, column=18, value="X")
            test_found = 1
    if test_found == 0:
        # the append method takes a list
        list1 = [svn_ws.cell(row=i, column=5).value]
        not_in_rm_ws.append(list1)

# save the roadmap to a different file
wb.save("BPM_Automation_Roadmap2.xlsx")

# todo list
# Loop through the Sync, Async, "Manual Intervention", "User Interaction(GUI)", and "Out of Scoped" tabs row by row
# Keep on iterating through the worksheets, if there are an other matches for this particular test, update column R
# There will be duplicate test names, (different versions though...versions won't be accounted for in this script) If
#  the MasterScript test is not in the road map, update a worksheet with the test name to signify that the test is
# not in the road map After the script is run, there will be two new tabs, one for the MasterScript tests,
# and one for the MasterScript tests that are not in the road map Use the pandas module to find how many test cases
# have been automated and submitted to the MasterScript repo
