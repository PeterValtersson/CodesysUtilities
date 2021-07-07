from __future__ import print_function
from scriptengine import *

import clr
clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel

#from System.Runtime.InteropServices import Marshal
#Excel = Marshal.GetActiveObject("Excel.Application")
database = system.ui.open_file_dialog(title = "Choose Excel database to import", default_extension = ".xlsl")

print("Opening {}".format(database))

ex = Excel.ApplicationClass()
ex.Visible = False
workbook = ex.Workbooks.Open(database)
parameters = workbook.Sheets("Parameters")
ids = parameters.Columns("D")
print(ids)







def updateHandler():

    #Select the workbook and worksheet
    #workbook=Excel.ActiveWorkbook
     #worksheet=Excel.ActiveSheet
    for sheet in workbook.sheets:
        print(sheet.Name)
    #print(worksheet.Range["A3"].Text)
    def test():
        #Define Key ranges in the Workbook
        ExcelCell_A = worksheet.Range["A3"].Text
        ExcelCell_B = worksheet.Range["B3"].Text

        #Get The Workbench Parameters
        # P1,P2,... not the "Name"
        lengthParam_A = Parameters.GetParameter(Name="P1")
        lengthParam_B = Parameters.GetParameter(Name="P2")

        #Assign values to the input parameters
        lengthParam_A.Expression = ExcelCell_A

        #Run the project update
        Update()

        #Update the workbook value from the WB parameter, !!dosent work!!
        ExcelCell_B = lengthParam_B


#Run the update
#updateHandler()