import os
import xlrd
from xlwt import Workbook 

# Give the location of the file 
Excel_Path = ("S:/Sahu_Group/B1_Research/B1_Automation-ExcelCellUpdate/A1_Input/aa.xlsx")
Excel_Path_Split = Excel_Path.split('/')
Excel_Name = Excel_Path_Split[len(Excel_Path_Split)-1]
print(Excel_Name)

# Operation-Open
Excel_Open = xlrd.open_workbook(Excel_Path,on_demand=True) 

"""
# Operation-File-Find
# Length
Excel_Sheet_Len = len(Excel_Open.sheet_names()) 
print(Excel_Sheet_Len)

# Operation-Sheet-Single-Read-Whole
Sheet_Structure = Excel_Open.sheet_by_index(0) 
# Operation-Sheet-Single-Read-Cell
Sheet_Cell_Val= Sheet_Structure.cell_value(0,0) 

# Operation-Sheet-Single-Find
# Row-Col
Sheet_Structure_Row = Sheet_Structure.nrows
Sheet_Structure_Col = Sheet_Structure.ncols
print("No. of rows:", Sheet_Structure_Row)               
print("No. of columns:", Sheet_Structure_Col) 
"""

Replacement_KeyPair = [
                        ['Co000000001Space','Fuschia'],
                        ['Oranges','Lemons']
                      ]
# Operation-Sheet-Multi-Read-Whole
# --
WExcel_Path = 'S:/Sahu_Group/B1_Research/B1_Automation-Excel/C1_Output'
#WriteFile_Name = '2.xxx_Information.xlsx'
WriteFile_Name = Excel_Name
New_WBook_Write = Workbook()
# --
for Sheet_Index in range(len(Excel_Open.sheet_names())) :
    # Operation-Sheet-Single-Read-Whole
    Sheet_Structure = Excel_Open.sheet_by_index(Sheet_Index) 
    # Operation-Sheet-Single-Read-Cell
    Sheet_Cell_Val= Sheet_Structure.cell_value(0,0) 
    # Operation-Sheet-Single-Find
    # Row-Col
    Sheet_Structure_Row = Sheet_Structure.nrows
    Sheet_Structure_Col = Sheet_Structure.ncols
    print("No. of rows:", Sheet_Structure_Row)               
    print("No. of columns:", Sheet_Structure_Col) 
    
    # Operation-Sheet-Multi-Replace
    New_WBook_Write_Sheet_Structure = New_WBook_Write.add_sheet('Sheet'+str(Sheet_Index+1),cell_overwrite_ok=False)   
    for Sheet_Row_Index in range(Sheet_Structure.nrows):
        for Sheet_Col_Index in range(Sheet_Structure.ncols):
            # -- To Avoid Over Write First As Normal
            New_Str = str(Sheet_Structure.cell_value(Sheet_Row_Index,Sheet_Col_Index)) 
            for Replacement_KeyPair_Index in range(0,len(Replacement_KeyPair)):
                if str(Sheet_Structure.cell_value(Sheet_Row_Index,Sheet_Col_Index)).find(Replacement_KeyPair[Replacement_KeyPair_Index][0])!=-1 :
                    New_Str = New_Str.replace(Replacement_KeyPair[Replacement_KeyPair_Index][0],Replacement_KeyPair[Replacement_KeyPair_Index][1])             
                    print(New_Str)
            New_WBook_Write_Sheet_Structure.write(Sheet_Row_Index,Sheet_Col_Index,New_Str)
# --
Chk_File = os.path.exists(WExcel_Path+str(chr(47))+WriteFile_Name) 
if Chk_File is True :
    New_WBook_Write.save(WExcel_Path+str(chr(47))+WriteFile_Name) 
    print(WExcel_Path+str(chr(47))+WriteFile_Name)
    print('File Succesfully OverWrited')
else :
    with open(os.path.join(WExcel_Path, WriteFile_Name), 'w') as fp: 
        New_WBook_Write.save(WExcel_Path+str(chr(47))+WriteFile_Name) 
        print(WExcel_Path+str(chr(47))+WriteFile_Name)
        print('File Succesfully NewWrited')
# ---------------------------------------------------------------------------------------------
#-end

