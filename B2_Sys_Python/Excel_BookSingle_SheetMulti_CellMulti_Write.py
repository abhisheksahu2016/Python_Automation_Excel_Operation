
# ------------ Write Multi Sheet ----------------
#From Excel
import os
import numpy as np
from numpy import * 
import xlrd 
import xlwt 
from xlwt import Workbook 
# ---------------------------------------------------------------------------------------------
RExcel_Path = ("C:/Users/ABHISHEK/Desktop/try/xxx.xlsx")
Excel_Open = xlrd.open_workbook(RExcel_Path) 

WExcel_Path = 'C:/Users/ABHISHEK/Desktop/try/t'
WriteFile_Name = '2.xxx_Information.xlsx'
wb = Workbook() 
# ---------------------------------------------------------------------------------------------
Content_Structure = Excel_Open.sheet_by_index(1) 
Y_Start = 2
Y_End = 2
while str(Content_Structure.cell(Y_End,1).value)!="Stop" :
    Y_End= Y_End + 1
Y_End = Y_End - 1
X_Start = 2
X_End = 2
while str(Content_Structure.cell(1,X_End).value)!="Stop" :
    X_End= X_End + 1
X_End = X_End - 1
Matrix_Row=(Y_End-Y_Start)+1
Matrix_Col=(X_End-X_Start)+1
# ---------------------------------------------------------------------------------------------
Write_Sheet1 = wb.add_sheet('Sheet1') 
# ---------------------------------------------------------------------------------------------
X_End = X_End + 1 
Y_End = Y_End + 1 # for range increase purppose
Content_Structure_Matrix=[]
temp_str = " "
for x in range(X_Start,X_End):
    for y in range(Y_Start,Y_End):
        temp_str = str(Content_Structure.cell(y,x).value)
        Write_Sheet1.write(y,x, temp_str) 
# ---------------------------------------------------------------------------------------------
Chk_File = os.path.exists(WExcel_Path+str(chr(47))+WriteFile_Name) 
if Chk_File is True :
    wb.save(WExcel_Path+str(chr(47))+WriteFile_Name) 
    print(WExcel_Path+str(chr(47))+WriteFile_Name)
    print('File Succesfully OverWrited')
else :
    with open(os.path.join(WExcel_Path, WriteFile_Name), 'w') as fp: 
        wb.save(WExcel_Path+str(chr(47))+WriteFile_Name) 
        print(WExcel_Path+str(chr(47))+WriteFile_Name)
        print('File Succesfully NewWrited')
# ---------------------------------------------------------------------------------------------
#-end
# ---------------------------------------------------------------------------------------------
# ---------------------------------------------------------------------------------------------
Content_Structure = Excel_Open.sheet_by_index(3) 
Y_Start = 2
Y_End = 2
while str(Content_Structure.cell(Y_End,1).value)!="Stop" :
    Y_End= Y_End + 1
Y_End = Y_End - 1
X_Start = 2
X_End = 2
while str(Content_Structure.cell(1,X_End).value)!="Stop" :
    X_End= X_End + 1
X_End = X_End - 1
Matrix_Row=(Y_End-Y_Start)+1
Matrix_Col=(X_End-X_Start)+1
# ---------------------------------------------------------------------------------------------
Write_Sheet1 = wb.add_sheet('Sheet2') 
# ---------------------------------------------------------------------------------------------
X_End = X_End + 1 
Y_End = Y_End + 1 # for range increase purppose
Content_Structure_Matrix=[]
temp_str = " "
for x in range(X_Start,X_End):
    for y in range(Y_Start,Y_End):
        temp_str = str(Content_Structure.cell(y,x).value)
        Write_Sheet1.write(y,x, temp_str) 
# ---------------------------------------------------------------------------------------------
Chk_File = os.path.exists(WExcel_Path+str(chr(47))+WriteFile_Name) 
if Chk_File is True :
    wb.save(WExcel_Path+str(chr(47))+WriteFile_Name) 
    print(WExcel_Path+str(chr(47))+WriteFile_Name)
    print('File Succesfully OverWrited')
else :
    with open(os.path.join(WExcel_Path, WriteFile_Name), 'w') as fp: 
        wb.save(WExcel_Path+str(chr(47))+WriteFile_Name) 
        print(WExcel_Path+str(chr(47))+WriteFile_Name)
        print('File Succesfully NewWrited')
# ---------------------------------------------------------------------------------------------