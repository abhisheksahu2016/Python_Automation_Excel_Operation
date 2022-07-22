# import the os module
import os
from numpy import * 
import shutil
import os
import xlrd 
from numpy import * 
import numpy as np
import xlwt 
from xlwt import Workbook 

# ------------ Numbering ----------------
Excel_Path = ("C:/Users/ABHISHEK/Desktop/Book2.xlsx")
Excel_Open = xlrd.open_workbook(Excel_Path) 
Folder_Structure = Excel_Open.sheet_by_index(0) 

wb = Workbook() 
Write_Sheet1 = wb.add_sheet('Sheet 1') 

X_Start = 5
X_End = 7
Y_Start = 3
Y_End = 37

Matrix_Row=(Y_End-Y_Start)+1
Matrix_Col=(X_End-X_Start)+1

# ---------------------------------------------------------------------------------------------
X_End = X_End + 1 
Y_End = Y_End + 1 # for range increase purppose
Folder_Structure_Matrix=[]
for x in range(X_Start,X_End):
    y_index = 0   
    for y in range(Y_Start,Y_End):
        if x==0 | x==(X_End-1) :
            if Folder_Structure.cell(y,x).value=="" :
                y_index=0
            else :
                if Folder_Structure.cell(y,x).value!="" :
                    y_index = y_index +1
                    temp_str = str(y_index) + "."
                    temp_str = temp_str + str(Folder_Structure.cell(y,x).value)
                    Write_Sheet1.write(y, x, temp_str) 
        else :
            if Folder_Structure.cell(y,x).value!="" : 
                if Folder_Structure.cell(y,x-1).value!="" :
                    y_index=1
                    temp_str = str(y_index) + "."
                    temp_str = temp_str + str(Folder_Structure.cell(y,x).value)
                    Write_Sheet1.write(y, x, temp_str) 
                else :
                    y_index= y_index+1
                    temp_str = str(y_index) + "."
                    temp_str = temp_str + str(Folder_Structure.cell(y,x).value)
                    Write_Sheet1.write(y, x, temp_str) 
   
# ---------------------------------------------------------------------------------------------
wb.save('example.xls') 
# ---------------------------------------------------------------------------------------------
