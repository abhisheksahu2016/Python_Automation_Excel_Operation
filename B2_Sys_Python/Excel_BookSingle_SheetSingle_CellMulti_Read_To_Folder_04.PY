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
Excel_Path = ("C:/Users/ABHISHEK/Desktop/example.xls")
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
Bad_Chars = ['1','2','3','4','5','6','7','8','9','.']
X_End = X_End + 1 
Y_End = Y_End + 1 # for range increase purppose
Folder_Structure_Matrix=[]
for x in range(X_Start,X_End):
    for y in range(Y_Start,Y_End):
        temp_str=str(Folder_Structure.cell(y,x).value)
        print(temp_str)
        for i in Bad_Chars : 
            test_string = temp_str.replace(i,'') 
        test_string = ''.join([i for i in test_string if not i.isdigit()])
        print(test_string)
        Write_Sheet1.write(y, x, test_string)   
# ---------------------------------------------------------------------------------------------
wb.save('example2.xls') 
# ---------------------------------------------------------------------------------------------
