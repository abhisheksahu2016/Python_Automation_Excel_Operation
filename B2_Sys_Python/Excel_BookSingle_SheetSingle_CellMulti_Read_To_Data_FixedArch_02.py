import os
import xlrd 
from numpy import * 
import numpy as np
import shutil

# Dragger
# Give the location of the file 
#Excel_Path = ("C:/Users/ABHISHEK/Desktop/abc/Book2.xlsx")  
Excel_Path = ("C:/Users/ABHISHEK/Desktop/Atmn_Grp_Data/A1_Input/aa.xlsx")
#F:/1.Fuschia_1.0.8/3.Pro_Data/2.Fuschia_DTsn/2.Kernal_Space/3.Application_Memory/Shaktiman/3.Pro_Data/2.StMn_Datation/3.Information/2.Data/22.Blueprint/2.Folder_Blpt.xlsx")
Excel_Open = xlrd.open_workbook(Excel_Path) 
Folder_Structure = Excel_Open.sheet_by_index(4) 

row_number = 0
column_number = 2
X_Start = 2
X_End = 11
Y_Start = 2
Y_End =33

Matrix_Row=(Y_End-Y_Start)+1
Matrix_Col=(X_End-X_Start)+1
# ---------------------------------------------------------------------------------------------
# Excel To Matrix-Nromal-Numpy
X_End = X_End + 1 
Y_End = Y_End + 1 # for range increase purppose
Folder_Structure_Matrix=[]
for y in range(Y_Start,Y_End):
    temp_row =[] 
    for x in range(X_Start,X_End):   
        cell_value = Folder_Structure.cell(y,x).value
        temp_row.append(cell_value) 
    Folder_Structure_Matrix.append(temp_row) 
#print(Folder_Structure_Matrix)
Folder_Structure_NumpyMatrix_BeforeFill = np.array(Folder_Structure_Matrix)
Folder_Structure_NumpyMatrix_AfterFill = np.array(Folder_Structure_Matrix)
#print(Folder_Structure_NumpyMatrix_BeforeFill)
# ---------------------------------------------------------------------------------------------
# Filling Right Side Blank-NIL
X_End=X_End-1
for y in range(len(Folder_Structure_NumpyMatrix_AfterFill)):
    for x in range((len(Folder_Structure_NumpyMatrix_AfterFill[y])-1),-1,-1):
         if Folder_Structure_NumpyMatrix_AfterFill[y][x]!="" :
             if x==(len(Folder_Structure_NumpyMatrix_AfterFill[y])-1):
                 break
             else :    
                 for k in range(x+1,(len(Folder_Structure_NumpyMatrix_AfterFill[y])),1):
                    Folder_Structure_NumpyMatrix_AfterFill[y][k]="NIL"
                 break
# ---------------------------------------------------------------------------------------------
# Removing Blank Row
lst = [] 
for y in range(len(Folder_Structure_NumpyMatrix_AfterFill)):
    cnt=0
    for x in range(len(Folder_Structure_NumpyMatrix_AfterFill[y])):
        if(Folder_Structure_NumpyMatrix_AfterFill[y][x]==""):
            #print(Folder_Structure_NumpyMatrix_AfterFill[y][x])   
            #print("y : %d x : %d",y,x)
            cnt=cnt+1
    if cnt==(Matrix_Col):
        lst.append(y)
for i in range( len(lst) - 1, -1, -1) :
    Folder_Structure_NumpyMatrix_AfterFill = np.delete(Folder_Structure_NumpyMatrix_AfterFill, lst[i],axis=0)
#print(Folder_Structure_NumpyMatrix_AfterFill.shape)
# ---------------------------------------------------------------------------------------------
# Filling Left Side Blank - Upper Cell
for y in range(len(Folder_Structure_NumpyMatrix_AfterFill)):
    for x in range(len(Folder_Structure_NumpyMatrix_AfterFill[y])):   
        if(Folder_Structure_NumpyMatrix_AfterFill[y][x]==""):
            Folder_Structure_NumpyMatrix_AfterFill[y][x]=Folder_Structure_NumpyMatrix_AfterFill[y-1][x]
#print(np.matrix(Folder_Structure_NumpyMatrix_AfterFill))
# ---------------------------------------------------------------------------------------------
