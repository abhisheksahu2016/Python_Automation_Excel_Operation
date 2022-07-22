# -----------File Name or Path List with Path to Excel ------
import os
import xlwt 
from xlwt import Workbook 

path ='C:/Users/ABHISHEK/Desktop/NT/8.NT_MemoryMap'
wb = Workbook() 
Write_Sheet1 = wb.add_sheet('Sheet 1') 
Y_Start = 0
Y_End = len(os.listdir(path))
Excel_YIndex = 0 
for File_Name in os.listdir(path):
    if str(File_Name).endswith(".png"):
        File_Path =""
        File_Path = path +str("/") +str(File_Name)
        # print(File_Name)
        print(File_Path)
        Write_Sheet1.write(Excel_YIndex, 0, File_Path)   
        Excel_YIndex = Excel_YIndex+1
# ---------------------------------------------------------------------------------------------
wb.save('example.xls') 
# ---------------------------------------------------------------------------------------------
