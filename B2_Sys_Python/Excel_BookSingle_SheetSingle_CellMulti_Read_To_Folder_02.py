import os
import xlrd 
from numpy import * 
import numpy as np
import shutil

# Give the location of the file 

Excel_Path = ("F:/1.Fuschia_1.0.8/New/ex.xlsx")  
Excel_Open = xlrd.open_workbook(Excel_Path) 
Folder_Structure = Excel_Open.sheet_by_index(0) 

row_number = 0
column_number = 2
X_Start = 2
X_End = 7
Y_Start = 2
Y_End = 1532
#740

Matrix_Row=(Y_End-Y_Start)+1
Matrix_Col=(X_End-X_Start)+1

# ---------------------------------------------------------------------------------------------
X_End = X_End + 1 
Y_End = Y_End + 1 # for range increase purppose
Folder_Structure_Matrix=[]
for y in range(Y_Start,Y_End):
    temp_row =[] 
    for x in range(X_Start,X_End):   
        cell_value = Folder_Structure.cell(y,x).value
        temp_row.append(cell_value) 
    Folder_Structure_Matrix.append(temp_row) 
print(np.matrix(Folder_Structure_Matrix))
# ---------------------------------------------------------------------------------------------
Folder_Structure_NumpyMatrix = np.array(Folder_Structure_Matrix)
X_End=X_End-1
for y in range(len(Folder_Structure_NumpyMatrix)):
    for x in range((len(Folder_Structure_NumpyMatrix[y])-1),-1,-1):
         if Folder_Structure_NumpyMatrix[y][x]!="" :
             if x==(len(Folder_Structure_NumpyMatrix[y])-1):
                 break
             else :    
                 for k in range(x+1,(len(Folder_Structure_NumpyMatrix[y])),1):
                    Folder_Structure_NumpyMatrix[y][k]="NIL"
                 break
print(np.matrix(Folder_Structure_NumpyMatrix))
print(Folder_Structure_NumpyMatrix.shape)
# ---------------------------------------------------------------------------------------------
lst = [] 
for y in range(len(Folder_Structure_NumpyMatrix)):
    cnt=0
    for x in range(len(Folder_Structure_NumpyMatrix[y])):
        if(Folder_Structure_NumpyMatrix[y][x]==""):
            #print(Folder_Structure_NumpyMatrix[y][x])   
            #print("y : %d x : %d",y,x)
            cnt=cnt+1
    if cnt==(Matrix_Col):
        lst.append(y)
for i in range( len(lst) - 1, -1, -1) :
    Folder_Structure_NumpyMatrix = np.delete(Folder_Structure_NumpyMatrix, lst[i],axis=0)
print(np.matrix(Folder_Structure_NumpyMatrix))
print(Folder_Structure_NumpyMatrix.shape)
# ---------------------------------------------------------------------------------------------
print(Folder_Structure_NumpyMatrix.shape)
for y in range(len(Folder_Structure_NumpyMatrix)):
    for x in range(len(Folder_Structure_NumpyMatrix[y])):   
        if(Folder_Structure_NumpyMatrix[y][x]==""):
            Folder_Structure_NumpyMatrix[y][x]=Folder_Structure_NumpyMatrix[y-1][x]
print(np.matrix(Folder_Structure_NumpyMatrix))
print(Folder_Structure_NumpyMatrix.shape)
# ---------------------------------------------------------------------------------------------
# path = "C:/Users/ABHISHEK/Desktop/Project"
#path = "P:/Professional_Mission/1.Wipro/1.E-FullTime/1.Esubject/1.Trng/1.PBT(PJP)/2.Action/2.Preparation/2.Study/3.PJP-Knlg_Kit/1.PJPCore(.NET)/try"
path = "F:/1.Fuschia_1.0.8/New"
temp_path=""
for i in range(len(Folder_Structure_NumpyMatrix)):
        temp_path = path  
        for j in range(len(Folder_Structure_NumpyMatrix[i])):
                if Folder_Structure_NumpyMatrix[i][j] !="NIL" :
                    # ------------------
                    File_Extension = (".txt", ".pdf", ".docx" ,".pptx",".psd",".png","xlsx",".lnk",".skp")
                    temp_str=str(Folder_Structure_NumpyMatrix[i][j])
                    Chk_CellValueType = 0
                    Chk_CellValueType = temp_str.endswith(File_Extension)   
                    #-------------------------------------------------------------
                    if Chk_CellValueType is True :
                        Chk_Shortcut = temp_str.endswith(".lnk")   
                        if Chk_Shortcut is True : 
                            if 'V-0.0.0.txt.lnk' in temp_str:
                                Src_Path = 'F:/1.Fuschia_1.0.8/3.Pro_Data/2.Fuschia_DTsn/2.Kernal_Space/3.Application_Memory/1.Shaktiman/3.Pro_Data/2.StMn_Datation/3.Information/2.Data/21.Brough/1.xxx_V-0.0.0.txt.lnk'
                                #F:/1.Fuschia_1.0.8/3.Pro_Data/2.Fuschia_DTsn/2.Kernal_Space/3.Application_Memory/Shaktiman/3.Pro_Data/2.StMn_Datation/3.Information/2.Data/21.Brough/1.xxx_V-0.0.0.txt.lnk'
                                Src_Path_Split =Src_Path.split('/')
                                Src_BroughFile =Src_Path_Split[len(Src_Path_Split)-1]
                            else :
                                print('File Wrong')
                            Dst_Path = temp_path
                            Temp_List = Src_Path.split('/')
                            Src_File = Temp_List[len(Temp_List)-1]
                            print(Src_File)
                            Chk_File = os.path.exists(Dst_Path+str(chr(47))+str(Src_File)) 
                            print(Dst_Path+str(chr(47))+str(Src_File))
                            print(Chk_File)
                            if Chk_File is True :
                                    print(Dst_Path+str(chr(47))+str(Src_File))
                                    print('File Existed')
                            else :
                                shutil.copy(Src_Path,Dst_Path)
                                Chk_File = os.path.exists(Dst_Path+str(chr(47))+str(Src_File)) 
                                if Chk_File is True :
                                    print(Dst_Path+str(chr(47))+str(Src_File))
                                    print('File Succesfully Copied')
                                    SrcPath = Dst_Path+str(chr(47))+str(Src_BroughFile)
                                    DstPath = Dst_Path+str(chr(47))+str(Folder_Structure_NumpyMatrix[i][j])
                                    try : 
                                        os.rename(SrcPath, DstPath) 
                                        print('File Succesfully Renamed')
                                    except OSError as error: 
                                        print(error)
                                    print(Dst_Path+str(chr(47))+str(Src_File))
                                else :
                                    print(Dst_Path+str(chr(47))+str(Src_File))
                                    print('File Not Succesfully Copied')
                        else :
                            Dst_Path = temp_path
                            Chk_AlreadFile = os.path.exists(Dst_Path+str(chr(47))+str(Folder_Structure_NumpyMatrix[i][j])) 
                            if Chk_AlreadFile is True :
                                    print(temp_path+str(chr(47))+Folder_Structure_NumpyMatrix[i][j])
                                    print('File Existed')
                            else :
                                with open(os.path.join(temp_path, Folder_Structure_NumpyMatrix[i][j]), 'w') as fp: 
                                    Chk_File = os.path.exists(path+str(chr(47))+Folder_Structure_NumpyMatrix[i][j]) 
                                    if Chk_File is True :
                                        print(temp_path+str(chr(47))+Folder_Structure_NumpyMatrix[i][j])
                                        print('File Succesfully Created')
                                    else :
                                        print(temp_path+str(chr(47))+Folder_Structure_NumpyMatrix[i][j])
                                        print('File Succesfully Not Created')                  
                    #-------------------------------------------------------------
                    else :
                        temp_path = temp_path + "/" 
                        temp_path = temp_path + Folder_Structure_NumpyMatrix[i][j]
                        print(temp_path)
                        isdir = os.path.isdir(temp_path) 
                        t = True
                        f = False
                        if bool(isdir==t) : 
                            print("Folder Existed ")
                        elif bool(isdir==f) :
                            try:
                                os.makedirs(temp_path)
                            except OSError:
                                print ("Successfully not created the directory/n%s" % temp_path)
                            else:
                                print ("Successfully created the directory/n%s " % temp_path)
# ---------------------------------------------------------------------------------------------
#F:/1.Fuschia_1.0.8/3.Pro_Data/2.Fuschia_DTsn/2.Kernal_Space/3.Application_Memory/Shaktiman/3.Pro_Data/2.StMn_Datation/3.Information/2.Data/12.Program/2.Directory/312.FileFolder_Create_Mass.py