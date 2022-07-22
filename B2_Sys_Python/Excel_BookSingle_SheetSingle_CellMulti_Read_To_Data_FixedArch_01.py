import os
import xlrd 
from numpy import * 
import numpy as np
import shutil

# Dragger
# Give the location of the file 
#Excel_Path = ("C:/Users/ABHISHEK/Desktop/abc/Book2.xlsx")  
Excel_Path = ("S:/Sahu_Group/B1_Research/B1_Automation-ExcelToData/A1_Input/Fixed_Arch_bb.xlsx")
#F:/1.Fuschia_1.0.8/3.Pro_Data/2.Fuschia_DTsn/2.Kernal_Space/3.Application_Memory/Shaktiman/3.Pro_Data/2.StMn_Datation/3.Information/2.Data/22.Blueprint/2.Folder_Blpt.xlsx")
Excel_Open = xlrd.open_workbook(Excel_Path) 
Folder_Structure = Excel_Open.sheet_by_index(0) 

row_number = 0
column_number = 2
X_Start = 2
X_End = 11
Y_Start = 2
Y_End =16

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
# Mapper
# --
Should_Map_Row = ['Company_Name','SysApp_Name','Module_Name','Dimension_Name','Operation_Name']
Should_Map_Row_Index_X = [3,5,8,9,11]
Row_Index_Y = 4

Company_Name_StartInMatrix_Index = 0
SysApp_Name_StartInMatrix_Index = 0
Module_Name_StartInMatrix_Index = 0
Dimension_Name_StartInMatrix_Index = 0
Operation_Name_StartInMatrix_Index = 0

Group_List=['SahuGroup']
Company_List=[]
Company_SysApp_List=[]
Company_SysApp_Module_List=[]
Company_SysApp_Module_Dimension_List=[]
Company_SysApp_Module_Dimension_Operation_List=[]

User_Index = '000000001'
Group_Index = '001'
Company_Index ='000000000001'
Company_SysApp_Index = '000000000000000000001'
Company_SysApp_Module_Index = '000000000000000000000000000001'
Company_SysApp_Module_Index = '000000000000000000000000000001'
# Map Index
for i in range(0,len(Should_Map_Row)):
    if(Should_Map_Row[i]=='Company_Name'):
        Company_Name_StartInMatrix_Index = Should_Map_Row_Index_X[int(Should_Map_Row.index(Should_Map_Row[i]))]-X_Start
    if(Should_Map_Row[i]=='SysApp_Name'):
        SysApp_Name_StartInMatrix_Index = Should_Map_Row_Index_X[int(Should_Map_Row.index(Should_Map_Row[i]))]-X_Start
    if(Should_Map_Row[i]=='Module_Name'):
        Module_Name_StartInMatrix_Index = Should_Map_Row_Index_X[int(Should_Map_Row.index(Should_Map_Row[i]))]-X_Start
    if(Should_Map_Row[i]=='Dimension_Name'):
        Dimension_Name_StartInMatrix_Index = Should_Map_Row_Index_X[int(Should_Map_Row.index(Should_Map_Row[i]))]-X_Start
    if(Should_Map_Row[i]=='Operation_Name'):
        Operation_Name_StartInMatrix_Index = Should_Map_Row_Index_X[int(Should_Map_Row.index(Should_Map_Row[i]))]-X_Start

# Map Company
for i in range(Row_Index_Y-Y_Start,len(Folder_Structure_NumpyMatrix_BeforeFill)):      
    if Folder_Structure_NumpyMatrix_BeforeFill[i][Company_Name_StartInMatrix_Index] !="NIL" and Folder_Structure_NumpyMatrix_BeforeFill[i][Company_Name_StartInMatrix_Index] !="" :
        Company_List.append(str(Folder_Structure_NumpyMatrix_BeforeFill[i][Company_Name_StartInMatrix_Index]))       
print(Company_List)

# Map SysApp
"""
for i in range(Row_Index_Y-Y_Start,len(Folder_Structure_NumpyMatrix_BeforeFill)):      
    if Folder_Structure_NumpyMatrix_BeforeFill[i][SysApp_Name_StartInMatrix_Index] !="NIL" and Folder_Structure_NumpyMatrix_BeforeFill[i][SysApp_Name_StartInMatrix_Index] !="" :
        Company_SysApp_List.append(str(Folder_Structure_NumpyMatrix_BeforeFill[i][SysApp_Name_StartInMatrix_Index]))       
print(Company_SysApp_List)
"""
for i in range(Row_Index_Y-Y_Start,len(Folder_Structure_NumpyMatrix_BeforeFill)):      
    if Folder_Structure_NumpyMatrix_BeforeFill[i][Company_Name_StartInMatrix_Index] != 'Null' and Folder_Structure_NumpyMatrix_BeforeFill[i][Company_Name_StartInMatrix_Index] != '':
        #print(Folder_Structure_NumpyMatrix_BeforeFill[i][Company_Name_StartInMatrix_Index])
        # ---
        Temp_Company_SysApp_List = []
        for j in range(i,len(Folder_Structure_NumpyMatrix_BeforeFill)):      
            #print(j)
            if Folder_Structure_NumpyMatrix_BeforeFill[j][SysApp_Name_StartInMatrix_Index] !="NIL" and Folder_Structure_NumpyMatrix_BeforeFill[j][SysApp_Name_StartInMatrix_Index] !="" :
                Temp_Company_SysApp_List.append(str(Folder_Structure_NumpyMatrix_BeforeFill[j][SysApp_Name_StartInMatrix_Index]))       
       
            if(j!=len(Folder_Structure_NumpyMatrix_BeforeFill)-1):
                if Folder_Structure_NumpyMatrix_BeforeFill[j+1][Company_Name_StartInMatrix_Index] !="NIL" and Folder_Structure_NumpyMatrix_BeforeFill[j+1][Company_Name_StartInMatrix_Index] !="":                  
                    break                                  
            elif(j==len(Folder_Structure_NumpyMatrix_BeforeFill)-1):
                if Folder_Structure_NumpyMatrix_BeforeFill[j][Company_Name_StartInMatrix_Index] !="NIL" and Folder_Structure_NumpyMatrix_BeforeFill[j][Company_Name_StartInMatrix_Index] !="":                  
                    break                                  
        # ---
        Company_SysApp_List.append(Temp_Company_SysApp_List)
print(Company_SysApp_List)

# Map Module
"""
for i in range(Row_Index_Y-Y_Start,len(Folder_Structure_NumpyMatrix_BeforeFill)):      
    if Folder_Structure_NumpyMatrix_BeforeFill[i][Module_Name_StartInMatrix_Index] !="NIL" and Folder_Structure_NumpyMatrix_BeforeFill[i][Module_Name_StartInMatrix_Index] !="" :
        Company_SysApp_Module_List.append(str(Folder_Structure_NumpyMatrix_BeforeFill[i][Module_Name_StartInMatrix_Index]))       
print(Company_SysApp_Module_List)
"""
for i in range(Row_Index_Y-Y_Start,len(Folder_Structure_NumpyMatrix_BeforeFill)):      
    if Folder_Structure_NumpyMatrix_BeforeFill[i][SysApp_Name_StartInMatrix_Index] != 'Null' and Folder_Structure_NumpyMatrix_BeforeFill[i][SysApp_Name_StartInMatrix_Index] != '':
        #print(Folder_Structure_NumpyMatrix_BeforeFill[i][SysApp_Name_StartInMatrix_Index])
        # ---
        Temp_Company_SysApp_Module_List = []
        for j in range(i,len(Folder_Structure_NumpyMatrix_BeforeFill)):      
            #print(j)
            if Folder_Structure_NumpyMatrix_BeforeFill[j][Module_Name_StartInMatrix_Index] !="NIL" and Folder_Structure_NumpyMatrix_BeforeFill[j][Module_Name_StartInMatrix_Index] !="" :
                Temp_Company_SysApp_Module_List.append(str(Folder_Structure_NumpyMatrix_BeforeFill[j][Module_Name_StartInMatrix_Index]))       
            if(j!=len(Folder_Structure_NumpyMatrix_BeforeFill)-1):
                if Folder_Structure_NumpyMatrix_BeforeFill[j+1][SysApp_Name_StartInMatrix_Index] !="NIL" and Folder_Structure_NumpyMatrix_BeforeFill[j+1][SysApp_Name_StartInMatrix_Index] !="":                  
                    break                                  
            elif(j==len(Folder_Structure_NumpyMatrix_BeforeFill)-1):
                if Folder_Structure_NumpyMatrix_BeforeFill[j][SysApp_Name_StartInMatrix_Index] !="NIL" and Folder_Structure_NumpyMatrix_BeforeFill[j][SysApp_Name_StartInMatrix_Index] !="":                  
                    break                                  
        # ---
        Company_SysApp_Module_List.append(Temp_Company_SysApp_Module_List)
print(Company_SysApp_Module_List)

# Map Dimension
"""
for i in range(Row_Index_Y-Y_Start,len(Folder_Structure_NumpyMatrix_BeforeFill)):      
    if Folder_Structure_NumpyMatrix_BeforeFill[i][Dimension_Name_StartInMatrix_Index] !="NIL" and Folder_Structure_NumpyMatrix_BeforeFill[i][Dimension_Name_StartInMatrix_Index] !="" :
        Company_SysApp_Module_Dimension_List.append(str(Folder_Structure_NumpyMatrix_BeforeFill[i][Dimension_Name_StartInMatrix_Index]))       
print(Company_SysApp_Module_Dimension_List)
"""
for i in range(Row_Index_Y-Y_Start,len(Folder_Structure_NumpyMatrix_BeforeFill)):      
    if Folder_Structure_NumpyMatrix_BeforeFill[i][Module_Name_StartInMatrix_Index] != 'Null' and Folder_Structure_NumpyMatrix_BeforeFill[i][Module_Name_StartInMatrix_Index] != '':
        #print(Folder_Structure_NumpyMatrix_BeforeFill[i][SysApp_Name_StartInMatrix_Index])
        # ---
        Temp_Company_SysApp_Module_Dimension_List = []
        for j in range(i,len(Folder_Structure_NumpyMatrix_BeforeFill)):      
            #print(j)
            if Folder_Structure_NumpyMatrix_BeforeFill[j][Dimension_Name_StartInMatrix_Index] !="NIL" and Folder_Structure_NumpyMatrix_BeforeFill[j][Dimension_Name_StartInMatrix_Index] !="" :
                Temp_Company_SysApp_Module_Dimension_List.append(str(Folder_Structure_NumpyMatrix_BeforeFill[j][Dimension_Name_StartInMatrix_Index]))       
            if(j!=len(Folder_Structure_NumpyMatrix_BeforeFill)-1):
                if Folder_Structure_NumpyMatrix_BeforeFill[j+1][Module_Name_StartInMatrix_Index] !="NIL" and Folder_Structure_NumpyMatrix_BeforeFill[j+1][Module_Name_StartInMatrix_Index] !="":                  
                    break                                  
            elif(j==len(Folder_Structure_NumpyMatrix_BeforeFill)-1):
                if Folder_Structure_NumpyMatrix_BeforeFill[j][Module_Name_StartInMatrix_Index] !="NIL" and Folder_Structure_NumpyMatrix_BeforeFill[j][Module_Name_StartInMatrix_Index] !="":                  
                    break                                  
        # ---
        Company_SysApp_Module_Dimension_List.append(Temp_Company_SysApp_Module_Dimension_List)
print(Company_SysApp_Module_Dimension_List)

# Map Dimension
"""
for i in range(Row_Index_Y-Y_Start,len(Folder_Structure_NumpyMatrix_BeforeFill)):      
    if Folder_Structure_NumpyMatrix_BeforeFill[i][Dimension_Name_StartInMatrix_Index] !="NIL" and Folder_Structure_NumpyMatrix_BeforeFill[i][Dimension_Name_StartInMatrix_Index] !="" :
        Company_SysApp_Module_Dimension_Operation_List.append(str(Folder_Structure_NumpyMatrix_BeforeFill[i][Dimension_Name_StartInMatrix_Index]))       
print(Company_SysApp_Module_Dimension_Operation_List)
"""
for i in range(Row_Index_Y-Y_Start,len(Folder_Structure_NumpyMatrix_BeforeFill)):      
    if Folder_Structure_NumpyMatrix_BeforeFill[i][Dimension_Name_StartInMatrix_Index] != 'Null' and Folder_Structure_NumpyMatrix_BeforeFill[i][Dimension_Name_StartInMatrix_Index] != '':
        #print(Folder_Structure_NumpyMatrix_BeforeFill[i][SysApp_Name_StartInMatrix_Index])
        # ---
        Temp_Company_SysApp_Module_Dimension_Operation_List = []
        for j in range(i,len(Folder_Structure_NumpyMatrix_BeforeFill)):      
            #print(j)
            if Folder_Structure_NumpyMatrix_BeforeFill[j][Operation_Name_StartInMatrix_Index] !="NIL" and Folder_Structure_NumpyMatrix_BeforeFill[j][Operation_Name_StartInMatrix_Index] !="" :
                Temp_Company_SysApp_Module_Dimension_Operation_List.append(str(Folder_Structure_NumpyMatrix_BeforeFill[j][Operation_Name_StartInMatrix_Index]))       
            if(j!=len(Folder_Structure_NumpyMatrix_BeforeFill)-1):
                if Folder_Structure_NumpyMatrix_BeforeFill[j+1][Dimension_Name_StartInMatrix_Index] !="NIL" and Folder_Structure_NumpyMatrix_BeforeFill[j+1][Dimension_Name_StartInMatrix_Index] !="":                  
                    break                                  
            elif(j==len(Folder_Structure_NumpyMatrix_BeforeFill)-1):
                if Folder_Structure_NumpyMatrix_BeforeFill[j][Dimension_Name_StartInMatrix_Index] !="NIL" and Folder_Structure_NumpyMatrix_BeforeFill[j][Dimension_Name_StartInMatrix_Index] !="":                  
                    break                                  
        # ---
        Company_SysApp_Module_Dimension_Operation_List.append(Temp_Company_SysApp_Module_Dimension_Operation_List)
print(Company_SysApp_Module_Dimension_Operation_List)
# Map Space
Space_Row_Index_Y = 0
Space_Row_Index_X = 0

Default_Space_List = []
Vertical_Space_List = []
Horizontal_Space_List = []
#--
Space_Row_Index_Y = 2
Space_Row_Index_X = 3

Space_Row_Index_Y_InMatrix = Space_Row_Index_Y - Y_Start
Space_Row_Index_X_InMatrix = Space_Row_Index_X - X_Start
"""
Temp_Company_Space_List=[]
for i in range(Space_Row_Index_X_InMatrix,len(Folder_Structure_NumpyMatrix_BeforeFill[0])):
    if Folder_Structure_NumpyMatrix_BeforeFill[Space_Row_Index_Y_InMatrix][i] != 'Null' and Folder_Structure_NumpyMatrix_BeforeFill[Space_Row_Index_Y_InMatrix][i] != '':
        #print(Folder_Structure_NumpyMatrix_BeforeFill[Space_Row_Index_Y_InMatrix][i])
        Temp_Company_Space_List.append(Folder_Structure_NumpyMatrix_BeforeFill[Space_Row_Index_Y_InMatrix][i])
Company_Space_List.append(Temp_Company_Space_List)
print(Company_Space_List)
"""
Default_Space_List = ['A1_Indx','B1_Group','C1_Company','D1_User']
Vertical_Space_List = ['Indx_Space','Group_Space','Company_Space','User_Space']
Horizontal_Space_List = [[],['HW_Space'],['HW_Space','Krnl_Space','App_Space','Mdl_Space','On_Space'],[]]
# Map Srusti_Start_Name
Srusti_Start_Name='Sahu_Group'

# --------------------------------------------------------------------------------------------
# Generator
"""
Company_List = ['Fuschia', 'Halal']
Company_SysApp_List = [['My Shaktiman ', 'My Kalal2'], ['My Shaktiman 2', 'My Kalal']]
Company_SysApp_Module_List = [['My Profile', 'My Body', 'My Time'], ['My Time'], ['My Profile', 'My Body', 'My Time'], ['My Time']]
"""
Output_Path = 'S:/Sahu_Group/B1_Research/B1_Automation-ExcelToData/C1_Output'
# Generator-A1_Plan
# Generator-B1_Code
# Generator-C1_Data
Default_Data_Folder=[['C1_Data'],['A1_DBFile_SQLServer','B1_Script_SQL','C1_Raw_FldrFile','D1_Core_Excel']]
Default_Data_Folder_Location ="C:/Users/ABHISHEK/Desktop/Temp/On00000000"

Default_Data_Folder_Location_Dir_List = os.listdir(Default_Data_Folder_Location)

# --
for Level_Index in range(0,len(Default_Data_Folder[1])):
    temp_path = Output_Path + "/" + str(Default_Data_Folder[0][0]) +  "/" + str(Default_Data_Folder[1][Level_Index])
    #print(temp_path)
    isdir = os.path.isdir(temp_path) 
    t = True
    f = False
    if bool(isdir==t) : 
        pass
        #print("Folder Existed ")
    elif bool(isdir==f) :
        try:
            os.makedirs(temp_path)
        except OSError:
            print ("Successfully not created the directory\n%s" % temp_path)
        else:
            print ("Successfully created the directory\n%s " % temp_path)

# Generator-C1_Data-C1_Raw_FldrFile
B1_Script_SQL_Path = Output_Path + "/" + str(Default_Data_Folder[0][0]) +  "/" + str(Default_Data_Folder[1][1])
#print(B1_Script_SQL_Path)
C1_Raw_FldrFile_Path = Output_Path + "/" + str(Default_Data_Folder[0][0]) +  "/" + str(Default_Data_Folder[1][2])
#print(C1_Raw_FldrFile_Path)
# --
Default_Data_Folder_Location = "C:/Users/ABHISHEK/Desktop/Temp/On00000000"
# --
Vertical_CharPrefix_StartFrom = 65
Vertical_CharPrefix_Num = 1
# --
for i in range(0,len(Vertical_Space_List)):
    """   
    # --   
    if(Default_Space_List[i].find('Indx')!=-1 or Default_Space_List[i].find('User')!=-1 ):
        if(len(Horizontal_Space_List[i])==0):
            #--
            # Need 2 string to AVOID 2 directory(1 with space,2 with out)
            temp_path = C1_Raw_FldrFile_Path + "/" +str(chr(Vertical_CharPrefix_StartFrom))+str(Vertical_CharPrefix_Num)+'_'+Vertical_Space_List[i]+'Space'
            print(temp_path)
            isdir = os.path.isdir(temp_path) 
            t = True
            f = False
            if bool(isdir==t): 
                # --
                print("Folder Existed ")
                # --
                for i in range(0,len(Default_Data_Folder_Location_Dir_List)):                            
                    Data_Path = temp_path+"/"+str(Default_Data_Folder_Location_Dir_List[i])
                    isdir1 = os.path.isdir(Data_Path) 
                    t1 = True
                    f1 = False
                    if bool(isdir1==t1): 
                        pass
                    else :
                        try:
                            # --Dummy Copy Folder
                            from distutils.dir_util import copy_tree
                            copy_tree(Default_Data_Folder_Location+str(Default_Data_Folder_Location_Dir_List[i]),temp_path)
                        except OSError:
                            print ("Successfully not created the data directory in Existed Folder\n%s" % data_path)
                        else:   
                            print ("Successfully created the data directory in Existed Folder\n%s" % data_path)
                print("Data Folder Checked ")
            elif bool(isdir==f) :
                try:
                    os.makedirs(temp_path)
                except OSError:
                    print ("Successfully not created the directory\n%s" % temp_path)
                else:
                    #--
                    print ("Successfully created the directory\n%s " % temp_path)
                    # -- Dummy Copy Folder
                    from distutils.dir_util import copy_tree
                    copy_tree(Default_Data_Folder_Location,temp_path)
                    Current_Data_Folder_Location_Dir_List = os.listdir(temp_path)
                    # --
                    import collections
                    if collections.Counter(Current_Data_Folder_Location_Dir_List) == collections.Counter(Current_Data_Folder_Location_Dir_List):
                        print ("Successfully created the data directory inside\n%s" % temp_path)
                    else :
                        print ("Successfully not created the data directory inside\n%s" % temp_path)
            #--
        elif(len(Horizontal_Space_List[i])>0):
            pass
    # --
    # --
    if(Default_Space_List[i].find('Group')!=-1 ):
        if(len(Horizontal_Space_List[i])==0):
            #--
            Group_Serial_Index = 0
            # --
            Default_Data_Folder_Location_Dir_List = os.listdir(Default_Data_Folder_Location)
            # --
            for Group_List_Index in range(0,len(Group_List)):    
                temp_path = C1_Raw_FldrFile_Path + "/" +str(chr(Vertical_CharPrefix_StartFrom))+str(Vertical_CharPrefix_Num)+'_'+ Group_List[Group_Serial_Index]+'Space'
                print(temp_path)
                isdir = os.path.isdir(temp_path) 
                t = True
                f = False
                if bool(isdir==t): 
                    # --
                    print("Folder Existed ")
                    # --
                    for i in range(0,len(Default_Data_Folder_Location_Dir_List)):                            
                        Data_Path = temp_path+"/"+str(Default_Data_Folder_Location_Dir_List[i])
                        isdir1 = os.path.isdir(Data_Path) 
                        t1 = True
                        f1 = False
                        if bool(isdir1==t1): 
                            pass
                        else :
                            try:
                                # --Dummy Copy Folder
                                from distutils.dir_util import copy_tree
                                copy_tree(Default_Data_Folder_Location+str(Default_Data_Folder_Location_Dir_List[i]),temp_path)
                            except OSError:
                                print ("Successfully not created the data directory in Existed Folder\n%s" % data_path)
                            else:   
                                print ("Successfully created the data directory in Existed Folder\n%s" % data_path)
                    print("Data Folder Checked ")
                elif bool(isdir==f) :
                    try:
                        os.makedirs(temp_path)
                    except OSError:
                        print ("Successfully not created the directory\n%s" % temp_path)
                    else:
                        #--
                        print ("Successfully created the directory\n%s " % temp_path)
                        # -- Dummy Copy Folder
                        from distutils.dir_util import copy_tree
                        copy_tree(Default_Data_Folder_Location,temp_path)
                        Current_Data_Folder_Location_Dir_List = os.listdir(temp_path)
                        # --
                        import collections
                        if collections.Counter(Current_Data_Folder_Location_Dir_List) == collections.Counter(Current_Data_Folder_Location_Dir_List):
                            print ("Successfully created the data directory inside\n%s" % temp_path)
                        else :
                            print ("Successfully not created the data directory inside\n%s" % temp_path)
                #--
                #--
                Group_Serial_Index = Group_Serial_Index + 1
        elif(len(Horizontal_Space_List[i])>0):
            Horizontal_CharPrefix_StartFrom = 65
            Horizontal_CharPrefix_Num = 1
            for j in range(0,len(Horizontal_Space_List[i])):
                #--
                Group_Serial_Index = 0
                # --
                Default_Data_Folder_Location_Dir_List = os.listdir(Default_Data_Folder_Location)
                # --
                for Group_List_Index in range(0,len(Group_List)):    
                    temp_path = C1_Raw_FldrFile_Path + "/" +str(chr(Vertical_CharPrefix_StartFrom))+str(Vertical_CharPrefix_Num)+'_'+ Group_List[Group_Serial_Index]+'Space'+str(chr(Horizontal_CharPrefix_StartFrom))+str(Horizontal_CharPrefix_Num)+'_'+ Horizontal_Space_List[i]+'Space'
                    print(temp_path)
                    isdir = os.path.isdir(temp_path) 
                    t = True
                    f = False
                    if bool(isdir==t): 
                        # --
                        print("Folder Existed ")
                        # --
                        for i in range(0,len(Default_Data_Folder_Location_Dir_List)):                            
                            Data_Path = temp_path+"/"+str(Default_Data_Folder_Location_Dir_List[i])
                            isdir1 = os.path.isdir(Data_Path) 
                            t1 = True
                            f1 = False
                            if bool(isdir1==t1): 
                                pass
                            else :
                                try:
                                    # --Dummy Copy Folder
                                    from distutils.dir_util import copy_tree
                                    copy_tree(Default_Data_Folder_Location+str(Default_Data_Folder_Location_Dir_List[i]),temp_path)
                                except OSError:
                                    print ("Successfully not created the data directory in Existed Folder\n%s" % data_path)
                                else:   
                                    print ("Successfully created the data directory in Existed Folder\n%s" % data_path)
                        print("Data Folder Checked ")
                    elif bool(isdir==f) :
                        try:
                            os.makedirs(temp_path)
                        except OSError:
                            print ("Successfully not created the directory\n%s" % temp_path)
                        else:
                            #--
                            print ("Successfully created the directory\n%s " % temp_path)
                            # -- Dummy Copy Folder
                            from distutils.dir_util import copy_tree
                            copy_tree(Default_Data_Folder_Location,temp_path)
                            Current_Data_Folder_Location_Dir_List = os.listdir(temp_path)
                            # --
                            import collections
                            if collections.Counter(Current_Data_Folder_Location_Dir_List) == collections.Counter(Current_Data_Folder_Location_Dir_List):
                                print ("Successfully created the data directory inside\n%s" % temp_path)
                            else :
                                print ("Successfully not created the data directory inside\n%s" % temp_path)
                    #--
                    #--
                    Group_Serial_Index = Group_Serial_Index + 1
            Horizontal_CharPrefix_StartFrom = Horizontal_CharPrefix_StartFrom +1
    # --
    """
    # --   
    if(Default_Space_List[i].find('Company')!=-1):
        #--
        Company_Serial_Index = 0
        SysApp_Serial_Index = 0
        Module_Serial_Index = 0
        Dimension_Serial_Index = 0
        Operation_Serial_Index = 0
        # --
        Default_Data_Folder_Location_Dir_List = os.listdir(Default_Data_Folder_Location)
        # --
        for Company_List_Index in range(0,len(Company_List)):    
            #print(Company_List[Company_List_Index])
            #print(Company_Serial_Index)
            #--
            for Company_SysApp_List_Index in range(0,len(Company_SysApp_List[Company_Serial_Index])):    
                #print(Company_List[Company_Serial_Index]+ "/" + Company_SysApp_List[Company_Serial_Index][Company_SysApp_List_Index])
                #print(SysApp_Serial_Index)
                #--
                for Company_SysApp_Module_List_Index in range(0,len( Company_SysApp_Module_List[SysApp_Serial_Index])):    
                    #print(Company_List[Company_Serial_Index]+ "/" + Company_SysApp_List[Company_Serial_Index][Company_SysApp_List_Index] + "/" + Company_SysApp_Module_List[SysApp_Serial_Index][Company_SysApp_Module_List_Index])
                    #print(Module_Serial_Index)
                    #--
                    for Company_SysApp_Module_Dimension_List_Index in range(0,len( Company_SysApp_Module_Dimension_List[Module_Serial_Index])):    
                        #print(Company_List[Company_Serial_Index]+ "/" + Company_SysApp_List[Company_Serial_Index][Company_SysApp_List_Index] + "/" + Company_SysApp_Module_List[SysApp_Serial_Index][Company_SysApp_Module_List_Index] + "/" + Company_SysApp_Module_Dimension_List[Module_Serial_Index][Company_SysApp_Module_Dimension_List_Index])
                        #print(Dimension_Serial_Index)
                        #--
                        for Company_SysApp_Module_Dimension_Operation_List_Index in range(0,len( Company_SysApp_Module_Dimension_Operation_List[Dimension_Serial_Index])):    
                            #print(Company_List[Company_Serial_Index]+ "/" + Company_SysApp_List[Company_Serial_Index][Company_SysApp_List_Index] + "/" + Company_SysApp_Module_List[SysApp_Serial_Index][Company_SysApp_Module_List_Index] + "/" + Company_SysApp_Module_Dimension_List[Module_Serial_Index][Company_SysApp_Module_Dimension_List_Index] + "/" + Company_SysApp_Module_Dimension_Operation_List[Dimension_Serial_Index][Company_SysApp_Module_Dimension_Operation_List_Index] )
                            #print(Operation_Serial_Index)
                            #--
                            # Need 2 string to AVOID 2 directory(1 with space,2 with out)
                            temp_path = C1_Raw_FldrFile_Path + "/" +str(chr(Vertical_CharPrefix_StartFrom))+str(Vertical_CharPrefix_Num)+'_'+ Company_List[Company_Serial_Index]+'Space'+"/" + Company_SysApp_List[Company_Serial_Index][Company_SysApp_List_Index] + "/" + Company_SysApp_Module_List[SysApp_Serial_Index][Company_SysApp_Module_List_Index] + "/" + Company_SysApp_Module_Dimension_List[Module_Serial_Index][Company_SysApp_Module_Dimension_List_Index] + "/" + Company_SysApp_Module_Dimension_Operation_List[Dimension_Serial_Index][Company_SysApp_Module_Dimension_Operation_List_Index]
                            print(temp_path)
                            isdir = os.path.isdir(temp_path) 
                            t = True
                            f = False
                            if bool(isdir==t): 
                                # --
                                print("Folder Existed ")
                                # --
                                for i in range(0,len(Default_Data_Folder_Location_Dir_List)):                            
                                    Data_Path = temp_path+"/"+str(Default_Data_Folder_Location_Dir_List[i])
                                    isdir1 = os.path.isdir(Data_Path) 
                                    t1 = True
                                    f1 = False
                                    if bool(isdir1==t1): 
                                        pass
                                    else :
                                        try:
                                            # --Dummy Copy Folder
                                            from distutils.dir_util import copy_tree
                                            copy_tree(Default_Data_Folder_Location+str(Default_Data_Folder_Location_Dir_List[i]),temp_path)
                                        except OSError:
                                            print ("Successfully not created the data directory in Existed Folder\n%s" % data_path)
                                        else:   
                                            print ("Successfully created the data directory in Existed Folder\n%s" % data_path)
                                print("Data Folder Checked ")
                            elif bool(isdir==f) :
                                try:
                                    os.makedirs(temp_path)
                                except OSError:
                                    print ("Successfully not created the directory\n%s" % temp_path)
                                else:
                                    #--
                                    print ("Successfully created the directory\n%s " % temp_path)
                                    # -- Dummy Copy Folder
                                    from distutils.dir_util import copy_tree
                                    copy_tree(Default_Data_Folder_Location,temp_path)
                                    Current_Data_Folder_Location_Dir_List = os.listdir(temp_path)
                                    # --
                                    import collections
                                    if collections.Counter(Current_Data_Folder_Location_Dir_List) == collections.Counter(Current_Data_Folder_Location_Dir_List):
                                        print ("Successfully created the data directory inside\n%s" % temp_path)
                                    else :
                                        print ("Successfully not created the data directory inside\n%s" % temp_path)
                            #--
                            Operation_Serial_Index = Operation_Serial_Index + 1
                        #--
                        Dimension_Serial_Index = Dimension_Serial_Index + 1
                    #--
                    Module_Serial_Index = Module_Serial_Index + 1 
                #--
                SysApp_Serial_Index = SysApp_Serial_Index + 1 
            #--
            Company_Serial_Index = Company_Serial_Index + 1
    Vertical_CharPrefix_StartFrom = Vertical_CharPrefix_StartFrom+1


# Generator-C1_Data-D1_Core_Excel
#-
D1_Core_Excel_Path = Output_Path + "/" + str(Default_Data_Folder[0][0]) +  "/" + str(Default_Data_Folder[1][3])
#print(D1_Core_Excel_Path)
# --
Vertical_CharPrefix_StartFrom = 65
Vertical_CharPrefix_Num = 1
# --
for i in range(0,len(Vertical_Space_List)):
    # --   
    if(Default_Space_List[i].find('Indx')!=-1 or Default_Space_List[i].find('User')!=-1):
        if(len(Horizontal_Space_List[i])==0):
            excel_name = str(chr(Vertical_CharPrefix_StartFrom))+str(Vertical_CharPrefix_Num)+'_'+Vertical_Space_List[i].replace("_Space",'Space') +'.xlsx'
            temp_excel_path  = D1_Core_Excel_Path + "/" + excel_name
            #print(temp_excel_path)
            Chk_AlreadFile = os.path.exists(temp_excel_path) 
            if Chk_AlreadFile is True :
                print('File Existed')
                print(temp_excel_path)
            else :
                with open(os.path.join(D1_Core_Excel_Path, excel_name), 'w') as fp: 
                    Chk_File = os.path.exists(temp_excel_path) 
                    if Chk_File is True :
                        print('File Succesfully Created')
                        print(temp_excel_path)
                    else :
                        print('File Succesfully Not Created')                  
                        print(temp_excel_path)
        elif(len(Horizontal_Space_List[i])>0):
            Horizontal_CharPrefix_StartFrom = 65
            Horizontal_CharPrefix_Num = 1
            for j in range(0,len(Horizontal_Space_List[i])):
                excel_name = str(chr(Vertical_CharPrefix_StartFrom))+str(Vertical_CharPrefix_Num)+'_'+Vertical_Space_List[i].replace("_Space",'Space')+'_'+str(chr(Horizontal_CharPrefix_StartFrom))+str(Horizontal_CharPrefix_Num)+'_'+Horizontal_Space_List[i][j].replace("_Space",'Space')+'.xlsx'
                temp_excel_path  = D1_Core_Excel_Path + "/" + excel_name
                #print(temp_excel_path)
                Chk_AlreadFile = os.path.exists(temp_excel_path) 
                if Chk_AlreadFile is True :
                    print('File Existed')
                    print(temp_excel_path)
                else :
                    with open(os.path.join(D1_Core_Excel_Path, excel_name), 'w') as fp: 
                        Chk_File = os.path.exists(temp_excel_path) 
                        if Chk_File is True :
                            print('File Succesfully Created')
                            print(temp_excel_path)
                        else :
                            print('File Succesfully Not Created')                  
                            print(temp_excel_path)
                # --
                Horizontal_CharPrefix_StartFrom = Horizontal_CharPrefix_StartFrom +1
    # --
    if(Default_Space_List[i].find('Group')!=-1 ):
        if(len(Horizontal_Space_List[i])==0):
            excel_name = str(chr(Vertical_CharPrefix_StartFrom))+str(Vertical_CharPrefix_Num)+'_'+Group_List[0]+'Space' +'.xlsx'
            temp_excel_path  = D1_Core_Excel_Path + "/" + excel_name
            #print(temp_excel_path)
            Chk_AlreadFile = os.path.exists(temp_excel_path) 
            if Chk_AlreadFile is True :
                print('File Existed')
                print(temp_excel_path)
            else :
                with open(os.path.join(D1_Core_Excel_Path, excel_name), 'w') as fp: 
                    Chk_File = os.path.exists(temp_excel_path) 
                    if Chk_File is True :
                        print('File Succesfully Created')
                        print(temp_excel_path)
                    else :
                        print('File Succesfully Not Created')                  
                        print(temp_excel_path)
        elif(len(Horizontal_Space_List[i])>0):
            Horizontal_CharPrefix_StartFrom = 65
            Horizontal_CharPrefix_Num = 1
            for j in range(0,len(Horizontal_Space_List[i])):
                excel_name = str(chr(Vertical_CharPrefix_StartFrom))+str(Vertical_CharPrefix_Num)+'_'+Group_List[0]+'Space'+'_'+str(chr(Horizontal_CharPrefix_StartFrom))+str(Horizontal_CharPrefix_Num)+'_'+Horizontal_Space_List[i][j].replace("_Space",'Space')+'.xlsx'
                temp_excel_path  = D1_Core_Excel_Path + "/" + excel_name
                #print(temp_excel_path)
                Chk_AlreadFile = os.path.exists(temp_excel_path) 
                if Chk_AlreadFile is True :
                    print('File Existed')
                    print(temp_excel_path)
                else :
                    with open(os.path.join(D1_Core_Excel_Path, excel_name), 'w') as fp: 
                        Chk_File = os.path.exists(temp_excel_path) 
                        if Chk_File is True :
                            print('File Succesfully Created')
                            print(temp_excel_path)
                        else :
                            print('File Succesfully Not Created')                  
                            print(temp_excel_path)
                # --
                Horizontal_CharPrefix_StartFrom = Horizontal_CharPrefix_StartFrom +1
    # --
    # --   
    if(Default_Space_List[i].find('Company')!=-1):
        if(len(Horizontal_Space_List[i])==0):           
            pass
        elif(len(Horizontal_Space_List[i])>0):
            for k in range(0,len(Company_List)):
                Horizontal_CharPrefix_StartFrom = 65
                Horizontal_CharPrefix_Num = 1
                for j in range(0,len(Horizontal_Space_List[i])):             
                    if(Horizontal_Space_List[i][j].find('On_Space')!=-1):
                        excel_name =""
                        #--
                        Company_Serial_Index = 0
                        SysApp_Serial_Index = 0
                        Module_Serial_Index = 0
                        Dimension_Serial_Index = 0
                        Operation_Serial_Index = 0
                        # --
                        for Company_List_Index in range(0,len(Company_List)):    
                            for Company_SysApp_List_Index in range(0,len(Company_SysApp_List[Company_Serial_Index])):    
                                for Company_SysApp_Module_List_Index in range(0,len( Company_SysApp_Module_List[SysApp_Serial_Index])):    
                                    excel_name = str(chr(Vertical_CharPrefix_StartFrom))+str(Vertical_CharPrefix_Num)+'_'+ Company_List[k]+'Space'+'_'+Company_SysApp_Module_List[SysApp_Serial_Index][Company_SysApp_Module_List_Index]+'_'+str(chr(Horizontal_CharPrefix_StartFrom))+str(Horizontal_CharPrefix_Num)+'_'+Horizontal_Space_List[i][j].replace("_Space",'Space')+'.xlsx'
                                    temp_excel_path  = D1_Core_Excel_Path + "/" + excel_name
                                    #print(temp_excel_path)
                                    Chk_AlreadFile = os.path.exists(temp_excel_path) 
                                    if Chk_AlreadFile is True :
                                        print('File Existed')
                                        print(temp_excel_path)
                                    else :
                                        with open(os.path.join(D1_Core_Excel_Path, excel_name), 'w') as fp: 
                                            Chk_File = os.path.exists(temp_excel_path) 
                                            if Chk_File is True :
                                                print('File Succesfully Created')
                                                print(temp_excel_path)
                                            else :
                                                print('File Succesfully Not Created')                  
                                                print(temp_excel_path)
                                    # --
                        Horizontal_CharPrefix_StartFrom = Horizontal_CharPrefix_StartFrom +1
                    else :
                        excel_name = str(chr(Vertical_CharPrefix_StartFrom))+str(Vertical_CharPrefix_Num)+'_'+ Company_List[k]+'Space'+'_'+str(chr(Horizontal_CharPrefix_StartFrom))+str(Horizontal_CharPrefix_Num)+'_'+Horizontal_Space_List[i][j].replace("_Space",'Space')+'.xlsx'
                        temp_excel_path  = D1_Core_Excel_Path + "/" + excel_name
                        #print(temp_excel_path)
                        Chk_AlreadFile = os.path.exists(temp_excel_path) 
                        if Chk_AlreadFile is True :
                            print('File Existed')
                            print(temp_excel_path)
                        else :
                            with open(os.path.join(D1_Core_Excel_Path, excel_name), 'w') as fp: 
                                Chk_File = os.path.exists(temp_excel_path) 
                                if Chk_File is True :
                                    print('File Succesfully Created')
                                    print(temp_excel_path)
                                else :
                                    print('File Succesfully Not Created')                  
                                    print(temp_excel_path)
                        # --
                        Horizontal_CharPrefix_StartFrom = Horizontal_CharPrefix_StartFrom +1
    Vertical_CharPrefix_StartFrom = Vertical_CharPrefix_StartFrom+1
# Generator-C1_Data-B1_Script_SQL
#-
B1_Script_SQL_Path = Output_Path + "/" + str(Default_Data_Folder[0][0]) +  "/" + str(Default_Data_Folder[1][1])
#print(B1_Script_SQL_Path)
# --
Vertical_CharPrefix_StartFrom = 65
Vertical_CharPrefix_Num = 1
# --
for i in range(0,len(Vertical_Space_List)):
    # --   
    if(Default_Space_List[i].find('Indx')!=-1 or Default_Space_List[i].find('User')!=-1):
        if(len(Horizontal_Space_List[i])==0):
            excel_name = str(chr(Vertical_CharPrefix_StartFrom))+str(Vertical_CharPrefix_Num)+'_'+Vertical_Space_List[i].replace("_Space",'Space') +'.sql'
            temp_excel_path  = B1_Script_SQL_Path + "/" + excel_name
            #print(temp_excel_path)
            Chk_AlreadFile = os.path.exists(temp_excel_path) 
            if Chk_AlreadFile is True :
                print('File Existed')
                print(temp_excel_path)
            else :
                with open(os.path.join(B1_Script_SQL_Path, excel_name), 'w') as fp: 
                    Chk_File = os.path.exists(temp_excel_path) 
                    if Chk_File is True :
                        print('File Succesfully Created')
                        print(temp_excel_path)
                    else :
                        print('File Succesfully Not Created')                  
                        print(temp_excel_path)
        elif(len(Horizontal_Space_List[i])>0):
            Horizontal_CharPrefix_StartFrom = 65
            Horizontal_CharPrefix_Num = 1
            for j in range(0,len(Horizontal_Space_List[i])):
                excel_name = str(chr(Vertical_CharPrefix_StartFrom))+str(Vertical_CharPrefix_Num)+'_'+Vertical_Space_List[i].replace("_Space",'Space')+'_'+str(chr(Horizontal_CharPrefix_StartFrom))+str(Horizontal_CharPrefix_Num)+'_'+Horizontal_Space_List[i][j].replace("_Space",'Space')+'.sql'
                temp_excel_path  = B1_Script_SQL_Path + "/" + excel_name
                #print(temp_excel_path)
                Chk_AlreadFile = os.path.exists(temp_excel_path) 
                if Chk_AlreadFile is True :
                    print('File Existed')
                    print(temp_excel_path)
                else :
                    with open(os.path.join(B1_Script_SQL_Path, excel_name), 'w') as fp: 
                        Chk_File = os.path.exists(temp_excel_path) 
                        if Chk_File is True :
                            print('File Succesfully Created')
                            print(temp_excel_path)
                        else :
                            print('File Succesfully Not Created')                  
                            print(temp_excel_path)
                # --
                Horizontal_CharPrefix_StartFrom = Horizontal_CharPrefix_StartFrom +1
    # --
    # --
    if(Default_Space_List[i].find('Group')!=-1 ):
        if(len(Horizontal_Space_List[i])==0):
            excel_name = str(chr(Vertical_CharPrefix_StartFrom))+str(Vertical_CharPrefix_Num)+'_'+Group_List[0]+'Space' +'.xlsx'
            temp_excel_path  = D1_Core_Excel_Path + "/" + excel_name
            #print(temp_excel_path)
            Chk_AlreadFile = os.path.exists(temp_excel_path) 
            if Chk_AlreadFile is True :
                print('File Existed')
                print(temp_excel_path)
            else :
                with open(os.path.join(D1_Core_Excel_Path, excel_name), 'w') as fp: 
                    Chk_File = os.path.exists(temp_excel_path) 
                    if Chk_File is True :
                        print('File Succesfully Created')
                        print(temp_excel_path)
                    else :
                        print('File Succesfully Not Created')                  
                        print(temp_excel_path)
        elif(len(Horizontal_Space_List[i])>0):
            Horizontal_CharPrefix_StartFrom = 65
            Horizontal_CharPrefix_Num = 1
            for j in range(0,len(Horizontal_Space_List[i])):
                excel_name = str(chr(Vertical_CharPrefix_StartFrom))+str(Vertical_CharPrefix_Num)+'_'+Group_List[0]+'Space'+'_'+str(chr(Horizontal_CharPrefix_StartFrom))+str(Horizontal_CharPrefix_Num)+'_'+Horizontal_Space_List[i][j].replace("_Space",'Space')+'.sql'
                temp_excel_path  = D1_Core_Excel_Path + "/" + excel_name
                #print(temp_excel_path)
                Chk_AlreadFile = os.path.exists(temp_excel_path) 
                if Chk_AlreadFile is True :
                    print('File Existed')
                    print(temp_excel_path)
                else :
                    with open(os.path.join(D1_Core_Excel_Path, excel_name), 'w') as fp: 
                        Chk_File = os.path.exists(temp_excel_path) 
                        if Chk_File is True :
                            print('File Succesfully Created')
                            print(temp_excel_path)
                        else :
                            print('File Succesfully Not Created')                  
                            print(temp_excel_path)
                # --
                Horizontal_CharPrefix_StartFrom = Horizontal_CharPrefix_StartFrom +1
    # --
    # --   
    if(Default_Space_List[i].find('Company')!=-1):
        if(len(Horizontal_Space_List[i])==0):           
            pass
        elif(len(Horizontal_Space_List[i])>0):
            for k in range(0,len(Company_List)):
                Horizontal_CharPrefix_StartFrom = 65
                Horizontal_CharPrefix_Num = 1
                for j in range(0,len(Horizontal_Space_List[i])):             
                    if(Horizontal_Space_List[i][j].find('On_Space')!=-1):
                        excel_name =""
                        #--
                        Company_Serial_Index = 0
                        SysApp_Serial_Index = 0
                        Module_Serial_Index = 0
                        Dimension_Serial_Index = 0
                        Operation_Serial_Index = 0
                        # --
                        for Company_List_Index in range(0,len(Company_List)):    
                            for Company_SysApp_List_Index in range(0,len(Company_SysApp_List[Company_Serial_Index])):    
                                for Company_SysApp_Module_List_Index in range(0,len( Company_SysApp_Module_List[SysApp_Serial_Index])):    
                                    excel_name = str(chr(Vertical_CharPrefix_StartFrom))+str(Vertical_CharPrefix_Num)+'_'+ Company_List[k]+'Space'+'_'+Company_SysApp_Module_List[SysApp_Serial_Index][Company_SysApp_Module_List_Index]+'_'+str(chr(Horizontal_CharPrefix_StartFrom))+str(Horizontal_CharPrefix_Num)+'_'+Horizontal_Space_List[i][j].replace("_Space",'Space')+'.sql'
                                    temp_excel_path  = B1_Script_SQL_Path + "/" + excel_name
                                    #print(temp_excel_path)
                                    Chk_AlreadFile = os.path.exists(temp_excel_path) 
                                    if Chk_AlreadFile is True :
                                        print('File Existed')
                                        print(temp_excel_path)
                                    else :
                                        with open(os.path.join(B1_Script_SQL_Path, excel_name), 'w') as fp: 
                                            Chk_File = os.path.exists(temp_excel_path) 
                                            if Chk_File is True :
                                                print('File Succesfully Created')
                                                print(temp_excel_path)
                                            else :
                                                print('File Succesfully Not Created')                  
                                                print(temp_excel_path)
                                    # --
                        Horizontal_CharPrefix_StartFrom = Horizontal_CharPrefix_StartFrom +1
                    else :
                        excel_name = str(chr(Vertical_CharPrefix_StartFrom))+str(Vertical_CharPrefix_Num)+'_'+ Company_List[k]+'Space'+'_'+str(chr(Horizontal_CharPrefix_StartFrom))+str(Horizontal_CharPrefix_Num)+'_'+Horizontal_Space_List[i][j].replace("_Space",'Space')+'.sql'
                        temp_excel_path  = B1_Script_SQL_Path + "/" + excel_name
                        #print(temp_excel_path)
                        Chk_AlreadFile = os.path.exists(temp_excel_path) 
                        if Chk_AlreadFile is True :
                            print('File Existed')
                            print(temp_excel_path)
                        else :
                            with open(os.path.join(B1_Script_SQL_Path, excel_name), 'w') as fp: 
                                Chk_File = os.path.exists(temp_excel_path) 
                                if Chk_File is True :
                                    print('File Succesfully Created')
                                    print(temp_excel_path)
                                else :
                                    print('File Succesfully Not Created')                  
                                    print(temp_excel_path)
                        # --
                        Horizontal_CharPrefix_StartFrom = Horizontal_CharPrefix_StartFrom +1
    Vertical_CharPrefix_StartFrom = Vertical_CharPrefix_StartFrom+1
# Generator-C1_Data-A1_DBFile_SQLServer
#--
Default_DB_File_Extension_List =['.mdf','.ldf']
Default_DB_File_Inddex_StartFrom = '001'
#-
A1_DBFile_SQLServer_Path = Output_Path + "/" + str(Default_Data_Folder[0][0]) +  "/" + str(Default_Data_Folder[1][0])
#print(A1_DBFile_SQLServer_Path)
# --
for Default_DB_File_Extension_List_Index in range(0,len(Default_DB_File_Extension_List)):   
    excel_name = 'A1_'+Srusti_Start_Name+'_'+'A1_'+Default_DB_File_Inddex_StartFrom+Default_DB_File_Extension_List[Default_DB_File_Extension_List_Index]
    temp_excel_path  = A1_DBFile_SQLServer_Path + "/" + excel_name
    #print(temp_excel_path)
    Chk_AlreadFile = os.path.exists(temp_excel_path) 
    if Chk_AlreadFile is True :
        print('File Existed')
        print(temp_excel_path)
    else :
        with open(os.path.join(A1_DBFile_SQLServer_Path, excel_name), 'w') as fp: 
            Chk_File = os.path.exists(temp_excel_path) 
            if Chk_File is True :
                print('File Succesfully Created')
                print(temp_excel_path)
            else :
                print('File Succesfully Not Created')                  
                print(temp_excel_path)

