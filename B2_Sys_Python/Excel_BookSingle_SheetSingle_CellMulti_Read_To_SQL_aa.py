# aa.py
import os
import xlrd 
from numpy import * 
import numpy as np
# pip install xlrd==1.2.0

# Step-01-Excel Data Extraction
# Step-01-02-Excel File Configuration
ExcelFile_Path =("P:/Professional_Mission/Part_2_2/Fuschia_Research/21.Automation/ExcelToSQL/03.Actual/11.xlsx") 


ExcelFile_Sheet_Index_List =[0,2,3,4]
ExcelFile_SheetXEnd_Index_List =[ [ 0, 1],[ 0, 3],[ 0, 3],[ 0, 5] ]
ExcelFile_SheetYEnd_Index_List =[ [ 0, 3],[ 0,16],[ 0,18],[ 0,16] ] 

# Step-01-02-Excel Data Extraction
# -----------
Database_Name=''
Table_List=[]
Column_List=[]
ColumnValue_List=[]
CURD_List=[]
DataType_List=[]
Constraint_List=[]
ForeignKeyReferenceTable_List= []
ForeignKeyReferenceTableColumn_List=[]
# -----------
Default_Row_List = ['Primary Key','Foreign Key','Not Null','Unique','Identity','Default','Operation','Data Type','Col Name','Col Value']
Default_Row_Index = [2,3,5,6,7,9,11,13,15,16]

# -----------
Excel_Open = xlrd.open_workbook(ExcelFile_Path) 
# -----------
Sheet_Row=[]
Sheet_Col=[]
for x in range(0,len(ExcelFile_Sheet_Index_List)):
    Sheet_Col.append(int(ExcelFile_SheetXEnd_Index_List[x][1])-int(ExcelFile_SheetXEnd_Index_List[x][0]))
    Sheet_Row.append(int(ExcelFile_SheetYEnd_Index_List[x][1])-int(ExcelFile_SheetYEnd_Index_List[x][0]))
    #print(Sheet_Row[Sheet_List_Index])
    #print(Sheet_Col[Sheet_List_Index])
# -----------
for x in range(0,len(ExcelFile_Sheet_Index_List)):
    Sheet_List_Index = ExcelFile_Sheet_Index_List[x]
    # -----------
    TableBlueprint_Structure = Excel_Open.sheet_by_index(Sheet_List_Index) 
    X_Start = 0
    Y_Start = 0    
    X_End = Sheet_Col[x] + 1 
    Y_End = Sheet_Row[x] + 1 # for range increase purppose
    # -----------
    Table_Structure_Matrix=[]
    for y in range(Y_Start,Y_End):
        temp_row =[] 
        for x in range(X_Start,X_End):   
            cell_value = TableBlueprint_Structure.cell(y,x).value
            temp_row.append(cell_value) 
        Table_Structure_Matrix.append(temp_row) 
    #print(np.matrix(Table_Structure_Matrix))
    Table_Structure_NumpyMatrix = np.array(Table_Structure_Matrix)
    # -----------

    # ---
    if Sheet_List_Index == 0 :
        Database_Name=str(Table_Structure_NumpyMatrix[1][0])
        # ---
        for Temp_Sheet_Y_Index in range(1,Y_End):  
            Table_List.append(Table_Structure_NumpyMatrix[Temp_Sheet_Y_Index][1] )
    else :
        # --
        PrimaryKey_Row = 0
        NotNull_Row = 0 
        ForeignKey_Row = 0
        Identity_Row = 0
        Default_Row = 0
        Operation_Row = 0
        DataType_Row = 0
        ColName_Row = 0 
        ColValue_Row = 0
        for i in range(0,len(Default_Row_List)):
            if Default_Row_List[i] == 'Primary Key':               
                PrimaryKey_Row = Default_Row_Index[i]
            if Default_Row_List[i] == 'Not Null':
                NotNull_Row = Default_Row_Index[i]
            if Default_Row_List[i] == 'Foreign Key':
                ForeignKey_Row = Default_Row_Index[i]
            if Default_Row_List[i] == 'Identity':
                Identity_Row = Default_Row_Index[i]
            if Default_Row_List[i] == 'Default':
                Default_Row = Default_Row_Index[i]
            if Default_Row_List[i] == 'Operation':               
                Operation_Row = Default_Row_Index[i]
            if Default_Row_List[i] == 'Data Type':
                DataType_Row = Default_Row_Index[i]
            if Default_Row_List[i] == 'Col Name':
                ColName_Row = Default_Row_Index[i]
            if Default_Row_List[i] == 'Col Value':
                ColValue_Row = Default_Row_Index[i]        
        # --- 
        Temp_Column_List=[]
        Temp_ColumnValue_ForEachTable_List=[]
        Temp_DataType_List=[]
        Temp_CURD_List=[]

        Temp_Constraint_ForEachTable_List=[]       
        Temp_ForeignKeyReferenceTable_ForEachTable_List= []
        Temp_ForeignKeyReferenceTableColumn_ForEachTable_List=[]
        # --- 
        for x in range(X_End):
            # ---
            Temp_DataType_List.append(Table_Structure_NumpyMatrix[DataType_Row+1][x])
            Temp_Column_List.append(Table_Structure_NumpyMatrix[ColName_Row][x])
            # --- 
            # -           
            Temp_Constraint_ForEachCol_List=[]
            Temp_ColumnValue_ForEachCol_List=[]
            # -           
            Temp_CURD_Val=''
            Temp_ForeignKeyReferenceTable_Val=''
            Temp_ForeignKeyReferenceTableColumn_Val=''

            for y in range(Y_Start,Y_End):   
                if str(Table_Structure_NumpyMatrix[y][x])=='Primary Key' :
                    Temp_Constraint_ForEachCol_List.append(Table_Structure_NumpyMatrix[y][x])
                if str(Table_Structure_NumpyMatrix[y][x])=='Foreign Key' :
                    Temp_Constraint_ForEachCol_List.append(Table_Structure_NumpyMatrix[y][x])
                    Temp_Str=str(Table_Structure_NumpyMatrix[y+1][x])
                    Temp_Str_Split = Temp_Str.split('-')
                    Temp_ForeignKeyReferenceTable_Val = Temp_Str_Split[0]
                    Temp_ForeignKeyReferenceTableColumn_Val = Temp_Str_Split[1]
                if str(Table_Structure_NumpyMatrix[y][x])=='Not Null' :
                    Temp_Constraint_ForEachCol_List.append(Table_Structure_NumpyMatrix[y][x])
                if str(Table_Structure_NumpyMatrix[y][x])=='Unique' :
                    Temp_Constraint_ForEachCol_List.append(Table_Structure_NumpyMatrix[y][x])
                if str(Table_Structure_NumpyMatrix[y][x]).startswith('Identity') :
                    Temp_Val=str(Table_Structure_NumpyMatrix[y+1][x]).split('.')[0]
                    Temp_Str = 'Identity('+str(Temp_Val)+',1)'
                    Temp_Constraint_ForEachCol_List.append(Temp_Str)
                if str(Table_Structure_NumpyMatrix[y][x]).startswith('Default') :
                    Temp_Val=str(Table_Structure_NumpyMatrix[y+1][x]).split('.')[0]
                    Temp_Str = str(Table_Structure_NumpyMatrix[y][x])+' '+str(Temp_Val) 
                    Temp_Constraint_ForEachCol_List.append(Temp_Str)
                if str(Table_Structure_NumpyMatrix[y][x])=='Operation' :
                    Temp_CURD_Val=str(Table_Structure_NumpyMatrix[y+1][x]).split('.')[0]
                if y>=ColValue_Row:
                    Temp_ColumnValue_ForEachCol_List.append(Table_Structure_NumpyMatrix[y][x])
  
            # -           
            Temp_Constraint_ForEachTable_List.append(Temp_Constraint_ForEachCol_List)
            Temp_ColumnValue_ForEachTable_List.append(Temp_ColumnValue_ForEachCol_List)

            Temp_ForeignKeyReferenceTable_ForEachTable_List.append(Temp_ForeignKeyReferenceTable_Val)
            Temp_ForeignKeyReferenceTableColumn_ForEachTable_List.append(Temp_ForeignKeyReferenceTableColumn_Val)
            # - 
            Temp_CURD_List.append(Temp_CURD_Val)
        # --- 
        Column_List.append(Temp_Column_List)
        DataType_List.append(Temp_DataType_List)
        CURD_List.append(Temp_CURD_List)

        Constraint_List.append(Temp_Constraint_ForEachTable_List)
        ColumnValue_List.append(Temp_ColumnValue_ForEachTable_List)
        ForeignKeyReferenceTable_List.append(Temp_ForeignKeyReferenceTable_ForEachTable_List)
        ForeignKeyReferenceTableColumn_List.append(Temp_ForeignKeyReferenceTableColumn_ForEachTable_List)
    # -----------
# -----------
print('Database_Name = '+"'"+str(Database_Name)+"'")
print('Table_List = '+str(Table_List))
print('Column_List = '+str(Column_List))
print('ColumnValue_List = '+str(ColumnValue_List))
print('CURD_List = '+str(CURD_List))
print('DataType_List = '+str(DataType_List))
print('Constraint_List = '+str(Constraint_List))
print('ForeignKeyReferenceTable_List = '+str(ForeignKeyReferenceTable_List))
print('ForeignKeyReferenceTableColumn_List = '+str(ForeignKeyReferenceTableColumn_List))