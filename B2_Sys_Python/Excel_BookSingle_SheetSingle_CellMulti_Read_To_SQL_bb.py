# bb.py
# Step-01-02-Excel Data Extraction
# From aa.py
Database_Name = 'SahuGroup_Dtbs_001'
Table_List = ['IndexSpace_Database_List', 'IndexSpace_Table_List']
Column_List = [['Database_Name', 'Database_ID', 'Record_DateTime', 'Upgrade_DateTime'], ['Database_ID', 'Table_Name', 'Table_ID', 'Record_DateTime', 'Upgrade_DateTime']]
ColumnValue_List = [[['SahuGroup_Dtbs_001'], ['001'], ['2019-02-23 20:02:21.550'], ['2019-02-23 20:02:21.550']], [['001'], ['IndexSpace_Database_List'], ['001'], ['2019-02-23 20:02:21.550'], ['2019-02-23 20:02:21.550']]]
CURD_List = [['', 'URD', '', ''], ['', '', 'URD', '', '']]
DataType_List = [['varchar(40) ', 'int ', 'datetime', 'datetime'], ['int ', 'varchar(80) ', 'int ', 'datetime', 'datetime']]
Constraint_List = [[['Not Null', 'Unique'], ['Primary Key', 'Identity(001,1)'], ['Default Current_Timestamp'], []], [['Foreign Key', 'Not Null'], ['Not Null', 'Unique'], ['Primary Key', 'Identity(001,1)'], ['Default Current_Timestamp'], []]]
ForeignKeyReferenceTable_List = [['', '', '', ''], ['IndexSpace_Database_List', '', '', '', '']]
ForeignKeyReferenceTableColumn_List = [['', '', '', ''], ['Database_ID', '', '', '', '']]

"""
DataType	DataValue	
Excel	SQL	Excel	SQL
int	int	Int_Value	100
varchar(40)	varchar(40)	Str_Value	aaa		
"""
DataType_SQL_List=['char(n)','varchar(n)','nchar','nvarchar','nvarchar(max)','ntext','binary(n)','varbinary','varbinary(max)','text','image','bit','tinyint','smallint','int','bigint','decimal(p,s)','numeric(p,s)','smallmoney','money','float(n)','real','date','datetime','datetime2','smalldatetime','time','datetimeoffset','timestamp']
DataType_Excel_List=['char(n)','varchar(n)','nchar','nvarchar','nvarchar(max)','ntext','binary(n)','varbinary','varbinary(max)','text','image','bit','tinyint','smallint','int','bigint','decimal(p,s)','numeric(p,s)','smallmoney','money','float(n)','real','date','datetime','datetime2','smalldatetime','time','datetimeoffset','timestamp']

DataTValue_SQL_List=['char(n)','Str_Val','nchar','nvarchar','nvarchar(max)','ntext','binary(n)','varbinary','varbinary(max)','text','image','bit','tinyint','smallint','int','bigint','decimal(p,s)','numeric(p,s)','smallmoney','money','float(n)','real','date','datetime','datetime2','smalldatetime','time','datetimeoffset','timestamp']
DataValue_Excel_List=['char(n)','aaa','nchar','nvarchar','nvarchar(max)','ntext','binary(n)','varbinary','varbinary(max)','text','image','bit','tinyint','smallint',100,'bigint','decimal(p,s)','numeric(p,s)','smallmoney','money','float(n)','real','date','Current_Timestamp','datetime2','smalldatetime','time','datetimeoffset','timestamp']


Default_Int_Val = 100
Default_Str_Val = 'aaa'

UpdateSearch_Col = 'Column_XX'
UpdateSearch_Col_DType = 'int'
UpdateSearch_Col_Val = '100'

ReadSearch_Col = 'Column_Xx'
ReadSearch_Col_DType = 'int'
ReadSearch_Col_Val = '100'

DeleteSearch_Col = 'Column_XX'
DeleteSearch_Col_DType = 'int'
DeleteSearch_Col_Val = '100'
#-
# Step-02-Excel Data Scriptation
# Step-02-01-Script File Ready
ScriptFile_Path = 'P:/Professional_Mission/Part_2_2/Fuschia_Research/Automation_SQL/03.Actual/22.sql'
"""
ExcelFile_Path_Split = ExcelFile_Path.split('/')
ScriptFile_Path =''
for i in range(0,len(ExcelFile_Path_Split)-1):
    if i ==len(ExcelFile_Path_Split)-2 :
        ScriptFile_Path = ScriptFile_Path + ExcelFile_Path_Split[i]  
    else :
        ScriptFile_Path = ScriptFile_Path + ExcelFile_Path_Split[i] +'/' 
ScriptFile_Name ='xxx.sql'
with open(ScriptFile_Path+'/'+ScriptFile_Name, 'w') as fp: 
    pass
"""
# Step-02-02-Script File Operation
ScriptFile = open(ScriptFile_Path,"w") 
ScriptFileComment = '--' 

# Step-02-03-Database Comment Line Work
Script_Database_CommentLine_List = ['Database-xxx','Operation']
Script_Database_CommentLineScript_List = []
for Temp_CommentLine in Script_Database_CommentLine_List :
    # print(Temp_CommentLine)
    if Temp_CommentLine=='Database-xxx':
        # Prepare String
        Temp_CommentLine_Script=''
        Temp_Str_01 = ScriptFileComment+' '+Temp_CommentLine.replace('xxx',str(Database_Name))+'\n'
        Temp_CommentLine_Script = Temp_CommentLine_Script + Temp_Str_01 
        # print(Temp_CommentLine_Script)
        # Write To File
        ScriptFile = open(ScriptFile_Path,"a") 
        ScriptFile.writelines(Temp_CommentLine_Script)
        ScriptFile.close()
        # Read From File
    elif Temp_CommentLine=='Operation':
        # Prepare String
        Temp_CommentLine_Script=''
        Temp_Str_01 = ScriptFileComment+' '+Temp_CommentLine+'\n'
        Temp_Str_02 = 'use'+' '+str(Database_Name)+'\n'
        Temp_Str_03 =  'GO'+'\n'
        Temp_CommentLine_Script = Temp_CommentLine_Script + Temp_Str_01 + Temp_Str_02 + Temp_Str_03
        # print(Temp_CommentLine_Script)
        # Write To File
        ScriptFile = open(ScriptFile_Path,"a") 
        ScriptFile.writelines(Temp_CommentLine_Script)
        ScriptFile.close()
        # Read From File
        
# Step-02-03-Table Comment Line Work
#Script_Table_CommentLine_List = ['Table-xxx-Start','Operation-Query-CURD','Create','Update-All','Read-All','Operation-Stored Procedure','Update-All','Update-Search','Read-All','Read-Search','Delete-Search','Table-xxx-End']
Script_Table_CommentLine_List = ['Table-xxx-Start','Operation-Query-CURD','Create','Update-All','Read-All','Table-xxx-End']

Script_Table_CommentLineScript_List = []

for Table_List_Index in range(0,len(Table_List)):   
    Temp_Table = Table_List[Table_List_Index]
    Is_CURD_Done = 0

    # Pre Work
    # print(CURD_List[Table_List_Index])
    for CURD_List_Index in range(0,len(CURD_List[Table_List_Index])):
        if CURD_List[Table_List_Index][CURD_List_Index]=='URD' or CURD_List[Table_List_Index][CURD_List_Index]=='CURD':
            UpdateSearch_Col = Column_List[Table_List_Index][CURD_List_Index]
            UpdateSearch_Col_DType = DataType_List[Table_List_Index][CURD_List_Index]
            UpdateSearch_Col_Val = ColumnValue_List[Table_List_Index][CURD_List_Index]

            ReadSearch_Col = Column_List[Table_List_Index][CURD_List_Index]
            ReadSearch_Col_DType = DataType_List[Table_List_Index][CURD_List_Index]
            ReadSearch_Col_Val = ColumnValue_List[Table_List_Index][CURD_List_Index]

            DeleteSearch_Col = Column_List[Table_List_Index][CURD_List_Index]
            DeleteSearch_Col_DType = DataType_List[Table_List_Index][CURD_List_Index]
            DeleteSearch_Col_Val = ColumnValue_List[Table_List_Index][CURD_List_Index]
    # --
    for Temp_CommentLine_Index in range(0,len(Script_Table_CommentLine_List)) :
        Temp_CommentLine = Script_Table_CommentLine_List[Temp_CommentLine_Index]
        # print(Temp_CommentLine)
        if Temp_CommentLine=='Table-xxx-Start':
            # Prepare String
            Temp_CommentLine_Script=''
            Temp_Str_01 = ScriptFileComment+' '+Temp_CommentLine.replace('xxx',str(Temp_Table))+'\n'
            Temp_CommentLine_Script = Temp_CommentLine_Script + Temp_Str_01 
            # print(Temp_CommentLine_Script)
            # Write To File
            ScriptFile = open(ScriptFile_Path,"a") 
            ScriptFile.writelines(Temp_CommentLine_Script)
            ScriptFile.close()
            # Read From File
        elif Temp_CommentLine=='Operation-Query-CURD':
            # Prepare String
            Temp_CommentLine_Script=''
            Temp_Str_01 = ScriptFileComment+' '+Temp_CommentLine+'\n'
            Temp_CommentLine_Script = Temp_CommentLine_Script + Temp_Str_01 
            # print(Temp_CommentLine_Script)
            # Write To File
            ScriptFile = open(ScriptFile_Path,"a") 
            ScriptFile.writelines(Temp_CommentLine_Script)
            ScriptFile.close()
            # Read From File
        elif Temp_CommentLine=='Create' and Is_CURD_Done==0:
            # Prepare String
            Temp_CommentLine_Script=''
            Temp_Str_01 = ScriptFileComment+' '+Temp_CommentLine+'\n'
            Temp_Str_02 = 'Create Table'+' '+Temp_Table+'\n'
            Temp_Str_03 = '(' + '\n'
            Temp_Str_04 = ''
            Add_End_Constraint=''
            Col_Constraint_Type = ''
            for Column_List_Index in range(0,len(Column_List[Table_List_Index])):
                Col_Name = Column_List[Table_List_Index][Column_List_Index]
                Col_Data_Type = DataType_List[Table_List_Index][Column_List_Index]
                Col_Constraint_Type = ''
                for Constarint_Index_List in range(0,len(Constraint_List[Table_List_Index][Column_List_Index])):
                    if  Constraint_List[Table_List_Index][Column_List_Index][Constarint_Index_List]=='Foreign Key':
                        Col_Constraint_Type = Col_Constraint_Type+' '
                        if Add_End_Constraint!='':
                            Add_End_Constraint = Add_End_Constraint +'	'+',Constraint'+' '+'FP_'+Col_Name+'_'+Temp_Table+' '+'Foreign Key'+'('+Col_Name+')'+' References'+' '+str(ForeignKeyReferenceTable_List[Table_List_Index][Column_List_Index])+'('+str(ForeignKeyReferenceTableColumn_List[Table_List_Index][Column_List_Index])+')'+'\n'
                        else : 
                            Add_End_Constraint = Add_End_Constraint +'	'+'Constraint'+' '+'FP_'+Col_Name+'_'+Temp_Table+' '+'Foreign Key'+'('+Col_Name+')'+' References'+' '+str(ForeignKeyReferenceTable_List[Table_List_Index][Column_List_Index])+'('+str(ForeignKeyReferenceTableColumn_List[Table_List_Index][Column_List_Index])+')'+'\n'
                    else: 
                        Col_Constraint_Type = Col_Constraint_Type+Constraint_List[Table_List_Index][Column_List_Index][Constarint_Index_List]+' '
                if Column_List_Index==len(Column_List[Table_List_Index])-1:
                    if Add_End_Constraint=='':
                        Temp_Str_04 = Temp_Str_04 + '	'+Col_Name+' '+Col_Data_Type+' '+Col_Constraint_Type+'\n'
                    else :
                        Temp_Str_04 = Temp_Str_04 + '	'+Col_Name+' '+Col_Data_Type+' '+Col_Constraint_Type+','+'\n'
                else :
                    Temp_Str_04 = Temp_Str_04 + '	'+Col_Name+' '+Col_Data_Type+' '+Col_Constraint_Type+','+'\n'
            Temp_Str_05 = ')' + '\n'
            if Add_End_Constraint=='':
                Temp_CommentLine_Script = Temp_CommentLine_Script + Temp_Str_01 + Temp_Str_02 + Temp_Str_03 + Temp_Str_04 + Temp_Str_05
            else :
                Temp_CommentLine_Script = Temp_CommentLine_Script + Temp_Str_01 + Temp_Str_02 + Temp_Str_03 + Temp_Str_04 + Add_End_Constraint+Temp_Str_05
            Temp_Str_06 = 'GO'+'\n'
            Temp_CommentLine_Script = Temp_CommentLine_Script +  Temp_Str_06  
            # print(Temp_CommentLine_Script)
            # Write To File
            ScriptFile = open(ScriptFile_Path,"a") 
            ScriptFile.writelines(Temp_CommentLine_Script)
            ScriptFile.close()
            # Read From File
        elif Temp_CommentLine=='Update-All'and Is_CURD_Done==0:
            # Prepare String
            Temp_CommentLine_Script=''
            Temp_Str_01 = ScriptFileComment+' '+Temp_CommentLine+'\n'
            Temp_Str_02 = 'Insert Into '+Temp_Table+'\n'
            Temp_Str_03 = '('

            Temp_Valid_Col_List=[]
            for Column_List_Index in range(0,len(Column_List[Table_List_Index])):
                Is_Identity_Default = False
                for i in range(0,len(Constraint_List[Table_List_Index][Column_List_Index])):
                    if Constraint_List[Table_List_Index][Column_List_Index][i].startswith('Identity') or Constraint_List[Table_List_Index][Column_List_Index][i].startswith('Default'):
                        Is_Identity_Default = True
                if(Is_Identity_Default==False) :
                    Temp_Valid_Col_List.append(Column_List_Index)

            for Temp_Valid_Col_List_Index in range(0,len(Temp_Valid_Col_List)):
                Temp_Valid_Col_Index = Temp_Valid_Col_List[Temp_Valid_Col_List_Index]
                if(Temp_Valid_Col_List_Index==len(Temp_Valid_Col_List)-1):
                    Temp_Str_03=Temp_Str_03+Column_List[Table_List_Index][Temp_Valid_Col_Index]                        
                else :
                    Temp_Str_03=Temp_Str_03+Column_List[Table_List_Index][Temp_Valid_Col_Index]+','                        


            Temp_Str_03 = Temp_Str_03 +')'+'\n'
            Temp_Str_04 = 'values\n'
            Temp_Str_05 = ''

            for ColumnValue_List_Row_Index in range(0,len(ColumnValue_List[Table_List_Index][0])):
                Temp_Str_05 = Temp_Str_05 + '(\n'
                # --
                Temp_Valid_Col_List=[]
                for Column_List_Index in range(0,len(Column_List[Table_List_Index])):
                    Is_Identity = False
                    for i in range(0,len(Constraint_List[Table_List_Index][Column_List_Index])):
                        if Constraint_List[Table_List_Index][Column_List_Index][i].startswith('Identity') or Constraint_List[Table_List_Index][Column_List_Index][i].startswith('Default'):
                            Is_Identity = True
                    if(Is_Identity==False) :
                        Temp_Valid_Col_List.append(Column_List_Index)
                #print(Temp_Valid_Col_List)
                for x in range(0,len(Temp_Valid_Col_List)):
                    Temp_Valid_Col_List_Index = Temp_Valid_Col_List[x]
                    if(x==(len(Temp_Valid_Col_List)-1)):
                        # Insert with out Comma
                        if(ColumnValue_List[Table_List_Index][Temp_Valid_Col_List_Index][ColumnValue_List_Row_Index]=='Str_Val'):
                            Temp_Str_05=Temp_Str_05+"'"+Default_Str_Val+"'"       
                        elif(ColumnValue_List[Table_List_Index][Temp_Valid_Col_List_Index][ColumnValue_List_Row_Index]=='Int_Val'):
                            Temp_Str_05=Temp_Str_05+ str(Default_Int_Val)
                        elif ColumnValue_List[Table_List_Index][Temp_Valid_Col_List_Index][ColumnValue_List_Row_Index].isnumeric()==True :
                            Temp_Str_05=Temp_Str_05+ str(ColumnValue_List[Table_List_Index][Temp_Valid_Col_List_Index][ColumnValue_List_Row_Index])                           
                        elif ColumnValue_List[Table_List_Index][Temp_Valid_Col_List_Index][ColumnValue_List_Row_Index].isnumeric()==False :
                            Temp_Str_05=Temp_Str_05+"'"+str(ColumnValue_List[Table_List_Index][Temp_Valid_Col_List_Index][ColumnValue_List_Row_Index])+"'"       
                    else :
                        # Insert with Comma
                        if(ColumnValue_List[Table_List_Index][Temp_Valid_Col_List_Index][ColumnValue_List_Row_Index]=='Str_Val'):
                            Temp_Str_05=Temp_Str_05+"'"+Default_Str_Val+"'"+','       
                        elif(ColumnValue_List[Table_List_Index][Temp_Valid_Col_List_Index][ColumnValue_List_Row_Index]=='Int_Val'):
                            Temp_Str_05=Temp_Str_05+ str(Default_Int_Val)+','                           
                        else :
                            Temp_Str_05=Temp_Str_05+"'"+str(ColumnValue_List[Table_List_Index][Temp_Valid_Col_List_Index][ColumnValue_List_Row_Index])+"'"+','       
                # --
                if ColumnValue_List_Row_Index == len(ColumnValue_List[Table_List_Index][0])-1:
                    Temp_Str_05 = Temp_Str_05 + '\n)\n'
                else :
                    Temp_Str_05 = Temp_Str_05 + '\n),\n'

            Temp_Str_06 = 'GO'+'\n'
            Temp_CommentLine_Script = Temp_CommentLine_Script + Temp_Str_01 + Temp_Str_02 + Temp_Str_03 + Temp_Str_04 + Temp_Str_05 + Temp_Str_06
            # print(Temp_CommentLine_Script)
            # Write To File
            ScriptFile = open(ScriptFile_Path,"a") 
            ScriptFile.writelines(Temp_CommentLine_Script)
            ScriptFile.close()
            # Read From File
        elif Temp_CommentLine=='Read-All'and Is_CURD_Done==0:
            # Prepare String
            Temp_CommentLine_Script=''
            Temp_Str_01 = ScriptFileComment+' '+Temp_CommentLine+'\n'
            Temp_Str_02 = 'Select * From '+Temp_Table+'\n'
            Temp_Str_03 = 'GO'+'\n'
            Temp_CommentLine_Script = Temp_CommentLine_Script + Temp_Str_01 + Temp_Str_02 + Temp_Str_03
            # print(Temp_CommentLine_Script)
            # Write To File
            ScriptFile = open(ScriptFile_Path,"a") 
            ScriptFile.writelines(Temp_CommentLine_Script)
            ScriptFile.close()
            # Read From File
        elif Temp_CommentLine=='Operation-Stored Procedure':
            # Prepare String
            Temp_CommentLine_Script=''
            Temp_Str_01 = ScriptFileComment+' '+Temp_CommentLine+'\n'
            Temp_CommentLine_Script = Temp_CommentLine_Script + Temp_Str_01 
            # print(Temp_CommentLine_Script)
            # Write To File
            ScriptFile = open(ScriptFile_Path,"a") 
            ScriptFile.writelines(Temp_CommentLine_Script)
            ScriptFile.close()
            # Read From File
            # Global Var Operation
            Is_CURD_Done = 1
        elif Temp_CommentLine=='Update-All' and Is_CURD_Done==1:
            # Prepare String
            Temp_CommentLine_Script=''
            Temp_Str_01 = ScriptFileComment+' '+Temp_CommentLine+'\n'
            Temp_Str_02 = 'Create Procedure'+' '+'SP_UpdateAll_'+Temp_Table+'\n'
            Temp_Str_03 = '(' + '\n'
            Temp_Str_04 = ''

            Temp_Valid_Col_List=[]
            for Column_List_Index in range(0,len(Column_List[Table_List_Index])):
                Is_Identity = False
                for i in range(0,len(Constraint_List[Table_List_Index][Column_List_Index])):
                    if Constraint_List[Table_List_Index][Column_List_Index][i].startswith('Identity') or Constraint_List[Table_List_Index][Column_List_Index][i].startswith('Default'):
                        Is_Identity = True
                if(Is_Identity==False) :
                    Temp_Valid_Col_List.append(Column_List_Index)

            for Temp_Valid_Col_List_Index in range(0,len(Temp_Valid_Col_List)):
                Temp_Valid_Col_Index = Temp_Valid_Col_List[Temp_Valid_Col_List_Index]
                Col_Name = Column_List[Table_List_Index] [Temp_Valid_Col_Index]
                Col_Data_Type =DataType_List[Table_List_Index] [Temp_Valid_Col_Index]

                if(Temp_Valid_Col_List_Index==len(Temp_Valid_Col_List)-1):
                    Temp_Str_04 = Temp_Str_04 + '	'+'@Var_'+Col_Name+' '+Col_Data_Type+'\n'
                else :
                    Temp_Str_04 = Temp_Str_04 + '	'+'@Var_'+Col_Name+' '+Col_Data_Type+','+'\n'


            Temp_Str_05 = ')' + '\n'
            Temp_Str_06 = 'As' + '\n'
            Temp_Str_07 = 'Begin' + '\n'
            Temp_Str_08 = ''
            #
            # Prepare String
            Temp_Str_08_01 = '	'+'Insert Into '+Temp_Table+' ('
            Temp_Str_08_02 =''

            Temp_Valid_Col_List=[]
            for Column_List_Index in range(0,len(Column_List[Table_List_Index])):
                Is_Identity = False
                for i in range(0,len(Constraint_List[Table_List_Index][Column_List_Index])):
                    if Constraint_List[Table_List_Index][Column_List_Index][i].startswith('Identity') or Constraint_List[Table_List_Index][Column_List_Index][i].startswith('Default'):
                        Is_Identity = True
                if(Is_Identity==False) :
                    Temp_Valid_Col_List.append(Column_List_Index)

            for Temp_Valid_Col_List_Index in range(0,len(Temp_Valid_Col_List)):
                Temp_Valid_Col_Index = Temp_Valid_Col_List[Temp_Valid_Col_List_Index]
                Col_Name = Column_List[Table_List_Index] [Temp_Valid_Col_Index]
                Col_Data_Type =DataType_List[Table_List_Index] [Temp_Valid_Col_Index]
                if(Temp_Valid_Col_List_Index==len(Temp_Valid_Col_List)-1):
                    Temp_Str_08_02=Temp_Str_08_02+Column_List[Table_List_Index][Temp_Valid_Col_Index]
                else :
                    Temp_Str_08_02=Temp_Str_08_02+Column_List[Table_List_Index][Temp_Valid_Col_Index]+','



            Temp_Str_08_02 = Temp_Str_08_02 +')'+' '
            Temp_Str_08_03 = 'values('

            Temp_Valid_Col_List=[]
            for Column_List_Index in range(0,len(Column_List[Table_List_Index])):
                Is_Identity = False
                for i in range(0,len(Constraint_List[Table_List_Index][Column_List_Index])):
                    if Constraint_List[Table_List_Index][Column_List_Index][i].startswith('Identity') or Constraint_List[Table_List_Index][Column_List_Index][i].startswith('Default'):
                        Is_Identity = True
                if(Is_Identity==False) :
                    Temp_Valid_Col_List.append(Column_List_Index)

            for Temp_Valid_Col_List_Index in range(0,len(Temp_Valid_Col_List)):
                Temp_Valid_Col_Index = Temp_Valid_Col_List[Temp_Valid_Col_List_Index]

                Col_Name = Column_List[Table_List_Index] [Temp_Valid_Col_Index]
                Col_Data_Type =DataType_List[Table_List_Index] [Temp_Valid_Col_Index]
                if(Temp_Valid_Col_List_Index==len(Temp_Valid_Col_List)-1):
                    Temp_Str_08_03 = Temp_Str_08_03 +'@Var_'+Col_Name
                else :
                    Temp_Str_08_03 = Temp_Str_08_03 +'@Var_'+Col_Name+','


            Temp_Str_08_03 = Temp_Str_08_03 +')'+'\n'
            #
            Temp_Str_08 = Temp_Str_08_01 + Temp_Str_08_02 + Temp_Str_08_03
            Temp_Str_09 = 'End' + '\n'
            Temp_Str_10 = 'GO'+'\n'
            Temp_CommentLine_Script = Temp_CommentLine_Script + Temp_Str_01 + Temp_Str_02 + Temp_Str_03 + Temp_Str_04 + Temp_Str_05 + Temp_Str_06 + Temp_Str_07+Temp_Str_08+Temp_Str_09+Temp_Str_10
            # print(Temp_CommentLine_Script)
            # Write To File
            ScriptFile = open(ScriptFile_Path,"a") 
            ScriptFile.writelines(Temp_CommentLine_Script)
            ScriptFile.close()
            # Read From File
        elif Temp_CommentLine=='Update-Search' and Is_CURD_Done==1:
            # Prepare String
            Temp_CommentLine_Script=''
            Temp_Str_01 = ScriptFileComment+' '+Temp_CommentLine+'\n'
            Temp_Str_02 = 'Create Procedure'+' '+'SP_UpdateSearch_'+Temp_Table+'\n'
            Temp_Str_03 = '(' + '\n'
            Temp_Str_04 = ''

            Temp_Valid_Col_List=[]
            for Column_List_Index in range(0,len(Column_List[Table_List_Index])):
                Is_Identity = False
                for i in range(0,len(Constraint_List[Table_List_Index][Column_List_Index])):
                    if Constraint_List[Table_List_Index][Column_List_Index][i].startswith('IdentityNoNeed'):
                        Is_Identity = True
                if(Is_Identity==False) :
                    Temp_Valid_Col_List.append(Column_List_Index)

            for Temp_Valid_Col_List_Index in range(0,len(Temp_Valid_Col_List)):
                Temp_Valid_Col_Index = Temp_Valid_Col_List[Temp_Valid_Col_List_Index]

                Col_Name = Column_List[Table_List_Index] [Temp_Valid_Col_Index]
                Col_Data_Type =DataType_List[Table_List_Index] [Temp_Valid_Col_Index]
                if(Temp_Valid_Col_List_Index==len(Temp_Valid_Col_List)-1):
                    Temp_Str_04 = Temp_Str_04 + '	'+'@Var_'+Col_Name+' '+Col_Data_Type+'\n'
                else :
                    Temp_Str_04 = Temp_Str_04 + '	'+'@Var_'+Col_Name+' '+Col_Data_Type+','+'\n'

            Temp_Str_05 = ')' + '\n'
            Temp_Str_06 = 'As' + '\n'
            Temp_Str_07 = 'Begin' + '\n'
            Temp_Str_08 = ''
            #
            # Prepare String
            Temp_Str_08_01 = '	'+'Update '+Temp_Table+' '
            Temp_Str_08_02 ='Set '

            Temp_Valid_Col_List=[]
            for Column_List_Index in range(0,len(Column_List[Table_List_Index])):
                Is_Identity = False
                for i in range(0,len(Constraint_List[Table_List_Index][Column_List_Index])):
                    if Constraint_List[Table_List_Index][Column_List_Index][i].startswith('Identity'):
                        Is_Identity = True
                if(Is_Identity==False) :
                    Temp_Valid_Col_List.append(Column_List_Index)

            for Temp_Valid_Col_List_Index in range(0,len(Temp_Valid_Col_List)):
                Temp_Valid_Col_Index = Temp_Valid_Col_List[Temp_Valid_Col_List_Index]

                Col_Name = Column_List[Table_List_Index] [Temp_Valid_Col_Index]
                Col_Data_Type =DataType_List[Table_List_Index] [Temp_Valid_Col_Index]
                if(Temp_Valid_Col_List_Index==len(Temp_Valid_Col_List)-1):
                    Temp_Str_08_02=Temp_Str_08_02+ ' '+Column_List[Table_List_Index][Temp_Valid_Col_Index]+'='+'@Var_'+Column_List[Table_List_Index][Temp_Valid_Col_Index]
                else :
                    Temp_Str_08_02=Temp_Str_08_02+ ' '+Column_List[Table_List_Index][Temp_Valid_Col_Index]+'='+'@Var_'+Column_List[Table_List_Index][Temp_Valid_Col_Index]+','


            Temp_Str_08_02 = Temp_Str_08_02+'\n'
            Temp_Str_08_03 = '	'+'where '+UpdateSearch_Col+'='+'@Var_'+UpdateSearch_Col+'\n'
            Temp_Str_08 = Temp_Str_08_01 + Temp_Str_08_02 + Temp_Str_08_03
            Temp_Str_09 = 'End' + '\n'
            Temp_Str_10 = 'GO'+'\n'
            Temp_CommentLine_Script = Temp_CommentLine_Script + Temp_Str_01 + Temp_Str_02 + Temp_Str_03 + Temp_Str_04 + Temp_Str_05 + Temp_Str_06 + Temp_Str_07+Temp_Str_08+Temp_Str_09+Temp_Str_10
            # print(Temp_CommentLine_Script)
            # Write To File
            ScriptFile = open(ScriptFile_Path,"a") 
            ScriptFile.writelines(Temp_CommentLine_Script)
            ScriptFile.close()
            # Read From File
        elif Temp_CommentLine=='Read-All' and Is_CURD_Done==1:
            # Prepare String
            Temp_CommentLine_Script=''
            Temp_Str_01 = ScriptFileComment+' '+Temp_CommentLine+'\n'
            Temp_Str_02 = 'Create Procedure'+' '+'SP_ReadAll_'+Temp_Table+'\n'
            Temp_Str_03 = 'As' + '\n'
            Temp_Str_04 = 'Begin' + '\n'
            Temp_Str_05 = '	'+'Select * from '+Temp_Table+'\n'
            Temp_Str_06 = 'End' + '\n'
            Temp_Str_07 = 'GO'+'\n'
            Temp_CommentLine_Script = Temp_CommentLine_Script + Temp_Str_01 + Temp_Str_02 + Temp_Str_03 + Temp_Str_04 + Temp_Str_05 + Temp_Str_06 + Temp_Str_07
            # print(Temp_CommentLine_Script)
            # Write To File
            ScriptFile = open(ScriptFile_Path,"a") 
            ScriptFile.writelines(Temp_CommentLine_Script)
            ScriptFile.close()
            # Read From File
        elif Temp_CommentLine=='Read-Search' and Is_CURD_Done==1:
            # Prepare String
            Temp_CommentLine_Script=''
            Temp_Str_01 = ScriptFileComment+' '+Temp_CommentLine+'\n'
            Temp_Str_02 = 'Create Procedure'+' '+'SP_ReadSearch_'+Temp_Table+'\n'
            Temp_Str_11 = '('+'\n'
            Temp_Str_03 = '	'+'@Var_'+ReadSearch_Col+' '+ReadSearch_Col_DType+'\n' 
            Temp_Str_05 = ')'+'\n'
            Temp_Str_06 = 'As' + '\n'
            Temp_Str_07 = 'Begin' + '\n'
            Temp_Str_08 = '	'+'Select * from '+Temp_Table+' '+'where '+ReadSearch_Col+'='+'@Var_'+ReadSearch_Col+' order by '+ReadSearch_Col+'\n'
            Temp_Str_09 = 'End' + '\n'
            Temp_Str_10 = 'GO'+'\n'
            Temp_CommentLine_Script = Temp_CommentLine_Script + Temp_Str_01 + Temp_Str_02 + Temp_Str_11 + Temp_Str_03 +  Temp_Str_05 + Temp_Str_06 + Temp_Str_07+Temp_Str_08+Temp_Str_09+Temp_Str_10
            # print(Temp_CommentLine_Script)
            # Write To File
            ScriptFile = open(ScriptFile_Path,"a") 
            ScriptFile.writelines(Temp_CommentLine_Script)
            ScriptFile.close()
            # Read From File
        elif Temp_CommentLine=='Delete-Search' and Is_CURD_Done==1:
            # Prepare String
            Temp_CommentLine_Script=''
            Temp_Str_01 = ScriptFileComment+' '+Temp_CommentLine+'\n'
            Temp_Str_02 = 'Create Procedure'+' '+'SP_DeleteSearch_'+Temp_Table+'\n'
            Temp_Str_11 = '('+'\n'
            Temp_Str_03 = '	'+'@Var_'+DeleteSearch_Col+' '+DeleteSearch_Col_DType +'\n' 
            Temp_Str_05 = ')'+'\n'
            Temp_Str_06 = 'As' + '\n'
            Temp_Str_07 = 'Begin' + '\n'
            Temp_Str_08 = '	'+'Delete from '+Temp_Table+' '+'where'+' '+DeleteSearch_Col+'='+'@Var_'+DeleteSearch_Col+'\n'
            Temp_Str_09 = 'End' + '\n'
            Temp_Str_10 = 'GO'+'\n'
            Temp_CommentLine_Script = Temp_CommentLine_Script + Temp_Str_01 + Temp_Str_02 + Temp_Str_11 + Temp_Str_03 +  Temp_Str_05 + Temp_Str_06 + Temp_Str_07+Temp_Str_08+Temp_Str_09+Temp_Str_10
            # print(Temp_CommentLine_Script)
            # Write To File
            ScriptFile = open(ScriptFile_Path,"a") 
            ScriptFile.writelines(Temp_CommentLine_Script)
            ScriptFile.close()
            # Read From File
        elif Temp_CommentLine=='Table-xxx-End':
            # Prepare String
            Temp_CommentLine_Script=''
            Temp_Str_01 = ScriptFileComment+' '+Temp_CommentLine.replace('xxx',str(Temp_Table))+'\n\n'
            Temp_CommentLine_Script = Temp_CommentLine_Script + Temp_Str_01 
            # print(Temp_CommentLine_Script)
            # Write To File
            ScriptFile = open(ScriptFile_Path,"a") 
            ScriptFile.writelines(Temp_CommentLine_Script)
            ScriptFile.close()
            # Read From File
# Step-02-03-Script File Operation
# Read From File
ScriptFile = open(ScriptFile_Path,"r+") 
ScriptFile.seek(0)
ScriptFile.read()
with open(ScriptFile_Path) as file:
    for line in file:
        pass
        # print(line)
ScriptFile.close()
