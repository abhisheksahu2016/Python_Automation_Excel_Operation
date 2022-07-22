
# ----  cc.py
# Step-01-01-Excel Data Extraction
# From aa.py
"""
Database_Name='Dtbs_xxxxxxxxx'
Table_List=['UserSpace_User_Category_List']
Column_List=[['UserCategory_Name', 'UserCategory_ID']]
ColumnValue_List=[[['Str_Val', 'Str_Val', 'Str_Val'], ['Int_Val', 'Int_Val', 'Int_Val']]]
CURD_List= [['', 'URD']]
DataType_List = [['varchar(40) ', 'int ']]
Constraint_List=[[['Not Null'], ['Identity(100000000,1)']]]
ForeignKeyReferenceTable_List=[['',''],['','']]
ForeignKeyReferenceTableColumn_List=[ ['',''], ['','']  ]
"""
# From bb.py
"""
Default_Int_Val = 100
Default_Str_Val = 'aaa'

UpdateSearch_Col = 'Column_XX'
UpdateSearch_Col_DType = 'int'
UpdateSearch_Col_Val = '100'

ReadSearch_Col = 'Column_XX'
ReadSearch_Col_DType = 'int'
ReadSearch_Col_Val = '100'

DeleteSearch_Col = 'Column_XX'
DeleteSearch_Col_DType = 'int'
DeleteSearch_Col_Val = '100'
"""
#-
# Step-01-02-Excel Data To Code Data Replacementation
xxx_Namespace = 'xxx_Project'
xxx_Database = 'xxx'
xxx_SQLConStr = "data source=DESKTOP-C5BSL0D\\\SQL2014;database="+str(Database_Name)+";integrated security = SSPI;"

# Step-02-01-Code File Ready
# --
"""
ExcelFile_Path_Split = ExcelFile_Path.split('/')
CodeFile_Path =''
for i in range(0,len(ExcelFile_Path_Split)-1):
    if i ==len(ExcelFile_Path_Split)-2 :
        CodeFile_Path = CodeFile_Path + ExcelFile_Path_Split[i]  
    else :
        CodeFile_Path = CodeFile_Path + ExcelFile_Path_Split[i] +'/' 
CodeFile_Name ='yyy.cs'
with open(CodeFile_Path+'/'+CodeFile_Name, 'w') as fp: 
    pass
"""
CodeFile_Path = 'P:/Professional_Mission/Part_2_2/Fuschia_Research/Automation_SQL/03.Actual/33.cs'
CodeFileComment = '//' 

# Step-02-02-Code File Operation
CodeFile = open(CodeFile_Path,"w") 

# Step-02-03-DataLayer Library Work
# Prepare String
Temp_xxxxLine_Code = ''
Temp_Str_01 = 'using System;\nusing System.Collections.Generic;\nusing System.Linq;\nusing System.Threading.Tasks;\n//\nusing System.Data;\nusing System.Data.SqlClient;\n'
Temp_xxxxLine_Code = Temp_Str_01
# print(Temp_Str_01)
# Write To File
CodeFile = open(CodeFile_Path,"a") 
CodeFile.writelines(Temp_xxxxLine_Code)
CodeFile.close()
# Read From File

# Step-02-04-Namespace xxx Line Work
# Step-02-03-Database Comment Line Work
# Namspeace-Start
Code_Namespace_xxxLine_List = ['Namespace-xxx;Curl-Start;Database-xxx;']
for Temp_xxxLine in Code_Namespace_xxxLine_List :
    # print(Temp_CommentLine)
    if Temp_xxxLine=='Namespace-xxx;Curl-Start;Database-xxx;':
        # Prepare String
        Temp_xxxLine_Code=''
        Temp_Str_01 =  'namespace '+str(xxx_Namespace)+'.Models'+'\n'
        Temp_Str_02 =  '{\n'
        Temp_Str_03 =  '    '+CodeFileComment+' '+'Database-'+Database_Name+'\n'

        Temp_xxxLine_Code = Temp_xxxLine_Code + Temp_Str_01 + Temp_Str_02 + Temp_Str_03
        # print(Temp_xxxLine_Code)
        # Write To File
        CodeFile = open(CodeFile_Path,"a") 
        CodeFile.writelines(Temp_xxxLine_Code)
        CodeFile.close()
        # Read From File

# Step-02-04-Table Work
Code_Table_CommentLine_List = ['Table-xxx-Start','Operation-Basic-SetUp','Operation-Query-SP-Func','Update-All','Update-Search','Read-All','Read-Search','Delete-Search','Table-xxx-Class-Start','Table-xxx-End']
Code_Table_CommentLineCode_List = []
for Table_List_Index in range(0,len(Table_List)):   
    Temp_Table = Table_List[Table_List_Index]
    Is_CURD_Done = 0
    for Temp_CommentLine_Index in range(0,len(Code_Table_CommentLine_List)) :
        Temp_CommentLine = Code_Table_CommentLine_List[Temp_CommentLine_Index]
        # print(Temp_CommentLine)
        if Temp_CommentLine=='Table-xxx-Start':
            # Prepare String
            Temp_CommentLine_Code=''
            Temp_Str_01 = '    '+CodeFileComment+' '+Temp_CommentLine.replace('xxx',str(Temp_Table))+'\n'
            Temp_Str_02 = '    '+'public class '+str(Temp_Table)+'_DAL'+'\n'
            Temp_Str_03 = '    '+'{'+'\n'

            Temp_CommentLine_Code = Temp_CommentLine_Code + Temp_Str_01 + Temp_Str_02 + Temp_Str_03
            # print(Temp_CommentLine_Code)
            # Write To File
            CodeFile = open(CodeFile_Path,"a") 
            CodeFile.writelines(Temp_CommentLine_Code)
            CodeFile.close()
            # Read From File
        elif Temp_CommentLine=='Operation-Basic-SetUp':
            # Prepare String
            Temp_CommentLine_Code=''
            Temp_Str_01 = '        '+CodeFileComment+' '+Temp_CommentLine+'\n'
            Temp_Str_02 = '        '+'string ConnectionString = '+'"'+xxx_SQLConStr+'"'+'\n'
            Temp_CommentLine_Code = Temp_CommentLine_Code + Temp_Str_01 + Temp_Str_02
            # print(Temp_CommentLine_Code)
            # Write To File
            CodeFile = open(CodeFile_Path,"a") 
            CodeFile.writelines(Temp_CommentLine_Code)
            CodeFile.close()
            # Read From File
        elif Temp_CommentLine=='Operation-Query-SP-Func':
            # Prepare String
            Temp_CommentLine_Code=''
            Temp_Str_01 = '        '+CodeFileComment+' '+Temp_CommentLine+'\n'
            Temp_CommentLine_Code = Temp_CommentLine_Code + Temp_Str_01 
            # print(Temp_CommentLine_Code)
            # Write To File
            CodeFile = open(CodeFile_Path,"a") 
            CodeFile.writelines(Temp_CommentLine_Code)
            CodeFile.close()
            # Read From File
        elif Temp_CommentLine=='Update-All':
            # Prepare String
            Temp_CommentLine_Code=''
            Temp_Str_00 = '        '+CodeFileComment+' '+Temp_CommentLine+'\n'
            Temp_Str_01 = '        '+'public void'+' '+'UpdateAll_'+str(Temp_Table)+'('+Temp_Table+'_Model'+' '+Temp_Table+'_Model_Var'+')'+'\n'
            Temp_Str_02 = '        '+'{'+'\n'
            Temp_Str_04 = '            '+'using(SqlConnection SQLCon = new SqlConnection(ConnectionString))'+'\n'
            Temp_Str_05 = '            '+'{'+'\n'
            Temp_Str_06 = '                '+'SqlCommand SQLCmd = new SqlCommand("'+'SP_UpdateAll_'+Temp_Table+'",SQLCon);'+'\n'
            Temp_Str_07 = '                '+'SQLCmd.CommandType = CommandType.StoredProcedure;'+'\n'

            Temp_Str_13 = ''
            for Column_List_Index in range(0,len(Column_List[Table_List_Index])):
                Col_Name = Column_List[Table_List_Index][Column_List_Index]
                Col_Data_Type = DataType_List[Table_List_Index][Column_List_Index]

                Is_Identity = False
                for i in range(0,len(Constraint_List[Table_List_Index][Column_List_Index])):
                    if Constraint_List[Table_List_Index][Column_List_Index][i].startswith('Identity') or Constraint_List[Table_List_Index][Column_List_Index][i].startswith('Default'):
                        Is_Identity = True
                if(Is_Identity==False) :
                    Temp_Str_13 = Temp_Str_13 +'                '+'SQLCmd.Parameters.AddWithValue("@Var_'+Col_Name+'",'+Temp_Table+'_Model_Var.'+Col_Name+');'+'\n'
            Temp_Str_08 = '                '+'SQLCon.Open();'+'\n'
            Temp_Str_20 = '                '+'SQLCmd.ExecuteNonQuery();'+'\n'
            Temp_Str_16 = '                '+'SQLCon.Close();'+'\n'
            Temp_Str_17 = '            '+'}'+'\n'
            Temp_Str_19 = '        '+'}'+'\n'

            Temp_CommentLine_Code = Temp_CommentLine_Code+ Temp_Str_00 + Temp_Str_01 + Temp_Str_02+Temp_Str_04+ Temp_Str_05 + Temp_Str_06+ Temp_Str_07 + Temp_Str_13 + Temp_Str_08+ Temp_Str_20 + Temp_Str_16+ Temp_Str_17+ Temp_Str_19
            # print(Temp_CommentLine_Code)
            # Write To File
            CodeFile = open(CodeFile_Path,"a") 
            CodeFile.writelines(Temp_CommentLine_Code)
            CodeFile.close()
            # Read From File
        elif Temp_CommentLine=='Update-Search':
            # Prepare String
            Temp_CommentLine_Code=''
            Temp_Str_00 = '        '+CodeFileComment+' '+Temp_CommentLine+'\n'
            Temp_Str_01 = '        '+'public void'+' '+'UpdateSearch_'+str(Temp_Table)+'('+Temp_Table+'_Model'+' '+Temp_Table+'_Model_Var'+')'+'\n'

            Temp_Str_02 = '        '+'{'+'\n'
            Temp_Str_04 = '            '+'using(SqlConnection SQLCon = new SqlConnection(ConnectionString))'+'\n'
            Temp_Str_05 = '            '+'{'+'\n'
            Temp_Str_06 = '                '+'SqlCommand SQLCmd = new SqlCommand("'+'SP_UpdateSearch_'+Temp_Table+'",SQLCon);'+'\n'
            Temp_Str_07 = '                '+'SQLCmd.CommandType = CommandType.StoredProcedure;'+'\n'

            Temp_Str_13 = ''
            for Column_List_Index in range(0,len(Column_List[Table_List_Index])):
                Col_Name = Column_List[Table_List_Index][Column_List_Index]
                Col_Data_Type = DataType_List[Table_List_Index][Column_List_Index]

                Is_Identity = False
                for i in range(0,len(Constraint_List[Table_List_Index][Column_List_Index])):
                    if Constraint_List[Table_List_Index][Column_List_Index][i].startswith('Identity'):
                        Is_Identity = True
                Temp_Str_13 = Temp_Str_13 +'                '+'SQLCmd.Parameters.AddWithValue("@Var_'+Col_Name+'",'+Temp_Table+'_Model_Var.'+Col_Name+');'+'\n'
    
            Temp_Str_08 = '                '+'SQLCon.Open();'+'\n'
            Temp_Str_20 = '                '+'SQLCmd.ExecuteNonQuery();'+'\n'
            Temp_Str_16 = '                '+'SQLCon.Close();'+'\n'
            Temp_Str_17 = '            '+'}'+'\n'
            Temp_Str_19 = '        '+'}'+'\n'

            Temp_CommentLine_Code = Temp_CommentLine_Code+ Temp_Str_00 + Temp_Str_01 + Temp_Str_02+Temp_Str_04+ Temp_Str_05 + Temp_Str_06+ Temp_Str_07 + Temp_Str_13 + Temp_Str_08+ Temp_Str_20 + Temp_Str_16+ Temp_Str_17+ Temp_Str_19
            # print(Temp_CommentLine_Code)
            # Write To File
            CodeFile = open(CodeFile_Path,"a") 
            CodeFile.writelines(Temp_CommentLine_Code)
            CodeFile.close()
            # Read From File
        if Temp_CommentLine=='Read-All':
            # Prepare String
            Temp_CommentLine_Code=''
            Temp_Str_00 = '        '+CodeFileComment+' '+Temp_CommentLine+'\n'
            Temp_Str_01 = '        '+'public IEnumerable<'+str(Temp_Table)+'_Model'+'>'+' '+'ReadAll_'+str(Temp_Table)+'()'+'\n'
            Temp_Str_02 = '        '+'{'+'\n'
            Temp_Str_03 = '            '+'List<'+str(Temp_Table)+'_Model'+'>'+' '+Temp_Table+'_List = new List<'+Temp_Table+'_Model'+'>();'+'\n'
            Temp_Str_04 = '            '+'using(SqlConnection SQLCon = new SqlConnection(ConnectionString))'+'\n'
            Temp_Str_05 = '            '+'{'+'\n'
            Temp_Str_06 = '                '+'SqlCommand SQLCmd = new SqlCommand("'+'SP_ReadAll_'+Temp_Table+'",SQLCon);'+'\n'
            Temp_Str_07 = '                '+'SQLCmd.CommandType = CommandType.StoredProcedure;'+'\n'
            Temp_Str_08 = '                '+'SQLCon.Open();'+'\n'
            Temp_Str_09 = '                '+'SqlDataReader SQLRdr = SQLCmd.ExecuteReader();'+'\n'
            Temp_Str_10 = '                '+'while (SQLRdr.Read())'+'\n'
            Temp_Str_11 = '                '+'{'+'\n'
            Temp_Str_12 = '                    '+Temp_Table+'_Model'+' '+Temp_Table+'_Model_Var'+' = new '+Temp_Table+'_Model();'+'\n'
            Temp_Str_13 = ''
            for Column_List_Index in range(0,len(Column_List[Table_List_Index])):
                Col_Name = Column_List[Table_List_Index][Column_List_Index]
                Col_Data_Type = DataType_List[Table_List_Index][Column_List_Index]
                if Col_Data_Type.startswith('int') == True: 
                    Temp_Str_13 = Temp_Str_13 +'                    '+Temp_Table+'_Model_Var'+'.'+Col_Name+' = Convert.ToInt32(SQLRdr["'+Col_Name+'"]);'+'\n'
                if Col_Data_Type.startswith('varchar') == True: 
                    Temp_Str_13 = Temp_Str_13 +'                    '+Temp_Table+'_Model_Var'+'.'+Col_Name+' = SQLRdr["'+Col_Name+'"].ToString();'+'\n'
            Temp_Str_14 = '                    '+Temp_Table+'_List.Add('+Temp_Table+'_Model_Var);'+'\n'
            Temp_Str_15 = '                '+'}'+'\n'
            Temp_Str_16 = '                '+'SQLCon.Close();'+'\n'
            Temp_Str_17 = '            '+'}'+'\n'
            Temp_Str_18 = '            '+'return '+Temp_Table+'_List;'+'\n'
            Temp_Str_19 = '        '+'}'+'\n'

            Temp_CommentLine_Code = Temp_CommentLine_Code+ Temp_Str_00 + Temp_Str_01 + Temp_Str_02+ Temp_Str_03 + Temp_Str_04+ Temp_Str_05 + Temp_Str_06+ Temp_Str_07 + Temp_Str_08+ Temp_Str_09 + Temp_Str_10+ Temp_Str_11 + Temp_Str_12+ Temp_Str_13 + Temp_Str_14+ Temp_Str_15 + Temp_Str_16+ Temp_Str_17+ Temp_Str_18+ Temp_Str_19
            # print(Temp_CommentLine_Code)
            # Write To File
            CodeFile = open(CodeFile_Path,"a") 
            CodeFile.writelines(Temp_CommentLine_Code)
            CodeFile.close()
            # Read From File
        if Temp_CommentLine=='Read-Search':
            # Prepare String
            Temp_CommentLine_Code=''
            Temp_Str_00 = '        '+CodeFileComment+' '+Temp_CommentLine+'\n'
            Temp_Str_01 = '        '+'public IEnumerable<'+str(Temp_Table)+'_Model'+'>'+' '+'ReadSearch_'+str(Temp_Table)+'('+'int ReadSearch_Var'+')'+'\n'

            Temp_Str_02 = '        '+'{'+'\n'
            Temp_Str_03 = '            '+'List<'+str(Temp_Table)+'_Model'+'>'+' '+Temp_Table+'_List = new List<'+Temp_Table+'_Model'+'>();'+'\n'
            Temp_Str_04 = '            '+'using(SqlConnection SQLCon = new SqlConnection(ConnectionString))'+'\n'
            Temp_Str_05 = '            '+'{'+'\n'
            Temp_Str_06 = '                '+'SqlCommand SQLCmd = new SqlCommand("'+'SP_ReadSearch_'+Temp_Table+'",SQLCon);'+'\n'
            Temp_Str_07 = '                '+'SQLCmd.CommandType = CommandType.StoredProcedure;'+'\n'

            Temp_Str_20 = '                '+'SQLCmd.Parameters.AddWithValue("@Var_Var_'+ReadSearch_Col+'",'+'ReadSearch_Var'+');'+'\n'

            Temp_Str_08 = '                '+'SQLCon.Open();'+'\n'
            Temp_Str_09 = '                '+'SqlDataReader SQLRdr = SQLCmd.ExecuteReader();'+'\n'
            Temp_Str_10 = '                '+'while (SQLRdr.Read())'+'\n'
            Temp_Str_11 = '                '+'{'+'\n'
            Temp_Str_12 = '                    '+Temp_Table+'_Model'+' '+Temp_Table+'_Model_Var'+' = new '+Temp_Table+'_Model'+'();'+'\n'
            Temp_Str_13 = ''
            for Column_List_Index in range(0,len(Column_List[Table_List_Index])):
                Col_Name = Column_List[Table_List_Index][Column_List_Index]
                Col_Data_Type = DataType_List[Table_List_Index][Column_List_Index]

                if Col_Data_Type.startswith('int') == True: 
                    Temp_Str_13 = Temp_Str_13 +'                    '+Temp_Table+'_Model_Var'+'.'+Col_Name+' = Convert.ToInt32(SQLRdr["'+Col_Name+'"]);'+'\n'
                if Col_Data_Type.startswith('varchar') == True: 
                    Temp_Str_13 = Temp_Str_13 +'                    '+Temp_Table+'_Model_Var'+'.'+Col_Name+' = SQLRdr["'+Col_Name+'"].ToString();'+'\n'
            Temp_Str_14 = '                    '+Temp_Table+'_List.Add('+Temp_Table+'_Model_Var);'+'\n'
            Temp_Str_15 = '                '+'}'+'\n'
            Temp_Str_16 = '                '+'SQLCon.Close();'+'\n'
            Temp_Str_17 = '            '+'}'+'\n'
            Temp_Str_18 = '            '+'return '+Temp_Table+'_List;'+'\n'
            Temp_Str_19 = '        '+'}'+'\n'

            Temp_CommentLine_Code = Temp_CommentLine_Code+ Temp_Str_00 + Temp_Str_01 + Temp_Str_02+ Temp_Str_03 + Temp_Str_04+ Temp_Str_05 + Temp_Str_06+ Temp_Str_07 + Temp_Str_20 + Temp_Str_08+ Temp_Str_09 + Temp_Str_10+ Temp_Str_11 + Temp_Str_12+ Temp_Str_13 + Temp_Str_14+ Temp_Str_15 + Temp_Str_16+ Temp_Str_17+ Temp_Str_18+ Temp_Str_19
            # print(Temp_CommentLine_Code)
            # Write To File
            CodeFile = open(CodeFile_Path,"a") 
            CodeFile.writelines(Temp_CommentLine_Code)
            CodeFile.close()
            # Read From File
        elif Temp_CommentLine=='Delete-Search':
            # Prepare String
            Temp_CommentLine_Code=''
            Temp_Str_00 = '        '+CodeFileComment+' '+Temp_CommentLine+'\n'
            Temp_Str_01 = '        '+'public '+'void'+' '+'DeleteSearch_'+str(Temp_Table)+'('+'int DeleteSearch_Var'+')'+'\n'
            Temp_Str_02 = '        '+'{'+'\n'
            Temp_Str_04 = '            '+'using(SqlConnection SQLCon = new SqlConnection(ConnectionString))'+'\n'
            Temp_Str_05 = '            '+'{'+'\n'
            Temp_Str_06 = '                '+'SqlCommand SQLCmd = new SqlCommand("'+'SP_DeleteSearch_'+Temp_Table+'",SQLCon);'+'\n'
            Temp_Str_07 = '                '+'SQLCmd.CommandType = CommandType.StoredProcedure;'+'\n'
            Temp_Str_13 = '                '+'SQLCmd.Parameters.AddWithValue("@Var_'+DeleteSearch_Col+'",'+'DeleteSearch_Var'+');'+'\n'
            Temp_Str_08 = '                '+'SQLCon.Open();'+'\n'
            Temp_Str_20 = '                '+'SQLCmd.ExecuteNonQuery();'+'\n'
            Temp_Str_16 = '                '+'SQLCon.Close();'+'\n'
            Temp_Str_17 = '            '+'}'+'\n'
            Temp_Str_19 = '        '+'}'+'\n'
            Temp_CommentLine_Code = Temp_CommentLine_Code+ Temp_Str_00 + Temp_Str_01 + Temp_Str_02+Temp_Str_04+ Temp_Str_05 + Temp_Str_06+ Temp_Str_07 + Temp_Str_13 + Temp_Str_08+ Temp_Str_20 + Temp_Str_16+ Temp_Str_17+ Temp_Str_19
            # print(Temp_CommentLine_Code)
            # Write To File
            CodeFile = open(CodeFile_Path,"a") 
            CodeFile.writelines(Temp_CommentLine_Code)
            CodeFile.close()
            # Read From File
        elif Temp_CommentLine=='Table-xxx-End':
            # Prepare String
            Temp_CommentLine_Code=''
            Temp_Str_01 = '    '+'}'+'\n'
            Temp_Str_02 = '    '+'}'+'\n'
            Temp_Str_03 = '    '+CodeFileComment+' '+Temp_CommentLine.replace('xxx',str(Temp_Table))+'\n'
            Temp_CommentLine_Code = Temp_CommentLine_Code + Temp_Str_02 + Temp_Str_03
            # print(Temp_CommentLine_Code)
            # Write To File
            CodeFile = open(CodeFile_Path,"a") 
            CodeFile.writelines(Temp_CommentLine_Code)
            CodeFile.close()
            # Read From File

# Namspeace-End
Code_Namespace_xxxLine_List = ['Curl-End']
for Temp_xxxLine in Code_Namespace_xxxLine_List :
    # print(Temp_CommentLine)
    if Temp_xxxLine=='Curl-End':
        # Prepare String
        Temp_xxxLine_Code=''
        Temp_Str_01 =  '}\n'
        Temp_xxxLine_Code = Temp_xxxLine_Code + Temp_Str_01
        # Write To File
        CodeFile = open(CodeFile_Path,"a") 
        CodeFile.writelines(Temp_xxxLine_Code)
        CodeFile.close()
        # Read From File 