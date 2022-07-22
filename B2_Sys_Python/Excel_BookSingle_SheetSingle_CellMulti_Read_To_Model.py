# ----  dd.py
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

# from cc.py
# Step-01-02-Excel Data To Code Data Replacementation
"""
xxx_Namespace = 'xxx_Project'
"""
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
CodeFile_Name ='zzz.cs'
with open(CodeFile_Path+'/'+CodeFile_Name, 'w') as fp: 
    pass
"""
CodeFile_Path = 'P:/Professional_Mission/Part_2_2/Fuschia_Research/Automation_SQL/03.Actual/44.cs'
CodeFileComment = '//' 

# Step-02-02-Code File Operation
CodeFile = open(CodeFile_Path,"w") 

# Step-02-03-DataLayer Library Work
# Prepare String
Temp_xxxxLine_Code = ''
Temp_Str_01 = 'using System;\nusing System.Collections.Generic;\nusing System.ComponentModel.DataAnnotations;\nusing System.Linq;\nusing System.Threading.Tasks;\n'        

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
Code_Table_CommentLine_List = ['Table-xxx-Start','Column Fill','Table-xxx-End']
Code_Table_CommentLineCode_List = []

Column_DataAnnotian_List = ['Key','Timestamp','ConcurrencyCheck','Required','MinLength','MaxLength','StringLength']

for Table_List_Index in range(0,len(Table_List)):   
    Temp_Table = Table_List[Table_List_Index]
    for Temp_CommentLine_Index in range(0,len(Code_Table_CommentLine_List)) :
        Temp_CommentLine = Code_Table_CommentLine_List[Temp_CommentLine_Index]
        # print(Temp_CommentLine)
        if Temp_CommentLine=='Table-xxx-Start':
            # Prepare String
            Temp_CommentLine_Code=''
            Temp_Str_01 = '    '+CodeFileComment+' '+Temp_CommentLine.replace('xxx',str(Temp_Table))+'\n'
            Temp_Str_02 = '    '+'public class '+str(Temp_Table)+'_Model'+'\n'
            Temp_Str_03 = '    '+'{'+'\n'

            Temp_CommentLine_Code = Temp_CommentLine_Code + Temp_Str_01 + Temp_Str_02 + Temp_Str_03
            # print(Temp_CommentLine_Code)
            # Write To File
            CodeFile = open(CodeFile_Path,"a") 
            CodeFile.writelines(Temp_CommentLine_Code)
            CodeFile.close()
            # Read From File
        elif Temp_CommentLine=='Column Fill':
            # Prepare String
            Temp_CommentLine_Code=''
            for Column_List_Index in range(0,len(Column_List[Table_List_Index])):
                Temp_DataAnnotian_Str = ''
                for i in range(0,len(Constraint_List[Table_List_Index][Column_List_Index])):
                    if Constraint_List[Table_List_Index][Column_List_Index][i].startswith('Not Null') or Constraint_List[Table_List_Index][Column_List_Index][i].startswith('Identity'):
                        Temp_DataAnnotian_Str = Temp_DataAnnotian_Str + '		'+'[Required]' + '\n'

                Temp_Column_Str = ''
                Col_Name = Column_List[Table_List_Index][Column_List_Index]
                Col_Data_Type = DataType_List[Table_List_Index][Column_List_Index]
                Temp_DataType = ''
                if Col_Data_Type.startswith('int'):
                    Temp_DataType = 'int'
                elif Col_Data_Type.startswith('varchar'):
                    Temp_DataType = 'string'
                elif Col_Data_Type.startswith('datetime'):
                    Temp_DataType = 'DateTime'
                Temp_Column_Str  = '		'+'public'+' '+Temp_DataType+' '+Col_Name+' { get; set; } '+'\n'
                Temp_CommentLine_Code = Temp_CommentLine_Code + Temp_DataAnnotian_Str + Temp_Column_Str

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
