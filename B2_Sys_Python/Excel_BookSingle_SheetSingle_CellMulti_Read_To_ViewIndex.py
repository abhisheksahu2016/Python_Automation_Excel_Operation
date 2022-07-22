# ----  ee.py
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
# Step-01-02-Excel Data To Code Data Replacementation
xxx_Project = 'CoreAdminLTE'

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
CodeFile_Path = 'P:/Professional_Mission/Part_2_2/Fuschia_Research/Automation_SQL/03.Actual/55.cshtml'
CodeFileComment = '//' 

# Step-02-02-Code File Operation
CodeFile = open(CodeFile_Path,"w") 
# Step-02-03-DataLayer Library Work
# Prepare String
Temp_xxxxLine_Code = ''
Temp_Str_01 = '@using '+xxx_Project+'\n'
Temp_Str_02 = '@addTagHelper *, Microsoft.AspNetCore.Mvc.TagHelpers\n@addTagHelper *, AuthoringTagHelpers\n'
Temp_Str_03 = ''
for Table_List_Index in range (0,len(Table_List)):
    Temp_Str_03 = Temp_Str_03 + '@model '+'IEnumerable<'+xxx_Project+'.Models.'+Table_List[Table_List_Index]+'_Model'+'>'+'\n'
Temp_Str_04 = '@{'+'\n'
Temp_Str_05 = '    '+'ViewData["Title"] = " Index Page";'+'\n'
Temp_Str_06 = '    '+'Layout = "~/Views/A0101Shared/A01Layouts/_A0202StarterIndexLayout.cshtml";'+'\n'
Temp_Str_07 = '}'+'\n'

Temp_xxxxLine_Code = Temp_Str_01 + Temp_Str_02 + Temp_Str_03+ Temp_Str_04+ Temp_Str_05+ Temp_Str_06+ Temp_Str_07
# print(Temp_Str_01)
# Write To File
CodeFile = open(CodeFile_Path,"a") 
CodeFile.writelines(Temp_xxxxLine_Code)
CodeFile.close()
# Read From File

# Step-02-04-Table Work
Code_Table_CommentLine_List = ['Table-Name-Start','Table-Type-Start','Row-Start','Card-Start','Card-Header-Start','Card-Header-Middle','Card-Header-End','Card-Body-Start','Table-Start','Thead-Start','Thead-Middle','Thead-End','Tbody-Start','Tbody-Middle','Tbody-End','Table-End;Card-Body-End;Card-End;Row-End;Table-Type-End;Table-Name-End']
for Table_List_Index in range(0,len(Table_List)):   
    Temp_Table = Table_List[Table_List_Index]
    for Temp_CommentLine_Index in range(0,len(Code_Table_CommentLine_List)) :
        Temp_CommentLine = Code_Table_CommentLine_List[Temp_CommentLine_Index]
        # print(Temp_CommentLine)
        if Temp_CommentLine=='Table-Name-Start':
            # Prepare String
            Temp_CommentLine_Code=''
            Temp_Str_01 = '<!-- '+'Table-'+Temp_Table+'-Start'+' -->'+'\n'
            Temp_CommentLine_Code = Temp_CommentLine_Code + Temp_Str_01 
            # print(Temp_CommentLine_Code)
            # Write To File
            CodeFile = open(CodeFile_Path,"a") 
            CodeFile.writelines(Temp_CommentLine_Code)
            CodeFile.close()
            # Read From File
        elif Temp_CommentLine=='Table-Type-Start':
            # Prepare String
            Temp_CommentLine_Code=''
            Temp_Str_01 = '<!-- Data-xxx-Table-06 -->'+'\n'
            Temp_CommentLine_Code = Temp_CommentLine_Code + Temp_Str_01 
            # print(Temp_CommentLine_Code)
            # Write To File
            CodeFile = open(CodeFile_Path,"a") 
            CodeFile.writelines(Temp_CommentLine_Code)
            CodeFile.close()
            # Read From File
        elif Temp_CommentLine=='Row-Start':
            # Prepare String
            Temp_CommentLine_Code=''
            Temp_Str_01 = '<div class="row">'+'\n'+'    '+'<div class="col-12">'+'\n'
            Temp_CommentLine_Code = Temp_CommentLine_Code + Temp_Str_01 
            # print(Temp_CommentLine_Code)
            # Write To File
            CodeFile = open(CodeFile_Path,"a") 
            CodeFile.writelines(Temp_CommentLine_Code)
            CodeFile.close()
            # Read From File
        elif Temp_CommentLine=='Card-Start':
            # Prepare String
            Temp_CommentLine_Code=''
            Temp_Str_01 = '        '+'<div class="card">'+'\n'
            Temp_CommentLine_Code = Temp_CommentLine_Code + Temp_Str_01 
            # print(Temp_CommentLine_Code)
            # Write To File
            CodeFile = open(CodeFile_Path,"a") 
            CodeFile.writelines(Temp_CommentLine_Code)
            CodeFile.close()
            # Read From File
        elif Temp_CommentLine=='Card-Header-Start':
            # Prepare String
            Temp_CommentLine_Code=''
            Temp_Str_01 = '            '+'<div class="card-header">'+'\n'+'                '+'<h3 class="card-title">'+Temp_Table+' Table</h3>'+'\n'
            Temp_CommentLine_Code = Temp_CommentLine_Code + Temp_Str_01 
            # print(Temp_CommentLine_Code)
            # Write To File
            CodeFile = open(CodeFile_Path,"a") 
            CodeFile.writelines(Temp_CommentLine_Code)
            CodeFile.close()
            # Read From File
        elif Temp_CommentLine=='Card-Header-Middle':
            # Prepare String
            Temp_CommentLine_Code=''
            Temp_Str_01 = '                '+'<div class="card-tools">'+'\n'+'                   '+'<div class="input-group input-group-sm" style="width: 150px;">'+'\n'+'                      '+'<input type="text" name="table_search" class="form-control float-right" placeholder="Search">'+'\n'+'                        '+'<div class="input-group-append">'+'\n'+'                           '+'<button type="submit" class="btn btn-default">'+'\n'+'                              '+'<i class="fas fa-search"></i>'+'\n'+'                            '+'</button>'+'\n'+'                        '+'</div>'+'\n'+'                    '+'</div>'+'\n'+'                '+'</div>'+'\n'
            Temp_CommentLine_Code = Temp_CommentLine_Code + Temp_Str_01 
            # print(Temp_CommentLine_Code)
            # Write To File
            CodeFile = open(CodeFile_Path,"a") 
            CodeFile.writelines(Temp_CommentLine_Code)
            CodeFile.close()
            # Read From File
        elif Temp_CommentLine=='Card-Header-End':
            # Prepare String
            Temp_CommentLine_Code=''
            Temp_Str_01 = '            '+'</div>'+'\n'+'            '+'<!-- /.card-header -->'+'\n'
            Temp_CommentLine_Code = Temp_CommentLine_Code + Temp_Str_01 
            # print(Temp_CommentLine_Code)
            # Write To File
            CodeFile = open(CodeFile_Path,"a") 
            CodeFile.writelines(Temp_CommentLine_Code)
            CodeFile.close()
            # Read From File
        elif Temp_CommentLine=='Card-Body-Start':
            # Prepare String
            Temp_CommentLine_Code=''
            Temp_Str_01 = '            '+'<div class="card-body table-responsive p-0" style="height: 300px;">'+'\n'
            Temp_CommentLine_Code = Temp_CommentLine_Code + Temp_Str_01 
            # print(Temp_CommentLine_Code)
            # Write To File
            CodeFile = open(CodeFile_Path,"a") 
            CodeFile.writelines(Temp_CommentLine_Code)
            CodeFile.close()
            # Read From File
        elif Temp_CommentLine=='Table-Start':
            # Prepare String
            Temp_CommentLine_Code=''
            Temp_Str_01 = '                '+'<table class="table table-head-fixed text-nowrap">'+'\n'
            Temp_CommentLine_Code = Temp_CommentLine_Code + Temp_Str_01 
            # print(Temp_CommentLine_Code)
            # Write To File
            CodeFile = open(CodeFile_Path,"a") 
            CodeFile.writelines(Temp_CommentLine_Code)
            CodeFile.close()
            # Read From File
        elif Temp_CommentLine=='Thead-Start':
            # Prepare String
            Temp_CommentLine_Code=''
            Temp_Str_01 = '                    '+'<thead>'+'\n'
            Temp_CommentLine_Code = Temp_CommentLine_Code + Temp_Str_01 
            # print(Temp_CommentLine_Code)
            # Write To File
            CodeFile = open(CodeFile_Path,"a") 
            CodeFile.writelines(Temp_CommentLine_Code)
            CodeFile.close()
            # Read From File
        elif Temp_CommentLine=='Thead-Middle':
            # Prepare String
            Temp_CommentLine_Code=''
            Temp_Str_01 = '                        '+'<tr>'+'\n'
            Temp_Str_02 = ''
            for Column_List_Index in range(0,len(Column_List[Table_List_Index])):
                Col_Name = Column_List[Table_List_Index][Column_List_Index]
                Temp_Str_02 = Temp_Str_02+ '                            '+'<th>@Html.DisplayNameFor(model => model.'+Col_Name+')</th>'+'\n'
            Temp_Str_03 = '                        '+'</tr>'+'\n'
            Temp_CommentLine_Code = Temp_CommentLine_Code + Temp_Str_01 + Temp_Str_02+ Temp_Str_03
            # print(Temp_CommentLine_Code)
            # Write To File
            CodeFile = open(CodeFile_Path,"a") 
            CodeFile.writelines(Temp_CommentLine_Code)
            CodeFile.close()
            # Read From File
        elif Temp_CommentLine=='Thead-End':
            # Prepare String
            Temp_CommentLine_Code=''
            Temp_Str_01 = '                    '+'</thead>'+'\n'
            Temp_CommentLine_Code = Temp_CommentLine_Code + Temp_Str_01 
            # print(Temp_CommentLine_Code)
            # Write To File
            CodeFile = open(CodeFile_Path,"a") 
            CodeFile.writelines(Temp_CommentLine_Code)
            CodeFile.close()
            # Read From File
        elif Temp_CommentLine=='Tbody-Start':
            # Prepare String
            Temp_CommentLine_Code=''
            Temp_Str_01 = '                    '+'<tbody>'+'\n'
            Temp_CommentLine_Code = Temp_CommentLine_Code + Temp_Str_01 
            # print(Temp_CommentLine_Code)
            # Write To File
            CodeFile = open(CodeFile_Path,"a") 
            CodeFile.writelines(Temp_CommentLine_Code)
            CodeFile.close()
            # Read From File
        elif Temp_CommentLine=='Tbody-Middle':
            # Prepare String
            Temp_CommentLine_Code=''
            Temp_Str_01 = '                        '+'@foreach (var item in Model)'+'\n'
            Temp_Str_02 = '                        '+'{'+'\n'
            Temp_Str_03 = '                            '+'<tr>'+'\n'
            Temp_Str_04 = ''
            for Column_List_Index in range(0,len(Column_List[Table_List_Index])):
                Col_Name = Column_List[Table_List_Index][Column_List_Index]
                Temp_Str_04 = Temp_Str_04+ '                                <td>@Html.DisplayFor(modelItem => item.'+Col_Name+')</td>'+'\n'
            Temp_Str_05 = '                                '+'<td>'+'\n'+'                                    '+'<a asp-action="Update" asp-route-id="@item.UserCategory_ID">Update</a> |'+'\n'+'                                    '+'<a asp-action="Read" asp-route-id="@item.UserCategory_ID">Read</a> |'+'\n'+'                                    '+'<a asp-action="Delete" asp-route-id="@item.UserCategory_ID">Delete</a>'+'\n'+'                                '+'</td>'+'\n'+'                            '+'</tr>'+'\n'+'                        '+'}'+'\n'
            Temp_CommentLine_Code = Temp_CommentLine_Code + Temp_Str_01  + Temp_Str_02 + Temp_Str_03  + Temp_Str_04 + Temp_Str_05
            # print(Temp_CommentLine_Code)
            # Write To File
            CodeFile = open(CodeFile_Path,"a") 
            CodeFile.writelines(Temp_CommentLine_Code)
            CodeFile.close()
            # Read From File
        elif Temp_CommentLine=='Table-End;Card-Body-End;Card-End;Row-End;Table-Type-End;Table-Name-End':
            # Prepare String
            Temp_CommentLine_Code=''
            Temp_Str_01 = '               '+'</table>'+'\n'+'            '+'</div>'+'\n'+'            '+'<!-- /.card-body -->'+'\n'+'        '+'</div>'+'\n'+'        '+'<!-- /.card -->'+'\n'+'    '+'</div>'+'\n'+'</div>'+'\n'+'<!-- ./Data-xxx-Table-06 -->'+'\n'+'<!-- '+'Table-'+Temp_Table+'-End'+' -->'+'\n'
            Temp_CommentLine_Code = Temp_CommentLine_Code + Temp_Str_01 
            # print(Temp_CommentLine_Code)
            # Write To File
            CodeFile = open(CodeFile_Path,"a") 
            CodeFile.writelines(Temp_CommentLine_Code)
            CodeFile.close()
            # Read From File