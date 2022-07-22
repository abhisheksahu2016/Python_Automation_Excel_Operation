from Service_Excel import *
# -----------------------------------
Excel_Workbook = XLWT_Create_Excel_Workbook_Blank()
# -----------------------------------
Excel_Workbook,Excel_Workbook_Worksheet = XLWT_Create_Excel_Workbook_Worksheet_Manual(Excel_Workbook,"Sheet1")
Excel_Workbook_Worksheet = XLWT_Write_Excel_Workbook_Worksheet_FromList(Excel_Workbook_Worksheet,[0,0],["AA","BB"],"V")
# -----------------------------------
Excel_Workbook = XLWT_Create_Excel_FromWorkbook(Excel_Workbook,Get_Excel_File_Path("003"))

Excel_Workbook_Open = XLRD_Open_Excel_Workbook_Mannual(Get_Excel_File_Path("003"))
Excel_Workbook_Worksheet = XLRD_Open_Excel_Workbook_Worksheet_ByIndex(Excel_Workbook_Open,0)
# -----------------------------------
