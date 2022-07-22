from Service_Excel import *
# -----------------------------------
Excel_Workbook = Create_Excel_Workbook_Mannual(Get_Excel_File_Path("002"))
Excel_Workbook_Format = Create_Excel_Workbook_Format(Excel_Workbook,"NumFormat:Currency-Color:red-FontBold:True-FontColor:Black-BgColor:Red-Border:1")
# -----------------------------------
Excel_Workbook_Worksheet = Create_Excel_Workbook_Worksheet(Excel_Workbook)
Excel_Workbook_Worksheet = Create_Excel_Workbook_Worksheet_Format(Excel_Workbook_Worksheet,[0,5],"NA",Excel_Workbook_Format)

Update_Excel_Workbook_Worksheet_Cell_ContentFormat(Excel_Workbook_Worksheet,"A1",1234.56,Excel_Workbook_Format)
Update_Excel_Workbook_Worksheet_Cell_ContentFormat(Excel_Workbook_Worksheet,[2,3],1234.56,Excel_Workbook_Format)
# -----------------------------------
Excel_Workbook.close()
# -----------------------------------
