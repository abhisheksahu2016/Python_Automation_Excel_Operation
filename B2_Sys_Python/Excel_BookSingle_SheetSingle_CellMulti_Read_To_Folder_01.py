# ------------File Folder Auditing by Program ----------------
import os
import xlwt 
from xlwt import Workbook 
from numpy import * 
import numpy as np
"""
For Auto
SubjectList=[]
SubjectIndex=[]
# ---------------------------------------------------------------------------------------------
PrevSubjectIndexStart = 1
SubjectHeading = 'Sj'
AfterSubjectIndexStart = 1 
NoOfSub=3
# ---------------------------------------------------------------------------------------------
#----- Crates Subject with after
for i in range(1,NoOfSub+1):
    SubjectList.append(str(SubjectHeading+'-'+str(AfterSubjectIndexStart)))
    AfterSubjectIndexStart=AfterSubjectIndexStart+1
#----- Crates before subject
SubjectIndex = [i for i in range(PrevSubjectIndexStart,PrevSubjectIndexStart+NoOfSub+1)]
"""
# ---------------------------------------------------------------------------------------------
#SubjectList = ['Sj-20-1-1','Sj-20-1-2','Sj-20-1-3','Sj-20-1-4','Sj-20-1-5','Sj-20-1-6','Sj-20-1-7','Sj-20-1-8','Sj-20-1-9','Sj-20-1-10','Sj-20-1-11','Sj-20-2-1','Sj-20-2-2','Sj-20-2-3','Sj-20-2-4','Sj-20-2-5']
#SubjectIndex = [i for i in range(PrevSubjectIndexStart,PrevSubjectIndexStart+len(SubjectList))]

#SubjectList = ['Sj-1','Sj-3','Sj-5','Sj-7','Sj-9','Sj-11']
#SubjectIndex = [1 ,3,5,7,9,11]

SubjectList = ['Mnemonics','Fitting','Mind_Map','X_Memory','Govinda_Memory','Application','Subconcious_Mind','Photography_Memory','Procedural_Memory','Artificial_AVR','Sensory_Memory','Garbage_Collector','Affrimator','Taprecorder_Memory','Application_Memory','X_Memory','Canvas_Memory','Memory_Of_Locci','Episodic_Memory','Semantic_Memory','Concious_Mind','Settings','Time_Liner','Visualization_Tool','Iron_Man','Control_Panel','Shatiman','Jarvis','Param','SpyBee_Logicmusk','Applications','Rocket','Memory_Palace','Photography_Memory','AVR_Memory','AtmaSarankhana_Unit','Arc_Reactor']
PrevSubjectIndexStart = 1
SubjectIndex = [i for i in range(PrevSubjectIndexStart,PrevSubjectIndexStart+len(SubjectList))]
# ---------------------------------------------------------------------------------------------

SubjectFolderList = ['Pro_Service','Pro_Gram','3.Pro_Data']
SubjectFolderIndex= [1            ,2         ,3           ]
SubjectFolderFileList = [['V-0.0.0.txt.lnk'],['ResourceMap.xlsx'],['ChapterMap.xlsx'],['TopicMap.xlsx'],['ConceptMap.xlsx'],['ObjectMap.xlsx'],['DataMap.xlsx'],['MPcMap.skp'],['ScriptMap.xlsx']]

#SubjectFolderList = ['Resource','FResource','PConeception','CConeception','Resource','FResource','PDesign','CDesign','Resource','FResource','PPresentation','CPresentation']
#SubjectFolderIndex= [11        , 12        ,13            ,14            ,21        ,22         ,23       ,24       ,31        ,32         ,33             ,34             ]
#SubjectFolderFileList = [['V-0.0.0.txt.lnk'],['V-0.0.0.txt.lnk'],['V-0.0.0.txt.lnk'],['V-0.0.0.txt.lnk'],['V-0.0.0.txt.lnk'],['V-0.0.0.txt.lnk'],['V-0.0.0.txt.lnk'],['V-0.0.0.txt.lnk'],['V-0.0.0.txt.lnk'],['V-0.0.0.txt.lnk'],['V-0.0.0.txt.lnk'],['V-0.0.0.txt.lnk'],['V-0.0.0.txt.lnk'],['V-0.0.0.txt.lnk'],['V-0.0.0.txt.lnk'],['V-0.0.0.txt.lnk']]
# ---------------------------------------------------------------------------------------------
wb = Workbook() 
Write_Sheet1 = wb.add_sheet('Sheet 1') 
XStart_Index = 2
YStart_Index = 2

XRun_Index = XStart_Index
YRun_Index = YStart_Index
# ---------------------------------------------------------------------------------------------
for i in range(0,len(SubjectList)):
    Write_Sheet1.write(YRun_Index, XStart_Index+0,str(str(SubjectIndex[i])  + '.' + str(SubjectList[i]) ) )
    for j in range(0,len(SubjectFolderList)) :
        Write_Sheet1.write(YRun_Index, XStart_Index+1,str(str(SubjectFolderIndex[j])+'.'+str(SubjectList[i])+'_'+str(SubjectFolderList[j])))
        for k in range(0,len(SubjectFolderFileList[j])):
            Write_Sheet1.write(YRun_Index, XStart_Index+2,str(str(k+1)+'.'+str(SubjectList[i]) +'_'+str(SubjectFolderFileList[j][k])) ) 
            YRun_Index = YRun_Index+1
    YRun_Index = YRun_Index + 1
# ---------------------------------------------------------------------------------------------
wb.save('example.xlsx') 
# ---------------------------------------------------------------------------------------------