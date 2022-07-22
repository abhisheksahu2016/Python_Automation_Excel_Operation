# File List in a Foldler Path 
import os
path='F:/Sahu_Group/C1_Data/B1_DBFile_SQLServer'
Chk_Dir = os.path.exists(path) 
print(Chk_Dir)
if Chk_Dir is True :
    try : 
         File_List=[]
         for dirpath, dirnames, filenames in os.walk(path):
            for file in filenames:
                File_List.append(file)
                print(file)
    except OSError as error: 
        print(error)
else :
    print("Root Directory is not Avilabel")
#print(File_List)
