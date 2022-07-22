import os
from numpy import * 
import numpy as np

path = "C:/Users/ABHISHEK/Desktop/try/x"
Chk_Dir = os.path.exists(path) 
#print(Chk_Dir)
if Chk_Dir is True :
    try : 
        Max_Row=10
        Max_Col=10
        Folder_Structure_SquizeMatrix = [[ '' for i in range(0,Max_Col)] for j in range(0,Max_Row)]
        Folder_Structure_SeqGuard = [ -1 for i in range(0,Max_Col)]
        #---------------
        rootpathlen = len(str(path).split('/'))
        rootpathsplit = str(path).split('/')
        rootFolder = rootpathsplit[len(rootpathsplit)-1]
        print(rootFolder)
        Folder_Structure_SquizeMatrix[0][0]=rootFolder
        #---------------
        Y_Level=0;
        for dirpath, dirnames, filenames in os.walk(path):
            #---------------------------------------------
            if len(dirnames)>0:
                for dirnamestr in dirnames:
            #---------------------------------------------
            
    except OSError as error: 
        print(error)
else :
    print("Root Directory is not Avilabel")
"""
Root	Folder	File
    	0	0
	    0	1
    	1	0
	    1	1
"""