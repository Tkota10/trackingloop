import numpy as np
import pandas as pd
from openpyxl import load_workbook
import openpyxl
import xlrd
import pyvo
import math
		
george=pd.read_excel('CatWISEDiscoveries.xlsx',sheet_name='maincandidates',header=None,na_values=['NA'],usecols="DR,DS,DW,DX,IG",skiprows=20,nrows=5)
george=george.to_numpy()
lipu=list(george)
i = 0 
while i < (len(lipu)):
    writer = pd.ExcelWriter('MarksData.xlsx', engine='openpyxl')
    # try to open an existing workbook
    writer.book = load_workbook('MarksData.xlsx')
    # copy existing sheets
    writer.sheets = dict((ws.title, ws) for ws in writer.book.worksheets)
    # read existing file
    reader = pd.read_excel(r'MarksData.xlsx')
    mango=pd.read_excel('CatWISEDiscoveries.xlsx',sheet_name='maincandidates',header=None,na_values=['NA'],usecols="DR,DS,DW,DX,DY,DZ,EG,EH,EJ,EK",skiprows=int(george[i][4]),nrows=1)
    mango.to_excel(writer,index=False,header=False,startrow=len(reader)+1) 
    newra = lipu[i][0] + (1999 - 2015)*(lipu[i][2]/(math.cos(3.14159*(lipu[i][1]/180))*3600))
    newdec= lipu[i][1]+ (1998- 2015)*lipu[i][3]/3600
    url = "https://irsa.ipac.caltech.edu/SCS?table=fp_psc"
    objects = pyvo.conesearch(url, pos=(newra, newdec), radius = 0.001)
    value=objects.table['ra','dec','j_m','j_msigcom', 'h_m', 'h_msigcom','k_m','k_msigcom']
    df=value.to_pandas()
    df.to_excel(writer,index=False,header=False,startrow=len(reader)+1,startcol=10)
    i+=1
    print("download is succesful")
    writer.close()