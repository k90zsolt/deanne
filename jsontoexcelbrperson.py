import pandas as pd
import numpy as np
from openpyxl import Workbook
import csv
import os
import xlsxwriter
from shutil import copyfile
from pathlib import Path
import json
from pandas.io.json import json_normalize

p = Path(r'C:\Data\json\backregisterperson.json')
with p.open('r', encoding='utf-8') as f:
    data = json.loads(f.read())
df = pd.json_normalize(data['Persons'])

#print (df)

export_csv = df.to_csv (r'C:\Data\json\Newdocumnet.csv', encoding= 'utf-8', index = False, header = True)

read_file = pd.read_csv (r'C:\Data\json\Newdocumnet.csv')

read_file.to_excel (r'C:\Data\json\backregsiterperson.xlsx', index = None, header= True)

os.remove(r'C:\Data\json\Newdocumnet.csv')

#https://www.youtube.com/watch?v=FXhED53VZ50

#https://xlsxwriter.readthedocs.io/working_with_pandas.html





