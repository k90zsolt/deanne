import pandas as pd
import numpy as np
from openpyxl import Workbook
import csv
import os
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell
from shutil import copyfile
from pathlib import Path
import json
from pandas.io.json import json_normalize

p = Path(r'c:\Data\json\BackRegisterlabel.json')
with p.open('r', encoding='utf-8') as f:
    data = json.loads(f.read())
df = pd.json_normalize(data['Labels'])

df = df.sort_values(by =['Label.Name', 'Timestamp'])

writer = pd.ExcelWriter('backregisterlabelchart.xlsx', engine='xlsxwriter')

df.to_excel(writer, sheet_name='Sheet1')
#df = df.sort_values(["Person.Index"], axis=0, ascending=[True])
#https://thispointer.com/pandas-sort-rows-or-columns-in-dataframe-based-on-values-using-dataframe-sort_values/
writer.save()
