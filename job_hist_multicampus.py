# -*- coding: utf-8 -*-
"""
Created on Wed Mar 11 14:22:02 2020

@author: sayers
"""

from src.emailautosend import mailthis
from src.emailautosend import getemail
import os
from src.cleansheet import *
import pandas as pd
import re
from datetime import datetime 
from matplotlib import pyplot as plt

def newest(path,fname):
    files = os.listdir(path)
    paths = [os.path.join(path, basename) for basename in files if basename.startswith(fname)]
    return max(paths, key=os.path.getmtime)

path = "S:\\Downloads\\"     # Give the location of the files
fname = "JOB_HIST"         # Give filename prefix
df = pd.read_excel(newest(path,fname))  #getting the newest of these files in the directory and converting to df
#stripping out the 2 metadata columns in CJR files
if re.match("R1013",df.columns.values[0]).group() == "R1013":
    new_header = df.iloc[1] #grab the first row for the header
    df = df[2:] #take the data less the header row
    df.columns = new_header #set the header row as the df header
#standardizing the column names
df.columns = df.columns.str.strip().str.lower().str.replace(' ', '_').str.replace('(', '').str.replace(')', '')


"""
for i in sorted(list(df.columns.values)):
    print(i)
    
column_list = df.columns.values.tolist()
for column_name in column_list:
      print(df[column_name].unique())
"""
multicampus = df[df.id.isin(df[~(df.unit=="YRK01")][df.pay_status.isin(['A','W','P','L'])].id.unique())]

multicampus.to_excel("S:\\Downloads\\multicampus.xls")
df.empl_class.unique()