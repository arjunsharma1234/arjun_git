# -*- coding: utf-8 -*-
import sys
from glob import  glob
import pandas as pd
import os
from os import  path
import string
# print list(string.ascii_lowercase)
fil='2018_feb_Zoho_Leads.csv'
df = pd.read_csv(fil)
df = df[["POTENTIALID"]]
df["LEAD STATUS"] = 'Lead Lost'
df.to_csv(fil,index=False)
csvfilename = open(fil, 'r').readlines()
file = 1
for j in range(len(csvfilename)):
 if j % 29999 == 0:
  open(str(fil)+ str(file) + '.csv', 'w+').writelines(csvfilename[j:j+29999])
 file += 1



for csv_file in glob('*.csv'):
    print csv_file
    Cov = pd.read_csv(csv_file, header=None)
    Cov.columns = ["LEADID", "LEAD STATUS"]
    Cov = pd.DataFrame(Cov)
    Cov.to_csv(csv_file,index=False)