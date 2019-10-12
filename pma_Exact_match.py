# -*- coding: utf-8 -*-


import pandas as pd
import os
os.chdir(r"D:\INT AM Analytics\PMA\2019\June")

#Loading data
sap = pd.read_excel('Collins_Part Numbers.xlsx')
pma = pd.read_table('PMA_Queryd.txt',sep= "|" ,engine='python')


pnum1 = sap.merge(pma, left_on = 'Material',right_on= 'ReplacedPartNumber',how="inner")
pnum2 = sap.merge(pma, left_on = 'Material',right_on= 'PartNumber',how="inner")
frame = [pnum1,pnum2]
pnum = pd.concat(frame)

pnum['pp'] = pnum['Holder'].str.strip().str.replace('Inc','').str.replace('.','').str.replace(' ','').str.replace(',','')
pnum = pnum.drop_duplicates(subset=['Material','pp'])

writer = pd.ExcelWriter('PMA Match_June.xlsx')
pnum.to_excel(writer)
writer.save()

#pma_AB = pma[pma["Models"].str.contains("Airbus")]
#pma_Boe = pma[pma["Models"].str.contains("Boeing")]
#pma_Bmb = pma[pma["Models"].str.contains("Bombardier")]
#pma_Emr = pma[pma["Models"].str.contains("Embraer")]



