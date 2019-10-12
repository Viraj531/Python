
import pandas as pd
import numpy as np
import os

os.chdir(r"D:\INT AM Analytics\Data\OHWV\New Data\1007")


bl = pd.read_excel("SIS_Open_Orders.xlsx")


bl = bl[['Profit Center (PC)','PC Market','Distribution Channel','Division','Customer (Sold To)',
         'Customer Name (Sold To)','Sales Document Type','Sales Document' ,'Item Category'
, 'Sales Document Item','Material Number', 'Material Description2','Confirmed Qty(SALE UOM)','Open Amount USD','Plant'
,'SAP Sales Order Item Create Date','Material Type','Material Group','Country (Ship To)','Request Date(Customer Dock)']]

bl = bl.rename(columns={'Profit Center (PC)':'Profit Center','PC Market':'Market','Material Description2':'Material Description','Confirmed Qty(SALE UOM)':'Quantity',
                       'Open Amount USD':'Sales Amount','SAP Sales Order Item Create Date':'Create Date',
                        'Request Date(Customer Dock)':'Request Date'})


sh = pd.read_excel("Ship Data.xlsx",sheet_name='Sheet1')



sh = sh[['Sold To','Sold To Name','SD Document Type','Sales Order/PO #','Item #','Ship Date','Create Date',
         'Request Date','Material Number','Material Description','Material type','Material type description',
         'Material group','Quantity','Total Sale','Ship to Country','Profit Center','Sales Distribution Channel',
         'Sales Division','Sales Doc. Item Cat.']]

sh = sh.rename(columns={'Sold To':'Customer (Sold To)','Sold To Name':'Customer Name (Sold To)',
                        'Material type':'Material Type','Material group':'Material Group',
                    'SD Document Type':'Sales Document Type','Sales Order/PO #':'Sales Document',
                    'Item #':'Sales Document Item','Total Sale':'Sales Amount','Ship to Country':'Country (Ship To)',
                       'Sales Distribution Channel':'Distribution Channel','Sales Division':'Division',
                      'Sales Doc. Item Cat.':'Item Category'})
    
sh = sh[sh['Ship Date'].notnull()]

bl['Order Status'] = 'Open'

sh['Order Status'] = 'Shipped'

frame = [sh, bl]

df = pd.concat(frame)

df['Create Month'] = pd.DatetimeIndex(df['Create Date']).month
df['Create Quarter'] = "Q" + (pd.DatetimeIndex(df['Create Date']).quarter).astype(str)
df['Create Year'] = pd.DatetimeIndex(df['Create Date']).year

df['Request Month'] = pd.DatetimeIndex(df['Request Date']).month
df['Request Quarter'] = "Q" + (pd.DatetimeIndex(df['Request Date']).quarter).astype(str)
df['Request Year'] = pd.DatetimeIndex(df['Request Date']).year

df['Ship/Promise Date'] = np.where(df['Order Status']=='Shipped',df['Ship Date'],df['Request Date'])

df['Ship/Promise Month'] = pd.DatetimeIndex(df['Ship/Promise Date']).month
df['Ship/Promise Quarter'] = "Q" + (pd.DatetimeIndex(df['Ship/Promise Date']).quarter).astype(str)
df['Ship/Promise Year'] = pd.DatetimeIndex(df['Ship/Promise Date']).year

df['VS']= 'OHWV'

#df.to_excel("OHWV Raw Data Jul 15_copy.xlsx")

writer = pd.ExcelWriter('OHWV Raw Data Oct 7.xlsx',engine='xlsxwriter',datetime_format='mm/dd/yyyy',date_format='mm/dd/yyyy')
df.to_excel(writer,sheet_name='Sheet1',index=False)
writer.save()
