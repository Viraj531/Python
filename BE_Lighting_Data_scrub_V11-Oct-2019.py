
import pandas as pd
import numpy as np
import os

os.chdir(r"D:\INT AM Analytics\Data\BE Lighting\2019\10.07.19")

boh_sh = pd.read_excel("Boh_Ship.xls",skiprows=[0,1])
nog_sh = pd.read_excel("Nog_Ship.xls",skiprows=[0,1])
#nb_sh = pd.read_excel("NB_Ship.xlsx",sheet_name='Sheet1'
#                      #,skiprows=[0,1]#
#                     )

boh_bl = pd.read_excel("Boh_Open.xls",skiprows=[0]
                        #sheet_name='Backlog',
                        )
nog_bl = pd.read_excel("Nog_Open.xls",skiprows=[0]
                      #,sheet_name='Backlog'
                      )
#nb_bl = pd.read_excel("NB_Open.xlsx",sheet_name='Sheet1'
#                      #,skiprows=[0,1,2]
#                     )


## NB Shipments 
pip_sh = pd.read_excel("9.30-10.05_NEW BERLIN_Pipeline shipments.xls",skiprows=[0,1,2])
pip_sh =  pip_sh.dropna(how='all')
#nb_pl_sh.to_excel("NB_Pipeline Ship wo blank rows.xlsx")
macro_sh = pd.read_excel("NB_Ship.xlsx",skiprows=[0,1])

macro_sh['Item Number'] = macro_sh['Item Number'].str.strip()
pip_sh['Order No.'] = pip_sh['Order No.'].astype('int64')

#Combining Order No. & Item No. to form keywords for merging
macro_sh['left on'] = (macro_sh['SO Order'].astype(str))  + " " + macro_sh['Item Number']
pip_sh['right on'] = (pip_sh['Order No.'].astype(str)) + " " + pip_sh['Item No']

#pip_sh.rename(columns={'Order No.':'SO Order', 'Item No':'Item Number'},inplace=True)
#pip_sh['SO Order'] = pip_sh['SO Order'].astype(int)

pip_sh = pip_sh.drop_duplicates(subset=['Order No.','Item No'])

NB_Ship = pd.merge(macro_sh,
                   pip_sh,
                   left_on = 'left on',
                   right_on = 'right on',
                   how='left')

NB_Ship['Ship Date_1'] = NB_Ship['Ship Date'].astype(str)
NB_Ship['Create Date'] = NB_Ship['SO Entry Date'].astype(str)
NB_Ship['Request Date'] = NB_Ship['Requested Date'].astype(str)

#NB_Ship['Ship Date'] = NB_Ship['Ship Date'].str[4:6] + "/" + NB_Ship['Ship Date'].str[6:] + "/" + NB_Ship['Ship Date'].str[:4] 
NB_Ship['Ship Date_1'] = pd.to_datetime(NB_Ship['Ship Date_1'])
NB_Ship['Create Date'] = pd.to_datetime(NB_Ship['Create Date'])
NB_Ship['Request Date'] = pd.to_datetime(NB_Ship['Request Date'])

## NB Backlogs
pip_bl = pd.read_excel("10.07_NEW BERLIN_Pipeline backlog.xls",skiprows=[0,1,2])
pip_bl =  pip_bl.dropna(how='all')
#pip_bl.to_excel("NB_Pipeline Open wo blank rows.xlsx")
macro_bl =  pd.read_excel("NB_Open.xlsx",skiprows=[0,1])

macro_bl['item_no'] = macro_bl['item_no'].str.strip()
pip_bl['Order No.'] = pip_bl['Order No.'].astype('int64')

#Combining Order No. & Item No. to form keywords for merging
macro_bl['left on'] = (macro_bl['ord_no'].astype(str))  + " " + macro_bl['item_no']
pip_bl['right on'] = (pip_bl['Order No.'].astype(str)) + " " + pip_bl['Item No']

#pip_bl.rename(columns={'Order No.':'SO Order', 'Item No':'Item Number'},inplace=True)

pip_bl = pip_bl.drop_duplicates(subset=['Order No.','Item No'])

NB_Open = pd.merge(macro_bl,
                   pip_bl,
                   left_on = 'left on',
                   right_on = 'right on',
                   how='left')

NB_Open['Create Date'] = NB_Open['entered_dt'].astype(str)
NB_Open['Request Date'] = NB_Open['req_ship_dt'].astype(str)

NB_Open['Create Date'] = pd.to_datetime(NB_Open['Create Date'])
NB_Open['Request Date'] = pd.to_datetime(NB_Open['Request Date'])

boh_sh = boh_sh[['Order Type', 'SO Order','SO LI', 'RMA','Loc', 'Cust No', 'Cust Name', 'Item Number',
       'Item Description', 'Prod Cat', 'Cust Type', 'Profit Center','Program', 'Model',
       'Qty Shipped', 'Unit Price', 'Ext Price','Ship Date','SO Entry Date','Req Date']]
boh_sh = boh_sh.rename(columns={'Prod Cat':'Product Category','Ext Price':'Sales Amount','SO Entry Date':'Create Date','Req Date':'Request Date'})
boh_sh['Order Status'],boh_sh['Site']=['Shipped','Bohemia']

nog_sh = nog_sh[['Order Type', 'SO Order',
       'SO LI', 'Loc', 'Cust No', 'Cust Name', 'Item Number',
       'Item Description', 'Prod Cat', 'Cust Type', 'Profit Center','Program', 'Model','RMA',
       'Qty Shipped', 'Unit Price', 'Ext Price','Ship Date','SO Entry Date','Req Date']]
nog_sh = nog_sh.rename(columns={'Prod Cat':'Product Category','Ext Price':'Sales Amount','SO Entry Date':'Create Date','Req Date':'Request Date'})
nog_sh['Order Status'],nog_sh['Site']=['Shipped','Nogales']

nb_sh = NB_Ship[['Ord Type', 'SO Order','SO LI',
       'Customer No', 'Loc', 'Customer Name', 'Item Number',
       'Item Description', 'Prod Cat', 'Customer Type', 'Profit Center','Program',
       'Model','RMA','Unit', 'Qty Shipped','Ship Date_1', 'Unit Price',
       'Ext$-Disc','Create Date','Request Date','Product Bucket']]
nb_sh = nb_sh.rename(columns={'Ord Type':'Order Type','Customer No':'Cust No','Customer Name':'Cust Name','Customer Type':'Cust Type','Prod Cat':'Product Category','Ext$-Disc':'Sales Amount','Ship Date_1':'Ship Date'})
nb_sh['Order Status'],nb_sh['Site']=['Shipped','New Berlin']

boh_bl = boh_bl[['Order Type', 'Order No','RMA No',
       'Line No','Profit Center', 'Product Cat',
       'Cust Type', 'Cust No', 'Bill To Name',
       'Req Ship Date', 'Entered Date', 'Item No',
       'Item Description', 'Qty Ordered', 'Unit Price', 'Ext Price',
       'Model','Project']]
boh_bl = boh_bl.rename(columns={'Product Cat':'Product Category','Order No':'SO Order','RMA No':'RMA','Line No':'SO LI','Bill To Name':'Cust Name','Req Ship Date':'Request Date','Ext Price':'Sales Amount','Entered Date':'Create Date','Item No':'Item Number','Qty Ordered':'Qty Shipped','Project':'Program'})
boh_bl['Order Status'],boh_bl['Site']=['Open','Bohemia']


nog_bl = nog_bl[['Order Type', 'Order No','RMA No',
       'Line No','Profit Center', 'Product Cat',
       'Cust Type', 'Cust No', 'Bill To Name',
       'Req Ship Date', 'Entered Date', 'Item No',
       'Item Description', 'Qty Ordered', 'Unit Price', 'Ext Price',
       'Model','Project']]
nog_bl = nog_bl.rename(columns={'Product Cat':'Product Category','Order No':'SO Order','RMA No':'RMA','Line No':'SO LI','Bill To Name':'Cust Name','Req Ship Date':'Request Date','Ext Price':'Sales Amount','Entered Date':'Create Date','Item No':'Item Number','Qty Ordered':'Qty Shipped','Project':'Program'})
nog_bl['Order Status'],nog_bl['Site']=['Open','Nogales']

nb_bl = NB_Open[['ord_type', 'ord_no','rma_no','line_no','cus_no', 'bill_to_name', 'item_no',
       'item_desc_1', 'prod_cat', 'customer_type', 'profit_center','Unit','Request Date','qty_ordered', 'unit_price',
       'ext_price','Create Date','Product Bucket']]
nb_bl = nb_bl.rename(columns={'ord_type':'Order Type','ord_no':'SO Order','rma_no':'RMA','line_no':'SO LI','item_no':'Item Number','item_desc_1':'Item Description','cus_no':'Cust No','bill_to_name':'Cust Name','prod_cat':'Product Category','customer_type':'Cust Type','profit_center':'Profit Center','ext_price':'Sales Amount','unit_price':'Unit Price','qty_ordered':'Qty Shipped'})
nb_bl['Order Status'],nb_bl['Site']=['Open','New Berlin']


frame1 = [boh_sh, nog_sh,
          boh_bl,nog_bl,
          nb_sh, nb_bl
         ]


df1 = pd.concat(frame1)


#For df1 frame containing Boh & Nog

df1['Create Month'] = pd.DatetimeIndex(df1['Create Date']).month
df1['Create Quarter'] = "Q" + (pd.DatetimeIndex(df1['Create Date']).quarter).astype(str)
df1['Create Year'] = pd.DatetimeIndex(df1['Create Date']).year

df1['Request Month'] = pd.DatetimeIndex(df1['Request Date']).month
df1['Request Quarter'] = "Q" + (pd.DatetimeIndex(df1['Request Date']).quarter).astype(str)
df1['Request Year'] = pd.DatetimeIndex(df1['Request Date']).year

df1['Ship/Promise Date'] = np.where(df1['Order Status']=='Shipped',df1['Ship Date'],df1['Request Date'])

df1['Ship/Promise Month'] = pd.DatetimeIndex(df1['Ship/Promise Date']).month
df1['Ship/Promise Quarter'] = "Q" + (pd.DatetimeIndex(df1['Ship/Promise Date']).quarter).astype(str)
df1['Ship/Promise Year'] = pd.DatetimeIndex(df1['Ship/Promise Date']).year

df1['VS'] = "BE Lighting"

df1['Cust Name'] = df1['Cust Name'].str.strip()

writer = pd.ExcelWriter('BE Raw Data Oct 7.xlsx',engine='xlsxwriter',datetime_format='mm/dd/yyyy',date_format='mm/dd/yyyy')
df1.to_excel(writer,sheet_name='Sheet1',index=False)
writer.save()

#df2.to_excel("NB _Wk of Jul 01.xlsx")


