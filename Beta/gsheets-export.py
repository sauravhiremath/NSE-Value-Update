# -*- coding: utf-8 -*-
"""
Created on Thu Nov 22 15:18:29 2018

@author: Saurav
"""

def sheetExport(df):
    #authorization
    gc = pygsheets.authorize(service_file='/Users/Saurav/Desktop/NSE-Index-8086c5d81c3e.json')

    #open the google spreadsheet (where 'PY to Gsheet Test' is the name of my sheet)
    sh = gc.open('NSE-Index')

    #select the first sheet 
    wks = sh[0]

    #update the first sheet with df, starting at cell B2. 
    wks.set_dataframe(df_updated, 
                    start = (1,1),
                    copy_head = False )
