from tkinter import *
from datetime import timedelta, datetime
from urllib.request import urlopen, Request, urlretrieve
import urllib
from urllib import request
from pathlib import Path
import urllib.error
from urllib.request import Request, urlopen
import os
import sys
import pandas as pd
import numpy as np
import requests
import csv
import io
import gspread_dataframe as gd
from lxml import html
from lxml import etree
from openpyxl import load_workbook


position_fii, position_dii = 0, 0
workbookPath = 'test.xlsx'


def dii_and_fii_data(date):
    """DIIs and FIIs Data Single Day"""

    # The given url requires date to be in the format ---- ddmmyyyy
    url = 'https://www.nseindia.com/content/nsccl/fao_participant_oi_' + \
        date.replace('-', '') + '.csv'

    hdr = {'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.11 (KHTML, like Gecko) Chrome/23.0.1271.64 Safari/537.11',
           'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
           'Accept-Charset': 'ISO-8859-1,utf-8;q=0.7,*;q=0.3',
           'Accept-Encoding': 'none',
           'Accept-Language': 'en-US,en;q=0.8',
           'Connection': 'keep-alive'}

    r = requests.get(url, headers=hdr)
    df = pd.read_csv(io.StringIO(r.content.decode('utf-8')))

    try:
        new_header = df.iloc[0]  # grab the first row for the header
        df = df[1:]  # take the data less the header row
        df.columns = new_header  # set the header row as the df header

        df.insert(loc=0, column="Date", value=date)

        df_right_dii = df.loc[2:2, ('Future Index Long',
                                    'Future Index Short',
                                    'Future Stock Long',
                                    'Future Stock Short',
                                    'Option Index Call Long',
                                    'Option Index Put Long',
                                    'Option Index Call Short',
                                    'Option Index Put Short')]

        df_right_fii = df.loc[3:3, ('Future Index Long',
                                    'Future Index Short',
                                    'Future Stock Long',
                                    'Future Stock Short',
                                    'Option Index Call Long',
                                    'Option Index Put Long',
                                    'Option Index Call Short',
                                    'Option Index Put Short')]

        df_date = df.loc[2:2, ('Date',)]

    except KeyError:
        print("[+] Sorry, content for %s is not available online,\nKindly try after 7:30 PM for Today's Contents"%(date))
        sys.exit(1)

    return df_date, df_right_dii, df_right_fii


def availableDate(date):
    """Find next available data on site.
This removes the possibilty of holidays in the list.
Returns working day DATE as str
Sub-module: <Only for use with nextDate function>
DO NOT TOUCH"""

    url = 'https://www.nseindia.com/content/nsccl/fao_participant_oi_' + date + '.csv'

    r = requests.get(url)

    tree = html.fromstring(r.content)

    checkDate = tree.findtext('.//title')
    # p Returns None if Data to be scrapped is found
    # p Returns 404 Not Found if Data to be scrapped is not found

    return checkDate


def findDate():
    """Returns the str of Last filled Date and next Date to be filled"""

    global position_dii, position_fii

    df = pd.read_excel(workbookPath)
    lastFilledDate = pd.isna(df['Unnamed: 16']).index[-1]

    # This gives the row index from which data can be started appending
    position_fii = len(df) + 1
    position_dii = len(df) + 1 - 378

    nextDate = (datetime.strptime(lastFilledDate, '%d-%m-%Y') +
                timedelta(days=1)).strftime('%d-%m-%Y')

    while availableDate(nextDate.replace('-', '')) == '404 Not Found':
        nextDate = (datetime.strptime(nextDate, '%d-%m-%Y') +
                    timedelta(days=1)).strftime('%d-%m-%Y')

    return lastFilledDate, nextDate


def niftySpot(date):
    """Returns the nifty closing value of the day as string"""

    # Requires date format to be dd-mm-yyyy
    url = "https://www.nseindia.com/products/dynaContent/equities/indices/historicalindices.jsp?indexType=NIFTY%2050&fromDate=" + date + "&toDate=" + date

    page = requests.get(url)
    tree = html.fromstring(page.content)

    try:
        nifty_close = tree.xpath('/html/body/table/tr/td[5]/text()')[0].strip()
        return nifty_close

    except IndexError:
        print("Sorry the nifty value of %s, has not been refreshed online yet. \nKindly try after 7:30 PM"%(date))
        sys.exit(1)




def dataAppend():

    # lastFilledDate = findDate()[0]

    # now.time() > datetime.time(hour=8)

    while datetime.now().strftime('%d-%m-%Y') != findDate()[0]:

        if datetime.now().strftime('%d-%m-%Y') == findDate()[0]:
            print("[+][+] Process Completed")
            break
        # lastFilledDate = findDate()[1]

        # Load current date inside the variable, thus changing according to the loop of the function
        date = findDate()[1]

        # Load the excel file into the script
        book = load_workbook(workbookPath)
        writer = pd.ExcelWriter(
            workbookPath, engine='openpyxl')
        writer.book = book
        writer.sheets = dict((ws.title, ws) for ws in book.worksheets)

        # Get the value to be appended for a given date in the loop
        df_date, df_right_dii, df_right_fii = dii_and_fii_data(date)
        nifty_close = niftySpot(date)

        # Appending Date to FIIs Data and DIIs Data
        print("[+] Appending Dates of FIIs and DIIs of date %s to row %s (FII) and row %s (DII)" %
              (date, position_fii, position_dii))

        df_date.to_excel(writer, "FII Activity", startrow=position_fii,
                         index=False, header=None)
        df_date.to_excel(writer, "DII", startrow=position_dii,
                         index=False, header=None)

        # Appending FII and nifty information to FIIs Data
        print("[+] Appending Data of FIIs and Nifty of date %s to row %s" %
              (date, position_fii))

        df_right_fii.to_excel(writer, "FII Activity", startrow=position_fii,
                              startcol=14, index=False, header=None)
        pd.DataFrame(data=[nifty_close]).to_excel(
            writer, "FII Activity", startrow=position_fii, startcol=9, index=False, header=None)

        # Appending DII and nifty information to DIIs
        print("[+] Appending Data of DIIs and Nifty of date %s to row %s" %
              (date, position_dii))

        df_right_dii.to_excel(writer, "DII", startrow=position_dii,
                              startcol=12, index=False, header=None)
        pd.DataFrame(data=[nifty_close]).to_excel(
            writer, "DII", startrow=position_dii, startcol=9, index=False, header=None)

        #Saving the excel file
        writer.save()

    print("Seems Done")




def main():
    top = Tk()

    top.title("NSE-Updater")
    top.geometry('500x200')

    index1 = Button(top, text="Update FIIs and DIIs and Nifty Data",
                    bg="black", fg="white", command=dataAppend)
    index1.grid(column=0, row=0, padx=2, pady=2)

    top.mainloop()

    # dataAppend()
    # stop = input()
    # sys.exit(0)


if __name__ == "__main__":
    main()
