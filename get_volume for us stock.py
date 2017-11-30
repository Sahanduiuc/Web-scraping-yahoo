import pandas as pd
import numpy as np
import datetime as dt
import dateutil.relativedelta
import requests
import time
import re
import openpyxl
import sys
import pythonLib  # user-defined
import random


# Helper function to get historical daily volume data (raw data)
# Do NOT consider exchange, applies to US stock
# Calculate average/medain 10 days and 3m volume
def get_volume_data_2(symbol, date):
    """
    input: ticker, end date

    output: average/median 10 days and 3m volume
    """

    print('Getting volume data for: ' + symbol)
    # start date is 3 month before date
    start_date = dt.datetime.strptime(date, "%d-%b-%y") - dateutil.relativedelta.relativedelta(months=3)
    start_date = str(int(start_date.timestamp()))
    end_date = dt.datetime.strptime(date, "%d-%b-%y")
    end_date = str(int(end_date.timestamp()))

    # get url for "download data" using regular expression(crumb characters are dynamic)
    session = requests.Session()
    url = 'https://finance.yahoo.com/quote/%s/history?p=%s' % (symbol, symbol)
    page = session.get(url).content.decode()
    pattern = re.compile('{"crumb":"(.{11})"}')
    #pattern = re.compile('{"crumb":"(\S+?)"}')
    #pattern = re.compile('"CrumbStore":{"crumb":"(.+?)"}')  # ('{"crumb":"({.+})"}')
    #pattern = re.compile('{"user":{"crumb":"(.+)","firstName"')
    try:
        crumb = re.findall(pattern, page)[0]
    except IndexError as a:
        return a
    print(crumb)
    url = 'https://query1.finance.yahoo.com/v7/finance/download/' \
              '%s?period1=%s&period2=%s&interval=1d&events=history&crumb=%s' % (
        symbol, start_date, end_date, crumb)

    try:
        response = session.get(url)
    except:
        try:
            time.sleep(5)
            response = session.get(url)
        except Exception as e:
            if hasattr(e, 'reason'):
                return e.reason
            elif hasattr(e, 'code'):
                return e.code
    print(response.status_code)
    data = response.content.decode().splitlines()
    # get the volume data
    vol = []
    for lines in data[1:]:
        temp = lines.split(',')[6]
        if temp != 'null':
            vol.append(float(temp))
    # check whether has enough data
    if len(vol) <= 10:
        print("Not enough data")
        return 'null', 'null', 'null', 'null'
    avg_10d_vol = np.mean(vol[:10])
    med_10d_vol = np.median(vol[:10])
    if len(vol) <= 60:
        print("Not enough data")
        return avg_10d_vol, med_10d_vol, 'null', 'null'
    med_3m_vol = np.median(vol)
    avg_3m_vol = np.mean(vol)
    return avg_10d_vol, med_10d_vol, avg_3m_vol, med_3m_vol



# Function to convert list_of_dates into list_of_rows
def conv_dates_to_rows(path_file, sheet_name, list_of_dates):
    """
    input: output file name (directory convention a little bit different)
           name of target sheet
           list of dates to be converted
    output: of list of row numbers (Type: Long)
    """
    df = pd.read_excel(path_file, sheet_name, usecols=[0])
    df['US Date'] = pd.to_datetime(df['US Date'])
    df['US Date'] = df['US Date'].apply(lambda x: x.strftime('%d-%b-%y') if not pd.isnull(x) else '')
    bool_list = []

    for my_date in list_of_dates:
        if not bool_list:
            bool_list = (df['US Date'] == my_date).tolist()
        else:
            bool_list = np.logical_or(bool_list, (df['US Date'] == my_date).tolist()).tolist()

    my_list = df[bool_list].index.tolist()
    my_list = (np.array(my_list) + 2).tolist()

    return my_list


# Function to write data into the file
def insert_data_sheet_date(list_of_dates):
    """
    input: date of current row
    """
    print('Inserting data into file:')
    main_file = pythonLib.root_out + pythonLib.OUTPUT_FILENAME
    wb = openpyxl.load_workbook(main_file)
    done_sh_Amr = wb.get_sheet_by_name('Amr Ratings')
    done_sh = done_sh_Amr
    list_of_rows_Amr = conv_dates_to_rows(pythonLib.root_pandas, 'Amr Ratings', list_of_dates)

    # get data
    Ticker_col = 'C'
    list_to_use = list_of_rows_Amr

    for index in list_to_use:

        stock_code = str(done_sh[Ticker_col + str(index)].value)
        get_date = str(done_sh['A' + str(index)].value)
        get_date = pd.to_datetime(get_date).strftime('%d-%b-%y')

        print('(Stock: ' + stock_code + ', Date: ' + get_date + ')')

        status = False
        trytime = 0
        # check whether the url is valid
        while not status and trytime <= 5:
            try:
                avg_10d_vol, med_10d_vol, avg_3m_vol, med_3m_vol = get_volume_data_2(stock_code, get_date)
                time.sleep(random.randint(1, 5))
                status = True
            except (TypeError, IndexError, ValueError):
                print(stock_code + ':Error')
                trytime += 1
                continue

        done_sh['AE' + str(index)] = avg_3m_vol
        done_sh['AF' + str(index)] = avg_10d_vol
        done_sh['AG' + str(index)] = med_10d_vol
        done_sh['AH' + str(index)] = med_3m_vol

    wb.save(main_file)


# start and end date (mm-dd-yyyy)
start_date = '06-13-2017'
end_date = '06-20-2017'

Date_list = []
for My_date in pd.date_range(start_date, end_date):
    Date_list.append(str(My_date.date().strftime('%d-%b-%y')))

try:
    # Call helper function to write into excel
    insert_data_sheet_date(Date_list)
    print("Completed")
except Exception as e:
    print("Cannot write into excel",)
    print(e, sys.exc_info())