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
import shutil


# Helper function to get historical daily volume data (raw data)
# Exchange Parameter added, applies to Non-US stock
# Calculate average 10 days volume, median 10 days volume and median 3 month volume
def get_volume_data(symbol, date):
    """
    input: ticker, end date

    output: average 10 days volume, median 10 days volume and median 3 month volume
    """

    # find corresponding exchange code
    exc_list = pythonLib.Exc_map_3.get(pythonLib.get_country_name(symbol), [''])
    ticker = symbol.split('.')[0]
    print('Getting volume data for: ' + ticker)
    # start date is 5 month before date,to ensure getting data for 60 trading days
    start_date = dt.datetime.strptime(date, "%d-%b-%y") - dateutil.relativedelta.relativedelta(months=5)
    start_date = str(int(start_date.timestamp()))
    end_date = dt.datetime.strptime(date, "%d-%b-%y")
    end_date = str(int(end_date.timestamp()))

    # get url for "download data" using regular expression("crumb" characters are dynamic)
    session = requests.Session()

    count = 0
    avg_10d_vol = []
    med_10d_vol = []
    med_3m_vol = []
    avg_3m_vol = []
    for i in range(len(exc_list)):  # from 0 to len-1
        new_ticker = ticker + exc_list[i]
        url = 'https://finance.yahoo.com/quote/%s/history?p=%s' % (new_ticker, new_ticker)
        page = session.get(url).content.decode()
        pattern = re.compile('{"crumb":"(.{11})"}')  # .{11} Matches exactly 11 consecutive characters.
        try:
            crumb = re.findall(pattern, page)[0]
        except IndexError as a:
            return a

        url_root1 = 'https://query1.finance.yahoo.com/v7/finance/download/'
        url_root2 = '?period1=%s&period2=%s&interval=1d&events=history&crumb=%s' % (start_date, end_date, crumb)
        url_root = url_root1 + new_ticker + url_root2
        # check whether the url is valid
        try:
            response = session.get(url_root)
        except:
            try:
                time.sleep(5)
                response = session.get(url_root)
            except Exception as e:
                if i == (len(exc_list) - 1):
                    if hasattr(e, 'reason'):
                        return e.reason
                    elif hasattr(e, 'code'):
                        return e.code
                else:
                    continue
        if response.status_code == 200:
            count += 1
            data = response.content.decode().splitlines()
            vol = []
            for lines in data[1:]:
                temp = lines.split(',')[6]
                if temp != 'null' and temp is not None:
                    vol.append(float(temp))
            # check whether has enough data
            if len(vol) <= 10:
                print("Not enough data")
                avg_10d_vol.append('null')
                med_10d_vol.append('null')
                med_3m_vol.append('null')
                avg_3m_vol.append('null')
            elif len(vol) > 60:
                avg_10d_vol.append(np.mean(vol[:10]))
                med_10d_vol.append(np.median(vol[:10]))
                med_3m_vol.append(np.median(vol[:60]))
                avg_3m_vol.append(np.mean(vol[:60]))
            else:
                print("Not enough data")
                avg_10d_vol.append(np.mean(vol[:10]))
                med_10d_vol.append(np.median(vol[:10]))
                med_3m_vol.append('null')
                avg_3m_vol.append('null')

    if count == 0:
        print("cannot find ticker")
        return 'null', 'null', 'null', 'null'
    elif count == 1:
        return avg_10d_vol[0], med_10d_vol[0], avg_3m_vol[0], med_3m_vol[0]
    else:
        m = avg_10d_vol.index(max(avg_10d_vol))  # use max(avg_10d_vol) as a reference to find the primary exchange code
        return avg_10d_vol[m], med_10d_vol[m], avg_3m_vol[m], med_3m_vol[m]


# Helper function to get historical daily volume data (raw data)
# Do NOT consider exchange, applies to US stock
# Calculate average/medain 10 days and 3m volume
def get_volume_data_2(symbol, date):
    """
    input: ticker, end date

    output: average/median 10 days and 3m volume
    """

    print('Getting volume data for: ' + symbol)
    # start date is 5 month before date,to ensure getting data for 60 trading days
    start_date = dt.datetime.strptime(date, "%d-%b-%y") - dateutil.relativedelta.relativedelta(months=5)
    start_date = str(int(start_date.timestamp()))
    end_date = dt.datetime.strptime(date, "%d-%b-%y")
    end_date = str(int(end_date.timestamp()))

    # get url for "download data" using regular expression(crumb characters are dynamic)
    session = requests.Session()
    url = 'https://finance.yahoo.com/quote/%s/history?p=%s' % (symbol, symbol)
    page = session.get(url).content.decode()
    pattern = re.compile('{"crumb":"(.{11})"}')
    # pattern = re.compile('{"crumb":"(\S+?)"}')
    # pattern = re.compile('"CrumbStore":{"crumb":"(.+?)"}')
    try:
        crumb = re.findall(pattern, page)[0]
    except IndexError as a:  # sometimes crumb would be NONE, so there is IndexError
        return a
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
    med_3m_vol = np.median(vol[:60])
    avg_3m_vol = np.mean(vol[:60])
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




def insert_data_sheet_volume(list_of_dates, action):

    print('Inserting data into file:')
    main_file = pythonLib.root_out + pythonLib.OUTPUT_FILENAME
    wb = openpyxl.load_workbook(main_file)

    done_sh_Amr = wb.get_sheet_by_name('Amr Ratings')
    done_sh_Glo = wb.get_sheet_by_name('Global Ratings')
    list_of_rows_Amr = conv_dates_to_rows(pythonLib.root_pandas, 'Amr Ratings', list_of_dates)
    list_of_rows_Glo = conv_dates_to_rows(pythonLib.root_pandas, 'Global Ratings', list_of_dates)

    if action.lower() == 'both':
        li_st = ['US', "Global"]
    elif action.lower() == 'US':
        li_st = ['US']
    else:
        li_st = ['Global']
    # get data
    for i in li_st:

        if i == 'US':
            done_sh = done_sh_Amr
            Ticker_col = 'C'
            list_to_use = list_of_rows_Amr
        else:
            done_sh = done_sh_Glo
            Ticker_col = 'F'
            list_to_use = list_of_rows_Glo

        for index in list_to_use:

            stock_code = str(done_sh[Ticker_col + str(index)].value)
            get_date = str(done_sh['A' + str(index)].value)
            get_date = pd.to_datetime(get_date).strftime('%d-%b-%y')

            print('(Stock: ' + stock_code + ', Date: ' + get_date + ')')

            status = False
            trytime = 0

            # get data and write data into excel
            if Ticker_col == 'C':
                # check whether the url is valid
                while not status and trytime <= 5:
                    try:
                        avg_10d_vol, med_10d_vol, avg_3m_vol, med_3m_vol = get_volume_data_2(stock_code, get_date)
                        time.sleep(random.randint(1, 5))  ##
                        status = True
                    except (TypeError, IndexError,
                            ValueError):  # important: if get_volume_data_2 returns an error, then try again
                        print(stock_code + ': Error')
                        trytime += 1
                        continue
                done_sh['AE' + str(index)] = avg_3m_vol
                done_sh['AF' + str(index)] = avg_10d_vol
                done_sh['AG' + str(index)] = med_10d_vol
                done_sh['AH' + str(index)] = med_3m_vol

            elif Ticker_col == 'F':
                # check whether the url is valid
                while not status and trytime <= 5:
                    try:
                        avg_10d_vol, med_10d_vol, avg_3m_vol, med_3m_vol = get_volume_data(stock_code, get_date)
                        time.sleep(random.randint(1, 5))
                        status = True
                    except (TypeError, IndexError, ValueError):
                        print(stock_code + ': Error')
                        trytime += 1
                        continue
                done_sh['AK' + str(index)] = avg_3m_vol
                done_sh['AL' + str(index)] = avg_10d_vol
                done_sh['AM' + str(index)] = med_10d_vol
                done_sh['AN' + str(index)] = med_3m_vol

    wb.save(main_file)
'''
    BACKUP_FILENAME = str(dt.datetime.today().date()) + ' backup.xlsx'
    backup_path = 'D:\\intern at Mommsen Global\\Project2\\'
    # backup_path = 'F:\\Data_collection_system\\Main Folder\\Street Account US & Euro\\Street_Act_price_backup\\'
    backup = backup_path + BACKUP_FILENAME
    shutil.copy(main_file, backup)
'''

#######################  main  #######################

# start and end date (mm-dd-yyyy)
start_date = '06-1-2017'
end_date = '06-1-2017'
action = 'Both'  # 'US'  'Both'

Date_list = []
for My_date in pd.date_range(start_date, end_date):
    Date_list.append(str(My_date.date().strftime('%d-%b-%y')))

try:
    # Call helper function to write into excel
    insert_data_sheet_volume(Date_list, action)
    print("Completed")
except Exception as e:
    print("Cannot write into excel", )
    print(e, sys.exc_info())
