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
    ticker = symbol.split('.')[0]  # e.g. IRI.AU
    print('Getting volume data for: ' + ticker)
    print(exc_list)
    # start date is 3 month before end date
    start_date = dt.datetime.strptime(date, "%d-%b-%y") - dateutil.relativedelta.relativedelta(months=3)
    start_date = str(int(start_date.timestamp()))
    end_date = dt.datetime.strptime(date, "%d-%b-%y")  # previous: - dateutil.relativedelta.relativedelta(months=1)
    end_date = str(int(end_date.timestamp()))

    # get url for "download data" using regular expression(crumb characters are dynamic)
    session = requests.Session()
    url = 'https://finance.yahoo.com/quote/%s/history?p=%s' % (symbol, symbol)
    page = session.get(url).content.decode()
    pattern = re.compile('{"crumb":"(.{11})"}')  # .{11} Matches exactly 11 consecutive characters.
    try:
        crumb = re.findall(pattern, page)[0]
    except IndexError as a:
        return a
    print(crumb)

    url_root1 = 'https://query1.finance.yahoo.com/v7/finance/download/'
    url_root2 = '?period1=%s&period2=%s&interval=1d&events=history&crumb=%s' % (start_date, end_date, crumb)
    count = 0
    avg_10d_vol = []
    med_10d_vol = []
    med_3m_vol = []
    avg_3m_vol = []
    for i in range(len(exc_list)):   # from 0 to len-1

        url_root = url_root1 + ticker + exc_list[i] + url_root2
        # check whether the url is valid
        try:
            response = session.get(url_root)
        except:
            try:
                time.sleep(5)
                response = session.get(url_root)
            except Exception as e:
                if i == (len(exc_list)-1):
                    if hasattr(e, 'reason'):
                        return e.reason
                    elif hasattr(e, 'code'):
                        return e.code
                else:
                    continue
        if response.status_code == 200:
            count += 1
            t = ticker + exc_list[i]
            print("ticker symbol for yahoo", i, t)
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
                med_3m_vol.append(np.median(vol))
                avg_3m_vol.append(np.mean(vol))
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
        m = avg_10d_vol.index(max(avg_10d_vol))
        print(m)
        return avg_10d_vol[m], med_10d_vol[m], avg_3m_vol[m], med_3m_vol[m]





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
    done_sh_Glo = wb.get_sheet_by_name('Global Ratings')
    list_of_rows_Glo = conv_dates_to_rows(pythonLib.root_pandas, 'Global Ratings', list_of_dates)

    # get data
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
        # check whether the url is valid
        while not status and trytime <= 5:
            try:
                avg_10d_vol, med_10d_vol, avg_3m_vol, med_3m_vol = get_volume_data(stock_code, get_date)
                time.sleep(random.randint(1, 5))
                status = True
            except (TypeError, IndexError, ValueError):
                print(stock_code + ':Error')
                trytime += 1
                continue
        done_sh['AK' + str(index)] = avg_3m_vol
        done_sh['AL' + str(index)] = avg_10d_vol
        done_sh['AM' + str(index)] = med_10d_vol
        done_sh['AN' + str(index)] = med_3m_vol

    wb.save(main_file)

    BACKUP_FILENAME = str(dt.datetime.today().date()) + ' backup.xlsx'
    backup_path = 'D:\\intern at Mommsen Global\\Project2\\'
    backup = backup_path + BACKUP_FILENAME
    shutil.copy(main_file, backup)


# start and end date (mm-dd-yyyy)
start_date = '05-11-2017'
end_date = '05-20-2017'

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
