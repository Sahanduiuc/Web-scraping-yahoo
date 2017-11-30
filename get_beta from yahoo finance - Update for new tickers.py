import pandas as pd
import numpy as np
import datetime as dt
import openpyxl
import sys
from bs4 import BeautifulSoup
import pythonLib
import requests
import time
import shutil


# Function to get Beta and average 3 month volume
# Do NOT consider exchange, applies to US stock and Non-US stock
def get_beta(symbol):
    """
    input: ticker

    output: beta
    """

    print('Getting beta for: ' + symbol)
    url_root = 'http://finance.yahoo.com/quote/'
    url_root += symbol + '/?p=' + symbol
    beta = 'null'
    status = False
    trytime = 0
    page = None
    # check whether the url is valid
    while not status and trytime <= 5:
        try:
            page = requests.get(url_root, timeout=30)
            status = True
        except (requests.HTTPError, requests.ConnectionError):
            return 'null'
        except:
            trytime += 1
            print('timeout')
            time.sleep(60)
    if not page:
        return 'null'
    soup = BeautifulSoup(page.text, "html5lib")
    table = soup.findAll('table')
    headings = [td.get_text() for i in table for td in i.find_all('td')]
    if 'Beta' in headings:
        beta_index = headings.index('Beta')
        beta = headings[beta_index + 1]
        if beta == 'N/A':
            beta = 'null'
    return beta


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
def insert_data_sheet_date(list_of_dates,action):
    """
    input: date of current row
    """
    print('Inserting data into file:')
    main_file = pythonLib.root_out + pythonLib.OUTPUT_FILENAME
    wb = openpyxl.load_workbook(main_file)
    done_sh_Amr = wb.get_sheet_by_name('Amr Ratings')
    done_sh_Glo = wb.get_sheet_by_name('Global Ratings')
    list_of_rows_Amr = conv_dates_to_rows(pythonLib.root_pandas, 'Amr Ratings', list_of_dates)
    list_of_rows_Glo = conv_dates_to_rows(pythonLib.root_pandas, 'Global Ratings', list_of_dates)

    if action.lower() == 'both':
        li_st = ['US', "Global"]
    elif action.lower() == 'us':
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
            Ticker_col = 'AR'   # on yahoo finance
            list_to_use = list_of_rows_Glo

        for index in list_to_use:

            stock_code = str(done_sh[Ticker_col + str(index)].value)

            if stock_code is None or stock_code == 'null':
                if Ticker_col == 'C':
                    done_sh['AJ' + str(index)] = 'null'
                else:
                    done_sh['AP' + str(index)] = 'null'
                continue

            get_date = str(done_sh['A' + str(index)].value)
            get_date = pd.to_datetime(get_date).strftime('%d-%b-%y')

            print('(Stock: ' + stock_code + ', Date: ' + get_date + ')')

            beta = get_beta(stock_code)
            # get data and write data into excel
            if Ticker_col == 'C':
                done_sh['AJ' + str(index)] = beta

            else:
                done_sh['AP' + str(index)] = beta

    wb.save(main_file)

    BACKUP_FILENAME = str(dt.datetime.today().date()) + ' backup.xlsx'
    backup_path = 'D:\\intern at Mommsen Global\\Project2\\'
    #backup_path = 'F:\\Data_collection_system\\Main Folder\\Street Account US & Euro\\Street_Act_price_backup\\'
    backup = backup_path + BACKUP_FILENAME
    shutil.copy(main_file, backup)


#######################  main  #######################

# start and end date (mm-dd-yyyy)
start_date = '05-27-2017'
end_date = '06-1-2017'
action = 'both'  # 3 choices: 'US'  'Both'  'Global'(case insensitive)

Date_list = []
for My_date in pd.date_range(start_date, end_date):
    Date_list.append(str(My_date.date().strftime('%d-%b-%y')))

try:
    # Call helper function to write into excel
    insert_data_sheet_date(Date_list, action)
    print("Completed")
except Exception as e:
    print("Cannot write into excel", )
    print(e, sys.exc_info())
