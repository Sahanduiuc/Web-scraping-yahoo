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
# Exchange Parameter added, applies to Non-US stock
def get_beta(symbol):
    """
    input: ticker

    output: beta and average 3month volume
    """

    exc_list = pythonLib.Exc_map_3.get(pythonLib.get_country_name(symbol), [''])
    beta = 'null'
    ticker = symbol.split('.')[0]
    print('Getting beta for: ' + ticker)

    for i in range(len(exc_list)):
        url_root = 'http://finance.yahoo.com/quote/'
        url_root += ticker + exc_list[i] + '/?p=' + ticker + exc_list[i]
        status = False
        trytime = 0
        page = None
        # check whether the url is valid
        while not status and trytime <= 5:
            try:
                # page = urllib2.urlopen(url_root, timeout=10)
                page = requests.get(url_root, timeout=30)
                # print(url_root)
                status = True
            except (requests.HTTPError, requests.ConnectionError):
                break
            except:
                trytime += 1
                print('timeout')
                time.sleep(60)
        if not page:
            continue
        # c = page.read().decode('utf-8')
        c = page.text
        soup = BeautifulSoup(c, "html5lib")
        page.close()
        table = soup.findAll('table')
        headings = [td.get_text() for i in table for td in i.find_all('td')]
        if 'Beta' in headings:
            beta_index = headings.index('Beta')
            beta = headings[beta_index + 1]
            if beta != 'N/A' and beta != 'null':
                break
            elif beta != 'N/A':
                beta = 'null'
    return beta


# Function to get Beta and average 3 month volume
# Do NOT consider exchange, applies to US stock
def get_beta_2(symbol):
    """
    input: ticker

    output: beta and average 3month volume
    """

    ticker = symbol.split('.')[0]
    print('Getting beta for: ' + ticker)
    url_root = 'http://finance.yahoo.com/quote/'
    url_root += ticker + '/?p=' + ticker
    #print(url_root)
    beta = 'null'
    status = False
    trytime = 0
    page = None
    # check whether the url is valid
    while not status and trytime <= 5:
        try:
            # page = urllib2.urlopen(url_root, timeout=10)
            page = requests.get(url_root, timeout=30)
            status = True
        except (requests.HTTPError, requests.ConnectionError):
            return 'null'
        except:
            trytime += 1
            print('timeout')
            time.sleep(60)
    if not page:
        return 'null' 'null'
    # c = page.read().decode('utf-8')
    c = page.text
    soup = BeautifulSoup(c, "html5lib")
    page.close()
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
def insert_data_sheet_date(list_of_dates):
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
    list_to_use = []

    # get data
    for done_sh in [done_sh_Amr, done_sh_Glo]:

        if done_sh == done_sh_Amr:
            Ticker_col = 'C'
            list_to_use = list_of_rows_Amr
        elif done_sh == done_sh_Glo:
            Ticker_col = 'F'
            list_to_use = list_of_rows_Glo

        for index in list_to_use:

            stock_code = str(done_sh[Ticker_col + str(index)].value)
            get_date = str(done_sh['A' + str(index)].value)
            get_date = pd.to_datetime(get_date).strftime('%d-%b-%y')

            print('(Stock: ' + stock_code + ', Date: ' + get_date + ')')

            # get data and write data into excel
            if Ticker_col == 'C':
                beta = get_beta_2(stock_code)
                done_sh['AJ' + str(index)] = beta

            elif Ticker_col == 'F':
                beta = get_beta(stock_code)
                done_sh['AP' + str(index)] = beta

    wb.save(main_file)

    BACKUP_FILENAME = str(dt.datetime.today().date()) + ' backup.xlsx'
    backup_path = 'D:\\intern at Mommsen Global\\Project2\\'
    #backup_path = 'F:\\Data_collection_system\\Main Folder\\Street Account US & Euro\\Street_Act_price_backup\\'
    backup = backup_path + BACKUP_FILENAME
    shutil.copy(main_file, backup)


#######################  main  #######################

# start and end date (mm-dd-yyyy)
start_date = '06-1-2017'
end_date = '06-1-2017'

Date_list = []
for My_date in pd.date_range(start_date, end_date):
    Date_list.append(str(My_date.date().strftime('%d-%b-%y')))

try:
    # Call helper function to write into excel
    insert_data_sheet_date(Date_list)
    print("Completed")
except Exception as e:
    print("Cannot write into excel", )
    print(e, sys.exc_info())
