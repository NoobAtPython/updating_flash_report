# Script downloads the latest prepaid card sales from ftp site, order summary from third party vendor site,
# and advertisement revenue from google sheets. It also updates the flash report.

import requests
import glob
import os
import time
import win32com.client
from datetime import datetime, timedelta
from gsheets import Sheets

# function logs into ftp website with credentials and downloads each CSV file until today's date.
def download_ftp_file():
    list_of_files = glob.glob("C:/Users/BReyes/Desktop/Bryans_Folder/Game/PrepaidCards/*")
    latest_file = max(list_of_files, key=os.path.getctime)
    previous_file = latest_file[87:-4]
    previous_file = datetime(year=int(previous_file[0:4]),month=int(previous_file[4:6]), day =int(previous_file[6:8])) + timedelta(days=1)
    login_url = "https://ftp.website.com"

# login credentials to access ftp website.
    payload = {
        "username": "userid",
        "password": "secret_password",
    }

# posts login credentials to ftp website login form.
    with requests.session() as session:
         session.post(login_url, data=payload)

# downloads each csv file until today's date. if this is not included, then a URL past today's date would be constructed.
# an empty CSV file would then be downloaded until the script is manually stopped.
    while previous_file <= datetime.today():
        todays_date = previous_file.strftime('%Y%m%d')
        csv_download = "https://ftp.website.com/" + str(todays_date) + ".csv"
        folder_path = "C:/Users/BReyes/Desktop/Bryans_Folder/Games/DailySales" + str(todays_date) + ".csv"

        download = session.get(csv_download)
        with open(folder_path, 'wb') as my_csv:
            my_csv.write(download.content)
        previous_file += timedelta(days=1)

    print('FTP files downloaded and saved!')


# downloads order summary excel file which contains a daily summary of all placed orders (subscriptions and game currency).
def download_order_summary():
    yesterday = (datetime.today()-timedelta(days = 1)).strftime('%Y-%m-%d')
    form_url = "https://admin.website.com/admin/login"
    login_url = "https://admin.website.com/admin/admin_login_check"
    folder_path = "C:/Users/BReyes/Desktop/Bryans_Folder/Game/Order Summary/OrderSummary.xlsx"
    download_url = "https://admin.website.com/admin/order_summary/export?form%5BstartDate%5D=2018-01-01&form%5BendDate%5D=" + yesterday + "&format=xls"

# credentials to use to successfully login to website.
    payload = {
        "_target_path": "https://admin.website.com/admin/dashboard",
        "_username": "userid",
        "_password": "secret_password",
        "_remember_me": "on",

    }

# logs into website and downloads order summary file.
    with requests.session() as session:
        session.get(form_url)
        session.post(login_url, data=payload, headers={'Referer': form_url})
        download = session.get(download_url)

# saves order summary file as excel file.
    with open(folder_path, 'wb') as excel_file:
        excel_file.write(download.content)

    print('Order summary file downloaded and saved!')


# downloads advertisement revenue from google sheets. saves it as an excel file.
# updates, transforms, and saves a connected excel workbook into a pdf, which pulls data from advertisement file.
def download_and_update_from_google_sheets():
    sheets = Sheets.from_files('~/client_secrets.json','~/storage.json')
    url = 'https://docs.google.com/spreadsheets/d/idnumberhere'
    s = sheets.get(url)
    s.sheets[2].to_csv('C:/Users/BReyes/Desktop/Bryans_Folder/Game/Ad Revenue/AdDailyRevenue2019.csv', encoding='utf-8', dialect='excel')

    xl_app = win32com.client.DispatchEx("Excel.Application")
    wb = xl_app.workbooks.open('C:/Users/BReyes/Desktop/Bryans_Folder/Analytics/Flash Reports/Current/Neopets Flash Report.xlsx')
    wb.RefreshAll()
    ws = wb.Worksheets[0]
    ws.Visible = 1
    print_area = 'A1:K42'
    ws.PageSetup.Zoom = False
    ws.PageSetup.FitToPagesTall = 1
    ws.PageSetup.FitToPagesWide = 1
    ws.PageSetup.PrintArea = print_area
    time.sleep(15)
    wb.Save()
    ws.ExportAsFixedFormat(0, 'C:/Users/BReyes/Desktop/Bryans_Folder/Analytics/Flash Reports/Current/Flash Report.pdf' )
    xl_app.Quit()

    print('Google sheets file downloaded and saved! Flash report has been updated.')


download_ftp_file()
download_order_summary()
download_and_update_from_google_sheets()
