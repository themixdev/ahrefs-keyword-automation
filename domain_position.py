import os
import re
import shutil
import gspread
import pandas as pd
from time import sleep
import undetected_chromedriver
from win32com.client import Dispatch
from main import get_latest_download, login
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from oauth2client.service_account import ServiceAccountCredentials


def get_version_via_com(filename):
    parser = Dispatch("Scripting.FileSystemObject")
    try:
        version = parser.GetFileVersion(filename)
    except Exception:
        return None
    return version


def get_google_sheet(url):
    # Use ServiceAccountCredentials to authenticate
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name('google_sheet.json', scope)
    client = gspread.authorize(creds)
    # Open the spreadsheet

    sheet = client.open_by_url(
        url).sheet1
    return sheet


def make_dirs():
    # Make dir "Temp" in current working directory
    parent_dir = os.getcwd()
    # set path and path to file
    path_to_temp = os.path.join(parent_dir, "Temp")
    path_to_downloads = os.path.join(parent_dir, "Downloads")
    # add option to download to given folder

    try:
        shutil.rmtree(path_to_temp)
        shutil.rmtree(path_to_downloads)
    except:
        print("Temp does not exists")
    os.mkdir(path_to_temp)
    os.mkdir(path_to_downloads)

    return path_to_temp, path_to_downloads


def write_to_sheet(df, arr):
    # global ro/w
    for i, j in enumerate(arr):
        col = 7 + i
        curr_url = df['Current URL'].iloc[j]
        sheet.update_cell(row, col, curr_url)
        curr_pos = df['Current position'].iloc[j]
        col += 3
        sheet.update_cell(row, col, curr_pos)
        keyword = df['Keyword'].iloc[j]
        col += 3
        sheet.update_cell(row, col, keyword)


def random_rows(row_len):
    rang = range(0, row_len)
    if row_len > 3:
        import random
        rand_col_ind = random.sample(rang, k=3)
        return rand_col_ind
    else:
        return rang


def modify_csv(domain):
    # url_pattern = re.compile(fr'(https://[www]?.{domain}+/[\w-]+)')
    url_pattern = re.compile(fr'(https://(www\.)?{domain}+/[\w-]+.*)')
    csv_file = get_latest_download()
    # file = pd.read_excel(csv_file)
    df = pd.read_csv(csv_file, error_bad_lines=False)
    df = df[['Keyword', 'Current URL', 'KD', 'CPC', 'Current organic traffic', 'Current position', 'Volume']]
    # drop rows with missing values
    df = df.dropna()
    try:
        df = df[~df['Current URL'].str.contains('contact')]
        df = df[~df['Current URL'].str.contains('about')]
    except:
        pass
    # filter the data
    df_tmp = df[(df['Volume'] > 10) &
                (df['KD'] >= 10) &
                (df['KD'] < 50) &
                (df['CPC'] > 1) &
                (df['Current organic traffic'] > 4) &
                (df['Current position'] > 10) &
                (df['Current URL'].str.match(url_pattern))]
    df_tmp.drop_duplicates(subset='Current URL', inplace=True)
    # print the filtered data
    # print("Filtered file is:\n", df_tmp)

    if df_tmp.empty:
        df_tmp2 = df[(df['Volume'] > 5) &
                    (df['KD'] >= 5) &
                    (df['KD'] < 80) &
                    (df['CPC'] > 0.2) &
                    (df['Current organic traffic'] > 1) &
                    (df['Current position'] > 6) &
                    (df['Current URL'].str.match(url_pattern))]
        df_tmp2.drop_duplicates(subset='Current URL')
        # print("LAST ONE file is \n", df_tmp2)
        if df_tmp2.empty:
            df.drop_duplicates(subset='Current URL')
            # df = df.query('Current URL.str.match(@url_pattern)')
            df = df[df['Current URL'].str.match(url_pattern,na=False)]
    else:
        df = df_tmp
    rand_rows = random_rows(len(df.index))
    write_to_sheet(df[['Keyword', 'Current position', 'Current URL']], rand_rows)
    os.remove(csv_file)


def read_column_from_google_sheet():
    # Use ServiceAccountCredentials to authenticate

    # Get the column by name
    column_index = sheet.row_values(1).index('Domains') + 1
    column = sheet.col_values(column_index)[1:]

    # Return the values as a tuple
    return tuple(column)


def export_excel(driver, url):
    driver.get(url)
    driver.get(url)
    # os.system("pause")
    sleep(8)
    try:
        export_btn = WebDriverWait(driver, 30).until(EC.presence_of_element_located(
            (By.XPATH,
             "//button[normalize-space()='Export']")))
        export_btn.click()
        sleep(6)
        # change to UTF-8
        utf8_checkbox = WebDriverWait(driver, 30).until(EC.presence_of_element_located(
            (By.XPATH,
             "//span[contains(text(), 'For OpenOffice, Libre and others (UTF-8)')]")))
        utf8_checkbox.click()
        # driver.find_element(By.XPATH, "//span[contains(text(), 'For OpenOffice, Libre and others (UTF-8)')]")
        # sleep(1)
        driver.find_elements(By.XPATH, "//button[normalize-space()='Export']")[1].click()
        sleep(10)
    except Exception as err:
        print(err)
        print(url)


if __name__ == '__main__':
    sheet = get_google_sheet(
        'https://docs.google.com/spreadsheets/d/1MzLcEKBu1M4tDyrvVFTy9BVn0ycYX1Y8oFNkIsjItfw/edit#gid=0')

    path_to_temp, path_to_download = make_dirs()
    #########
    paths = [r"C:\Program Files\Google\Chrome\Application\chrome.exe",
             r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"]
    version = list(filter(None, [get_version_via_com(p) for p in paths]))[0]
    chrome_version = version.split(".")[0]

    driver = undetected_chromedriver.Chrome(version_main = chrome_version)
        # login to ahrefs
    login(driver)
        # Find the export button
    domains = read_column_from_google_sheet()
    for row, domain in enumerate(domains):
        try:
            row += 2
            url = 'https://app.ahrefs.com/v2-site-explorer/organic-keywords/subdomains?compareData=dontCompare&compareDate=prev3Months&country=us&currentDate=today&isUserCountry=1&limit=50&&offset=0&sort=trafficInitial&sortDirection=desc&target='
            full_url = url + domain
            export_excel(driver, full_url)
            modify_csv(domain)
        except Exception as err:
            print("Error excepted for domain: ", domain, "Error \n", err)

    driver.close()
    driver.quit()
