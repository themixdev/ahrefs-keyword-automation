import os
import re
import glob
import shutil
import random
from urllib.parse import quote
import pandas as pd
from time import sleep
import undetected_chromedriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchWindowException, WebDriverException

RATE_LIMIT_COOLDOWN = 90   # seconds to wait after hitting a rate limit
MAX_RATE_LIMIT_RETRIES = 3

# Content Explorer URL. {keyword} is replaced with each keyword from keyword_list.xlsx (URL-encoded).
# Listicle search: title contains "best" OR "top" AND the keyword.
URL_TEMPLATE = (
    "https://app.ahrefs.com/content-explorer/results/en/score_desc?field=title"
    "&topic=(%22best%22%20OR%20%22top%22)%20AND%20{keyword}"
    "&one-article-per-domain=1&exclude-homepages=1&exclude-subdomains=1&filter-broken-pages=all&publish_type=all"
    "&filter-words=500-&published=365&first_seen=all&filter-subdomain-orgtraffic=1000-"
    "&filter-domain_rating=40-90&chart-range=last_6_months&page=1&results-per-page=50"
    "&report-languages=1%2CtotalTraffic%2Cdesc&report-websites=1%2CtotalTraffic%2Cdesc"
    "&report-authors=1%2CtotalTraffic%2Cdesc&safe-search=1"
)

# Previous guest-post URL, kept for reference:
# "https://app.ahrefs.com/content-explorer/results/en/score_desc?field=content&topic=(title%3A%22write%20for%20us%22%20OR%20url%3A%22write-for-us%22%20OR%20url%3A%22guest-post%22%20OR%20content%3A%22guest%20post%20guidelines%22)%20AND%20content%3A%22{keyword}%22&one-article-per-domain=1&filter-broken-pages=live&publish_type=all&filter-words=600-&published=2020-01-01T08%3A00%3A00Z%2C2023-09-14T08%3A00%3A00Z&first_seen=2019-12-31T20%3A00%3A00Z%2C2022-04-27T20%3A00%3A00Z&filter-subdomain-orgtraffic=2000-&filter-domain_rating=40-89&chart-range=last_6_months&page=1&results-per-page=50&report-languages=1%2CtotalTraffic%2Cdesc&report-websites=1%2CtotalTraffic%2Cdesc&report-authors=1%2CtotalTraffic%2Cdesc&safe-search=1"



def get_chrome_major_version():
    """Read the installed Chrome major version from the Windows registry (None if not found)."""
    try:
        import winreg
    except ImportError:
        return None
    keys = [
        (winreg.HKEY_CURRENT_USER, r"Software\Google\Chrome\BLBeacon", "version"),
        (winreg.HKEY_LOCAL_MACHINE, r"Software\Google\Chrome\BLBeacon", "version"),
        (winreg.HKEY_LOCAL_MACHINE, r"Software\WOW6432Node\Google\Chrome\BLBeacon", "version"),
        (winreg.HKEY_LOCAL_MACHINE, r"Software\Microsoft\Windows\CurrentVersion\Uninstall\Google Chrome", "DisplayVersion"),
        (winreg.HKEY_LOCAL_MACHINE, r"Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Google Chrome", "DisplayVersion"),
    ]
    for hive, path, name in keys:
        try:
            with winreg.OpenKey(hive, path) as k:
                version = winreg.QueryValueEx(k, name)[0]
                return int(str(version).split(".")[0])
        except (OSError, ValueError):
            continue
    return None



def ensure_window(driver):
    """Switch to an open Chrome tab. Newer Chrome versions sometimes close the first tab right after launch."""
    try:
        handles = driver.window_handles
    except WebDriverException:
        return False
    if not handles:
        return False
    driver.switch_to.window(handles[-1])
    return True


def create_driver(chrome_version, attempts=3):
    for attempt in range(1, attempts + 1):
        driver = undetected_chromedriver.Chrome(version_main=chrome_version)
        sleep(3)  # give Chrome time to settle its startup tabs
        if ensure_window(driver):
            return driver
        print(f"Chrome opened without a usable window (attempt {attempt}/{attempts}), restarting...")
        try:
            driver.quit()
        except Exception:
            pass
        sleep(3)
    raise RuntimeError("Could not start Chrome with an open window.")


def safe_get(driver, url, attempts=3):
    """driver.get() that recovers when the current tab was closed underneath it."""
    for attempt in range(1, attempts + 1):
        try:
            ensure_window(driver)
            driver.get(url)
            return
        except NoSuchWindowException:
            print(f"Tab was closed while loading (attempt {attempt}/{attempts}), switching tab...")
            sleep(2)
            if not ensure_window(driver):
                try:
                    driver.switch_to.new_window('tab')
                except WebDriverException:
                    pass
    raise RuntimeError(f"Could not load page after {attempts} attempts: {url}")


def is_rate_limited(driver):
    return "Rate limit exceeded" in driver.page_source or '"code":10029' in driver.page_source


def load_page(driver, web_page, keyword):
    for attempt in range(1, MAX_RATE_LIMIT_RETRIES + 1):
        safe_get(driver, web_page)
        sleep(3)
        if not is_rate_limited(driver):
            return True
        print(f"Rate limited on '{keyword}' (attempt {attempt}/{MAX_RATE_LIMIT_RETRIES}), "
              f"cooling down for {RATE_LIMIT_COOLDOWN}s...")
        sleep(RATE_LIMIT_COOLDOWN)
    print(f"Still rate limited after {MAX_RATE_LIMIT_RETRIES} attempts, skipping '{keyword}'.")
    return False


def get_latest_download():
    home = os.path.expanduser("~")
    downloadspath=os.path.join(home, "Downloads")
    list_of_files = glob.glob(os.path.join(downloadspath, "*.csv")) # * means all if need specific format then *.csv
    latest_file = max(list_of_files, key=os.path.getctime)
    return latest_file



def login(driver):
    safe_get(driver, "https://app.ahrefs.com/user/login")
    driver.maximize_window()
    sleep(10)

    # f = open("login.txt", "r")
    # then, we will send the keys for username_1
    # Find the username and password fields
    username_field = driver.find_element("name", 'email')
    password_field = driver.find_element('name', 'password')
    os.system("pause")
    # Fill the username and password fields
    # username_field.send_keys(f.readline())
    # password_field.send_keys(f.readline())
    # # Find the login button
    # login_button = driver.find_element(By.XPATH, "//button[@type='submit']")
    # login_button.click()


def rename_csv(keyword):
    try:
        file = get_latest_download()
        shutil.move(file, os.path.join(path_to_downloads, keyword + '.csv'))
    except Exception as err:
        print("Excepted renaming for keyword: ", keyword)
        print("Reason is: ", err)


if __name__ == '__main__':
    #################
    # Make dir "Temp" in current working directory
    parent_dir = os.getcwd()
    # set path and path to file
    path_to_temp = os.path.join(parent_dir, "Temp")
    path_to_downloads = os.path.join(parent_dir, "Downloads")

    try:
        shutil.rmtree(path_to_temp)
        shutil.rmtree(path_to_downloads)
    except:
        print("Temp does not exists")
    os.mkdir(path_to_temp)
    os.mkdir(path_to_downloads)

    chrome_version = get_chrome_major_version()
    print(f"Detected Chrome version: {chrome_version or 'unknown (letting undetected_chromedriver decide)'}")
    driver = create_driver(chrome_version)

    keyw = pd.read_excel('keyword_list.xlsx')
    key_list = keyw['list'].tolist()
    ###############
    login(driver)
    for i in range(0, len(key_list)):
        try:
            raw_keyword = str(key_list[i]).strip()
            if not raw_keyword or raw_keyword.lower() == 'nan':
                continue
            # file-safe name for the exported CSV
            keyword = re.sub(r'[\\/:*?"<>|]+', '-', re.sub(r'\s+', '_', raw_keyword))
            web_page = URL_TEMPLATE.format(keyword=quote(raw_keyword, safe=''))
            sleep(random.uniform(5, 10))
            if not load_page(driver, web_page, keyword):
                continue
            try:
                export_btn = WebDriverWait(driver, 30).until(EC.presence_of_element_located(
                    (By.XPATH,
                     "//button[normalize-space()='Export']")))
                export_btn.click()
                sleep(5)
                try:
                    driver.find_element(By.XPATH, "//span[contains(text(), 'One page per domain')]").click()
                    sleep(1)
                except:
                    try:
                        driver.find_element(By.XPATH, "//span[contains(text(), 'All')]").click()
                        sleep(1)
                    except:
                        pass
                driver.find_elements(By.XPATH, "//button[normalize-space()='Export']")[1].click()
                sleep(10)
            except:
                pass
            rename_csv(keyword)
        except:
            continue
    driver.quit()

