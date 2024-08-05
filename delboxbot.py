from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.select import Select
from openpyxl import Workbook, load_workbook
import setting as s
import os
import argparse
import time

def browser_init():
    # WARNING: AGAR BISA SET LOKASI DOWNLOAD PDF, MAKA NAMA PROFILE HARUS "DEFAULT" 
    cud = s.CHROME_USER_DATA
    cp = s.CHROME_PROFILE
    options = webdriver.ChromeOptions()
    options.add_argument("user-data-dir={}".format(cud))
    options.add_argument("profile-directory={}".format(cp))
    options.add_argument('--no-sandbox')
    options.add_argument("--log-level=3")
    options.add_argument("--window-size=800,600")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    profile = {"download.default_directory": s.CHROME_DOWNLOAD_PATH + os.sep, 
                "download.extensions_to_open": "applications/pdf",
                "download.prompt_for_download": False,
                'profile.default_content_setting_values.automatic_downloads': 1,
                "download.directory_upgrade": True,
                "plugins.always_open_pdf_externally": True                   
                }
    options.add_experimental_option("prefs", profile)

    return webdriver.Chrome(service=Service(executable_path=os.path.join(os.getcwd(), "chromedriver", "chromedriver.exe")), options=options)
def parse(start, end):
    url = 'https://arsip-sda.pusair-pu.go.id/login'
    driver = browser_init()
    driver.get(url)
    driver.maximize_window()
    try:
        driver.find_element(By.CSS_SELECTOR, "input[name='login']").clear()
        driver.find_element(By.CSS_SELECTOR, "input[name='password']").clear()    
        driver.find_element(By.CSS_SELECTOR, "input[name='login']").send_keys(s.USER)
        driver.find_element(By.CSS_SELECTOR, "input[name='password']").send_keys(s.PASSWORD)    
        driver.find_element(By.CSS_SELECTOR, "button[type='submit']").click()
    except:
        pass
    # breakpoint()
    for boxno in range(start, end+1):
        driver.get(f"https://arsip-sda.pusair-pu.go.id/admin/archive/{s.YEAR}")
        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.CSS_SELECTOR, "table#dt-box-year")))
        Select(driver.find_element(By.CSS_SELECTOR, "select[name='dt-box-year_length']")).select_by_visible_text('100')
        time.sleep(1)
        trs = driver.find_element(By.CSS_SELECTOR, "table[id='dt-box-year']").find_elements(By.CSS_SELECTOR, "tr")
        for idx, tr in enumerate(trs):
            if idx == 0:
                continue
            if tr.find_element(By.CSS_SELECTOR, "div[class='media'] h6").text == str(boxno):
                link = tr.find_elements(By.CSS_SELECTOR, "a")[1].get_attribute("href")
                break
        while True:
            driver.get(link)
            trs = driver.find_element(By.CSS_SELECTOR, "table[id='dt-box-year']").find_elements(By.CSS_SELECTOR, "tr")
            if trs[1].find_elements(By.CSS_SELECTOR, "td.dataTables_empty"):
                break
            trs[1].find_element(By.CSS_SELECTOR, "button[title='Hapus Arsip']").click()
            time.sleep(1)
            driver.find_element(By.CSS_SELECTOR, "div.swal2-actions").click()
            time.sleep(1)

def main():
    parser = argparse.ArgumentParser(description="BOX BOT")
    parser.add_argument('-s', '--start', type=str,help="Start number")
    parser.add_argument('-e', '--end', type=str,help="End number")

    args = parser.parse_args()
    if args.start == None or args.end == None:
        print('use: python boxbot.py -s <start_index> -e <end_index>')
        exit()
    parse(int(args.start), int(args.end))
    input("End Process...")

if __name__ == '__main__':
    main()
