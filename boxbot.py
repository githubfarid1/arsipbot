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
    
    for boxno in range(start, end+1):
        driver.get("https://arsip-sda.pusair-pu.go.id/admin/master/box")
        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.CSS_SELECTOR, "input[name='name_box']")))
        driver.find_element(By.CSS_SELECTOR, "input[name='name_box']").send_keys(boxno)
        driver.find_element(By.CSS_SELECTOR, "input[name='year_box']").send_keys(s.YEAR)
        driver.find_elements(By.CSS_SELECTOR, "span[class='select2-selection__rendered']")[0].click()
        driver.find_element(By.CSS_SELECTOR, "input[class='select2-search__field']").send_keys(s.RAK)
        time.sleep(1)
        driver.find_element(By.CSS_SELECTOR, "li[class='select2-results__option select2-results__option--highlighted']").click()

        submit = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button[type='submit']")))
        # breakpoint()
        try:
            submit.click()
        except:
            submit.click()
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
