from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.select import Select
from openpyxl import Workbook, load_workbook
import setting
import os
import argparse
import time
from pynput.keyboard import Key, Controller


def browser_init():
    # WARNING: AGAR BISA SET LOKASI DOWNLOAD PDF, MAKA NAMA PROFILE HARUS "DEFAULT" 
    cud = setting.CHROME_USER_DATA
    cp = setting.CHROME_PROFILE
    options = webdriver.ChromeOptions()
    options.add_argument("user-data-dir={}".format(cud))
    options.add_argument("profile-directory={}".format(cp))
    options.add_argument('--no-sandbox')
    options.add_argument("--log-level=3")
    options.add_argument("--window-size=800,600")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    profile = {"download.default_directory": setting.CHROME_DOWNLOAD_PATH + os.sep, 
                "download.extensions_to_open": "applications/pdf",
                "download.prompt_for_download": False,
                'profile.default_content_setting_values.automatic_downloads': 1,
                "download.directory_upgrade": True,
                "plugins.always_open_pdf_externally": True                   
                }
    options.add_experimental_option("prefs", profile)

    return webdriver.Chrome(service=Service(executable_path=os.path.join(os.getcwd(), "chromedriver", "chromedriver.exe")), options=options)

def parse(year, box):
    url = 'https://arsip-sda.pusair-pu.go.id/login'
    driver = browser_init()
    driver.get(url)
    driver.maximize_window()
    try:
        driver.find_element(By.CSS_SELECTOR, "input[name='login']").clear()
        driver.find_element(By.CSS_SELECTOR, "input[name='password']").clear()    
        driver.find_element(By.CSS_SELECTOR, "input[name='login']").send_keys(setting.USER)
        driver.find_element(By.CSS_SELECTOR, "input[name='password']").send_keys(setting.PASSWORD)    
        driver.find_element(By.CSS_SELECTOR, "button[type='submit']").click()
    except:
        pass
    url = 'https://arsip-sda.pusair-pu.go.id/admin/archive/{}'.format(year)
    driver.get(url)
    
    trs = driver.find_element(By.CSS_SELECTOR, "table[id='dt-box-year']").find_elements(By.CSS_SELECTOR, "tr")
    for idx, tr in enumerate(trs):
        if idx == 0:
            continue
        if tr.find_element(By.CSS_SELECTOR, "div[class='media'] h6").text == box:
            # tr.find_element(By.CSS_SELECTOR, "a[class='btn']")
            link = tr.find_elements(By.CSS_SELECTOR, "a")[1].get_attribute("href")
            break
    
    driver.get(link)
    # breakpoint()
    Select(driver.find_element(By.CSS_SELECTOR, "select[name='dt-box-year_length']")).select_by_visible_text('50')
    time.sleep(1)
    trs = driver.find_element(By.CSS_SELECTOR, "table[id='dt-box-year']").find_elements(By.CSS_SELECTOR, "tr")
    items = []
    for idx, tr in enumerate(trs):
        if idx == 0:
            continue
        # time.sleep(1)
        filename = "-".join([year,box, tr.find_elements(By.CSS_SELECTOR, "td")[1].text, tr.find_elements(By.CSS_SELECTOR, "td")[2].text])+".png"
        link = tr.find_elements(By.CSS_SELECTOR, "a")[0].get_attribute("href") + "?ct=png"
        isempty = False
        if len(tr.find_elements(By.CSS_SELECTOR, "td[class='text-center'] i[class='fas fa-check text-success']")) == 0:
            isempty = True
        
        items.append((link, filename, isempty))
    for item in items:
        if os.path.exists(setting.PNG_LOCATION + os.path.sep + item[1]):
            if item[2]:
                driver.get(item[0])
                WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.CSS_SELECTOR, "a[id='uploader_browse']")))
                driver.find_element(By.CSS_SELECTOR, "a[id='uploader_browse']").click()
                time.sleep(3)          
                keyboard = Controller()
                keyboard.type(setting.PNG_LOCATION + os.path.sep + item[1])
                keyboard.press(Key.enter)
                keyboard.release(Key.enter)
                time.sleep(1)
                WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "a[id='uploader_start']")))
                time.sleep(1)
                driver.find_element(By.CSS_SELECTOR, "a[id='uploader_start']").click()
                # breakpoint()

                WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.CSS_SELECTOR, "a[id='uploader_stop']")))
                time.sleep(5)



def main():
    parser = argparse.ArgumentParser(description="PNG BOT")
    parser.add_argument('-y', '--year', type=str,help="Year")
    parser.add_argument('-b', '--box', type=str,help="Box")

    args = parser.parse_args()
    if args.year == None or args.box == None:
        print('use: python filepngbot.py -y <year> -b <box>')
        exit()
    parse(args.year, args.box)
    input("End Process...")

if __name__ == '__main__':
    main()
