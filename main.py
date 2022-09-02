''' 
    Web Scraping Shopify stores and ranks from myip.ms
    @author ysfzkn
'''

from selenium import webdriver
from bs4 import BeautifulSoup
import time
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
import random
from selenium.webdriver.common.by import By
import pandas as pd
import os
from selenium.webdriver.common.proxy import *
from selenium.webdriver import ActionChains
import winsound

htmlFiles = list()
excelFiles = list()

service = Service("C:\chromedriver.exe")

def get_free_proxies(driver):
    driver.get('https://sslproxies.org')

    table = driver.find_element(By.TAG_NAME, 'table')
    thead = table.find_element(By.TAG_NAME, 'thead').find_elements(By.TAG_NAME, 'th')
    tbody = table.find_element(By.TAG_NAME, 'tbody').find_elements(By.TAG_NAME, 'tr')

    headers = []
    for th in thead:
        headers.append(th.text.strip())

    proxies = []
    for tr in tbody:
        proxy_data = {}
        tds = tr.find_elements(By.TAG_NAME, 'td')
        for i in range(len(headers)):
            proxy_data[headers[i]] = tds[i].text.strip()
        proxies.append(proxy_data)

    return proxies

def sleep(num):

    r =random.randint(num,num+3)
    time.sleep(r)

def getDataFromPageSource(pageNumber, content):

    urlList = list()
    rankList = list()

    soup = BeautifulSoup(content, 'lxml')

    i = 3
    while(i <= 99):
        for siteUrl in soup.select("#sites_tbl > tbody > tr:nth-child({}) > td.row_name > a".format(i)):
            url =  siteUrl.text
            for rankEl in soup.select("#sites_tbl > tbody > tr:nth-child({}) > td:nth-child(7) > span".format(i)):
                rank = rankEl.text
                print(url , rank)
            urlList.append(url)
            rankList.append(rank)

            i += 2

    df = pd.DataFrame([urlList , rankList]).T

    df.to_excel(f'page-{pageNumber}-data.xlsx', sheet_name='status')

    print(f"Page : {pageNumber} Successfully scraped ! \n")

def getDataWithSeleniumProxy(starterPage, PROXY):

    global service

    user_agent = "Mozilla/5.0 (Windows NT 6.0) AppleWebKit/536.5 (KHTML, like Gecko) Chrome/19.0.1084.36 Safari/536.5"

    options =  Options()
    
    print(PROXY)

    options.add_argument("user-agent="+user_agent)
    options.add_argument("--no-sandbox")
    options.add_argument("--window-size=1420,1080")
    options.add_argument("--disable-gpu")
    options.add_argument(f'--proxy-server={PROXY}')

    proxyItem = webdriver.Proxy()
    proxyItem.proxy_type = ProxyType.MANUAL
    proxyItem.auto_detect = False
    capabilities = webdriver.DesiredCapabilities.CHROME
    proxyItem.http_proxy = PROXY
    proxyItem.ssl_proxy = PROXY
    proxyItem.add_to_capabilities(capabilities)
    webdriver.DesiredCapabilities.CHROME['acceptSslCerts']=True

    driver = webdriver.Chrome(options=options,
                                service=service,desired_capabilities=capabilities)

    url = f"https://myip.ms/browse/sites/{starterPage}/ipID/23.227.38.0/ipIDii/23.227.38.255/sort/5/asc/1#sites_tbl_top"

    driver.get(url)

    # Click and Verification
    
    sleep(2)

    try:
        driver.find_element(by=By.XPATH, value="//*[@id='captcha_submit']").click()

        sleep(10)
    except:
        print("capthca gelmedi")

    driver.get(url)

    sleep(5)

    pageSource = driver.page_source
    getDataFromPageSource(starterPage, pageSource)

    global pageStarter
    pageStarter += 1

    # 1 2 3 4 5 6 7 8 9 10 11 12 13 14 15 16 17 18 19
    # 1 2 3 4 3 4 5 6 7 3  4  5  6  7  3  4  5  6  7
    for pageNumber in range(starterPage+1,starterPage+30):

        url = f"https://myip.ms/browse/sites/{pageNumber}/ipID/23.227.38.0/ipIDii/23.227.38.255/sort/5/asc/1#sites_tbl_top"

        driver.get(url)
        
        sleep(5)

        print("******** Page " + str(pageNumber) + "************")

        pageSource = driver.page_source

        getDataFromPageSource(pageNumber, pageSource)

        pageStarter += 1

def reset_router():

    browser = webdriver.Chrome(service=service)
    browser.get('http://192.168.1.1')

    sleep(3)

    username = browser.find_element(By.ID, 'AuthName') # Find ID of your own input element
    password = browser.find_element(By.ID, 'AuthPassword') # Find ID of your own input element

    username.send_keys("admin")    # Type your own username here
    password.send_keys("ozkan123") # Type your own password here

    sleep(2)

    browser.find_element(By.XPATH, '//*[@id="login"]/fieldset/ul/li[6]/input').click()

    sleep(5)

    a = ActionChains(browser)

    try:
        bakimButton = browser.find_element(by=By.LINK_TEXT, value="Bakım")
        a.move_to_element(bakimButton).perform()
        resetButton = browser.find_element(by=By.LINK_TEXT, value="Yeniden başlat")
        a.move_to_element(resetButton).click().perform()
        
        sleep(4)

        resetPage = browser.find_element(by=By.CSS_SELECTOR, value='#mainFrame')

        browser.switch_to.frame(resetPage)

        sleep(4)

        try: 
            browser.find_element(by=By.CSS_SELECTOR, 
            value='#contentPanel > div > div.data_frame > ul > li > div > ul > li.right_table > input[type=button]').click()
        except:
            bakimButton = browser.find_element(by=By.LINK_TEXT, value="Bakım")
            a.move_to_element(bakimButton).perform()
            resetButton = browser.find_element(by=By.LINK_TEXT, value="Yeniden başlat")
            a.move_to_element(resetButton).click().perform()
            
            sleep(4)

            resetPage = browser.find_element(by=By.CSS_SELECTOR, value='#mainFrame')

            browser.switch_to.frame(resetPage)

            sleep(4)
            browser.find_element(by=By.CSS_SELECTOR, 
            value='#contentPanel > div > div.data_frame > ul > li > div > ul > li.right_table > input[type=button]').click()

        print('Restarting...')
        sleep(5)
        print('Reset Successfull !')
        browser.close()

    except:
        duration = 300  # milliseconds 
        freq = 700  # Hz
        winsound.Beep(freq, duration)
        winsound.Beep(freq, duration)
        winsound.Beep(freq, duration)
        print('An Error Occured about Router Interface, Please Reset Manually!!!')

def getDataWithSelenium(starterPage):

    global service

    user_agent = "Mozilla/5.0 (Windows NT 6.0) AppleWebKit/536.5 (KHTML, like Gecko) Chrome/19.0.1084.36 Safari/536.5"

    options =  Options()

    options.add_argument("user-agent="+user_agent)
    options.add_argument("--no-sandbox")
    options.add_argument("--window-size=1420,1080")
    options.add_argument("--disable-gpu")

    driver = webdriver.Chrome(options=options,
                                service=service)

    url = f"https://myip.ms/browse/sites/{starterPage}/ipID/23.227.38.0/ipIDii/23.227.38.255/sort/5/asc/1#sites_tbl_top"

    driver.get(url)
    
    # Click and Verification
    
    sleep(3)

    try:
        driver.find_element(by=By.XPATH, value="//*[@id='captcha_submit']").click()

        sleep(4)
    except:
        print("capthca gelmedi")

    try:
        driver.find_element(by=By.CSS_SELECTOR, value="#tabs-1 > table > tbody > tr > td > div > p:nth-child(4)")
        print("Limit Exceed, Restarting...")
        driver.close()
        return

    except:
        print('Limit Control Passed!')

    global pageStarter # to control starter variable

    # Algorithm Logic of Click
    # 1 2 3 4 5 6 7 8 9 10 11 12 13 14 15 16 17 18 19
    # 1 2 3 4 3 4 5 6 7 3  4  5  6  7  3  4  5  6  7
    for pageNumber in range(starterPage ,starterPage+30):

        if pageNumber == 165:
            pageNumber == 7605

        url = f"https://myip.ms/browse/sites/{pageNumber}/ipID/23.227.38.0/ipIDii/23.227.38.255/sort/5/asc/1#sites_tbl_top"

        try: 
            driver.get(url)
            
            sleep(2)

            if pageNumber == starterPage or pageNumber == starterPage+1:
                try:
                    driver.find_element(by=By.XPATH, value="//*[@id='captcha_submit']").click()
                    sleep(3)
                except:
                    print("capthca gelmedi")

            try:
                driver.find_element(by=By.CSS_SELECTOR, value="#tabs-1 > table > tbody > tr > td > div > p:nth-child(4)")
                print("Limit Exceed, Restarting... Wait for 4 min to restart...")
                driver.close()
                return

            except:
                print('Limit Control Passed!')
                
            print("******** Page " + str(pageNumber) + "************")

            pageSource = driver.page_source

            getDataFromPageSource(pageNumber, pageSource)

            if pageNumber == starterPage+29:
                driver.close()

            pageStarter += 1

        except Exception as e:

            print('An error occured : ' + e)

            print('Trying to reset router...')

            reset_router()
            sleep(240)

            driver.get(url)
            
            sleep(3)

            print("******** Page " + str(pageNumber) + "************")

            pageSource = driver.page_source

            getDataFromPageSource(pageNumber, pageSource)

            if pageNumber == starterPage+29:
                driver.close()

            pageStarter += 1
            
def getAllHtmlFilesInDir():

    global htmlFiles

    directory = 'C:/Users/asus/Desktop/code/python/extract data from excel/'

    for root, dirnames, filenames in os.walk(directory):
        for filename in filenames:
            if filename.endswith('.html'):
                fname = os.path.join(root, filename)
                print('Filename: {}'.format(fname))
                htmlFiles.append(fname)

    return htmlFiles

def getDataFromHtmlFiles(htmlFiles):

    for file in htmlFiles:

        HTMLFileToBeOpened = open(file, "r")

        contents = HTMLFileToBeOpened.read()

        soup = BeautifulSoup(contents, 'lxml')

        i = 3
        while(i <= 99):
            for siteUrl in soup.select("#sites_tbl > tbody > tr:nth-child({}) > td.row_name > a".format(i)):
                if("com" in siteUrl.text):
                    for rank in soup.select("#sites_tbl > tbody > tr:nth-child({}) > td:nth-child(7) > span".format(i)):
                        print(siteUrl.text)
                        print(rank.text)
                i += 
                
        print(file + " successfully scraped !")

def getAllExcelFilesInDir():

    global excelFiles

    directory = 'C:/Users/asus/Desktop/code/python/extract data from excel/'
    
    for root, dirnames, filenames in os.walk(directory):
        for filename in filenames:
            if filename.endswith('.xlsx'):
                fname = os.path.join(root, filename)
                print('Filename: {}'.format(fname))
                excelFiles.append(fname)

    return excelFiles

def mergeExcelFiles(file_list):
    
    excel_list = list()
    excel_merged = pd.DataFrame()

    for file in file_list:
        excel_list.append(pd.read_excel(file))
    
    excel_merged = pd.concat(excel_list, ignore_index = True)
    excel_merged.to_excel('shopify_stores_data.xlsx', index=False)

def runScraper():
    pageStarter = 1 # Change with wherever you want from

    while pageStarter < 7607:
            
        getDataWithSelenium(pageStarter)

        duration = 1000  # milliseconds 
        freq = 600  # Hz
        winsound.Beep(freq, duration)
        
        reset_router()
        sleep(250) # 4-5 minutes waiting for Router Reset


if __name__ == "__main__":

    runScraper()
    excelList = getAllExcelFilesInDir()
    mergeExcelFiles(excelList)