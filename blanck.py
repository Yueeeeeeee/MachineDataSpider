from bs4 import BeautifulSoup
import re
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options

url = "https://www.doosanmachinetools.com/en/main/index.do"

optionChrome = Options()
optionChrome.add_argument('--headless')
optionChrome.add_argument('--disable-gpu')
optionChrome.add_argument('disable-plugins')
optionChrome.add_argument('disable-extensions')

driverChrome = webdriver.Chrome(options=optionChrome)
driverChrome.get(url)
time.sleep(2)
htmlResult = driverChrome.page_source
driverChrome.quit()


soupMachine = BeautifulSoup(htmlResult, 'html5lib')

soup = soupMachine.find_all('div', {"class": "forDep"})

MachineURL = str(re.findall(r'href=".*"', str(soup))).replace('href="', "").replace('"', "").split(',')

urlList = []
urlIterator = iter(MachineURL)
Format = "https://www.doosanmachinetools.com/"
counter = 0
for i in urlIterator:

    if counter == len(MachineURL)-1:
        url = Format + str(i)[3:][:-2]
        urlList.append(url)
        #print(url)

    else:
        url = Format + str(i)[3:][:-1]
        urlList.append(url)
        counter = counter + 1
        #print(url)

print(urlList)