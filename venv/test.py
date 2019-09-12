import re
import lxml
from bs4 import BeautifulSoup
from urllib.request import urlopen

def writeData(curl):
    print(curl)

url = 'https://www.doosanmachinetools.com/en/product/series/D221_69/view.do'
html = urlopen(url).read().decode('utf-8')
soup = BeautifulSoup(html, 'lxml')

curl = soup.find('div', attrs={"class": "productList"})
i = 1

while curl is not None:
    writeData(curl)
    print(i)
    i += 1
    curl = curl.find_next('div', attrs={"class": "productList"})
