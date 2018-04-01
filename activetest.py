import requests
from bs4 import BeautifulSoup
import pprint

r = requests.get('http://tests.stockfishchess.org/tests')
soup =BeautifulSoup(r.text,'html5lib') #lxml
div = soup.find("div", class_='span11')
ndiv = div.find('div',id = 'machines')
active = ndiv.find_next_sibling('table')
tests = active.find_all('tr')
activetest = {}
for a in tests:
    name = a.find_all('td')[-1].text
    llr =  a.pre.text
    activetest[name] = llr
pprint.pprint(activetest)


