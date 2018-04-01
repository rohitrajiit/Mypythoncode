import requests
import urllib
from bs4 import BeautifulSoup
import json
import re
import math
import sys

try:
    r = requests.get('http://tests.stockfishchess.org/tests?success_only=1')
except:
    print ("request failed")
    sys.exit()

try:
    soup =BeautifulSoup(r.text,'lxml')
except:
    print ("no valid html")
    sys.exit()


tag = soup.td.parent.find_all('td')
a = {}
a['date'] = tag[0].string
a['username'] =tag[1].string
a['description'] =tag[6].string
a['timecontrol']= tag[5].string
a['results'] = tag[4].pre.string 
d= a['results']
e= a['results']
f= a['results']
winpattern = re.compile('W: ([0-9]+)')
aa = winpattern.search(a['results'])
wins = aa.groups()[0]

drawpattern = re.compile('D: ([0-9]+)')
b = drawpattern.search(a['results'])
draws = b.groups()[0]

losspattern = re.compile('L: ([0-9]+)')
c = losspattern.search(a['results'])
losses = c.groups()[0]

score = float(wins) + float(draws)/2;

total = float(wins) + float(draws) + float(losses);

percentage = (score /  float(total));

eloperf = -400 * math.log(1 / percentage - 1) / math.log(10)
try:
 with open('workfile.py') as data_file:    
    data = json.load(data_file)
except:
 data =0
 
if a!=data:
 print('new succeful test:',a['description'])
 print(a['timecontrol'],a['results'])
 print('elo performance', eloperf, 'elo')
 with open('workfile.py', 'w') as outfile:
    json.dump(a, outfile)
else:
 print('no new test')
 print('last test:',a['description'])
 print(a['timecontrol'],a['results'])
 print('elo performance', eloperf, 'elo')
