# -*- coding: utf-8 -*-
"""
Created on Thu Mar 29 15:04:02 2018

@author: rohit
"""

import requests
import urllib
from bs4 import BeautifulSoup
import json
import re
import math
import sys
import pprint
import collections


html = requests.get('http://lczero.org/matches')
l0 =BeautifulSoup(html.text,'lxml')
fields =['match', 'candidateid', 'bestnet', 'result', 'score', 'elo', 'pass', 'time']
t0= l0.tbody.find_all('tr')
matches = collections.OrderedDict()
for t in t0:
    td = t.find_all('td')
    data ={}
    for i in range(1,len(td)):
        data[fields[i]]=td[i].text
    matches[td[0].text]= data

a=max(list(map(lambda x:float(x),list(matches.keys()))))

print('last match',matches[str(round(a))])

total_elo = sum(float(matches[match]['elo']) for match in matches if matches[match]['result']=='true')

print('total elo is:', total_elo)
