# -*- coding: utf-8 -*-
"""
Created on Sun Apr  8 11:08:28 2018

@author: rohit
"""

# -*- coding: utf-8 -*-
"""
Created on Sun Apr  1 23:28:18 2018
make chart from bse pv data

@author: rohit
"""
import matplotlib.pyplot as plt
from datetime import datetime
import csv

def makechart(scrip='532712'):
    path = 'C:\\Users\\rohit\\Downloads\\'+ scrip+ '.csv'
    pvdata =[]
    with open(path) as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            pvdata.append(row)        
    pvdata = list(reversed(pvdata))
    Price = [date['Close Price'] for date in pvdata]
    Dates = [date['Date'] for date in pvdata]
    Volume = [date['No.of Shares'] for date in pvdata]
    Volume = [int(x) for x in Volume]
    Price = [float(x) for x in Price]
    Dates = [datetime.strptime(x,'%d-%B-%Y') for x in Dates]
    plt.figure(1)
    plt.subplot(211)
    plt.plot(Dates,Price)
    plt.title('Price')
    plt.subplot(212)
    plt.bar(Dates,Volume, width=0.1)
    plt.title('Volume')
    plt.show()
