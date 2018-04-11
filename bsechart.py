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
    fig, ax1 = plt.subplots()
    ax1.plot(Dates, Price, 'b-' , label = 'Price')
    ax1.set_xlabel('Dates')
    ax1.set_ylabel('Price', color='b')
    ax1.tick_params('y', colors='b')
    ax2 = ax1.twinx()
    ax2.bar(Dates, Volume, width =0.5 , label = 'Volume')
    ax2.set_ylabel('Volume', color='r')
    ax2.tick_params('y', colors='r')
    fig.tight_layout()
    ax1.legend(loc ='upper right')
    ax2.legend(loc ='upper left')
    plt.show()
